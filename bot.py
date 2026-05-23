"""
Spotify AI Playlist Bot
Her 3 günde bir dinleme geçmişini analiz eder, AI ile playlist günceller.
"""

import os
import json
import time
import logging
import datetime
import openpyxl
import schedule
import spotipy
from spotipy.oauth2 import SpotifyOAuth
from anthropic import Anthropic

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("bot.log"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

# ── Ortam değişkenlerinden config ──────────────────────────────────────────
SPOTIFY_CLIENT_ID     = os.environ["SPOTIFY_CLIENT_ID"]
SPOTIFY_CLIENT_SECRET = os.environ["SPOTIFY_CLIENT_SECRET"]
SPOTIFY_REDIRECT_URI  = os.environ.get("SPOTIFY_REDIRECT_URI", "http://localhost:8888/callback")
SPOTIFY_REFRESH_TOKEN = os.environ["SPOTIFY_REFRESH_TOKEN"]  # ilk kurulumda doldurulacak
ANTHROPIC_API_KEY     = os.environ["ANTHROPIC_API_KEY"]
PLAYLIST_ID           = os.environ.get("PLAYLIST_ID", "")     # ilk çalıştırmada otomatik oluşur
PLAYLIST_NAME         = os.environ.get("PLAYLIST_NAME", "🤖 AI Daily Mix")
PLAYLIST_SIZE         = int(os.environ.get("PLAYLIST_SIZE", "40"))
HISTORY_FILE          = "playlist_history.xlsx"
STATE_FILE            = "bot_state.json"

# ──────────────────────────────────────────────────────────────────────────

def get_spotify() -> spotipy.Spotify:
    auth = SpotifyOAuth(
        client_id=SPOTIFY_CLIENT_ID,
        client_secret=SPOTIFY_CLIENT_SECRET,
        redirect_uri=SPOTIFY_REDIRECT_URI,
        scope=(
            "user-read-recently-played "
            "user-top-read "
            "user-library-read "
            "playlist-modify-public "
            "playlist-modify-private "
            "playlist-read-private"
        ),
    )
    token_info = auth.refresh_access_token(SPOTIFY_REFRESH_TOKEN)
    return spotipy.Spotify(auth=token_info["access_token"])


def load_state() -> dict:
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE) as f:
            return json.load(f)
    return {
        "playlist_id": PLAYLIST_ID,
        "last_update": None,
        "cycle": 0,
        "feedback_history": [],
        "ai_notes": ""
    }


def save_state(state: dict):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)


# ── Excel Geçmişi ──────────────────────────────────────────────────────────

def save_to_excel(tracks: list[dict], cycle: int, score: float | None = None):
    """Playlist şarkılarını Excel'e ekler."""
    wb = openpyxl.load_workbook(HISTORY_FILE) if os.path.exists(HISTORY_FILE) else openpyxl.Workbook()
    ws = wb.active
    ws.title = "Playlist History"

    if ws.max_row == 1 and ws.cell(1, 1).value is None:
        headers = ["Döngü", "Tarih", "Spotify ID", "Şarkı", "Sanatçı", "Albüm", "AI Skoru"]
        for col, h in enumerate(headers, 1):
            ws.cell(1, col).value = h

    date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    for track in tracks:
        ws.append([
            cycle,
            date_str,
            track["id"],
            track["name"],
            track["artist"],
            track["album"],
            score if score is not None else "—"
        ])

    wb.save(HISTORY_FILE)
    log.info(f"Excel'e {len(tracks)} şarkı kaydedildi (döngü {cycle})")


# ── Spotify Veri Toplama ───────────────────────────────────────────────────

def get_listening_data(sp: spotipy.Spotify) -> dict:
    """Son 3 gündeki dinleme verisi + top tracks + beğenilen şarkılar."""
    data = {}

    # Son dinlenenler (max 50)
    recent = sp.current_user_recently_played(limit=50)
    recent_tracks = []
    cutoff = datetime.datetime.utcnow() - datetime.timedelta(days=3)
    for item in recent["items"]:
        played_at = datetime.datetime.strptime(item["played_at"], "%Y-%m-%dT%H:%M:%S.%fZ")
        if played_at >= cutoff:
            t = item["track"]
            recent_tracks.append({
                "id": t["id"],
                "name": t["name"],
                "artist": t["artists"][0]["name"],
                "album": t["album"]["name"],
                "played_at": item["played_at"]
            })
    data["recent_tracks"] = recent_tracks

    # Kısa dönem top tracks
    top_short = sp.current_user_top_tracks(limit=50, time_range="short_term")
    data["top_short"] = [
        {"id": t["id"], "name": t["name"], "artist": t["artists"][0]["name"],
         "popularity": t["popularity"]}
        for t in top_short["items"]
    ]

    # Orta dönem top tracks
    top_medium = sp.current_user_top_tracks(limit=50, time_range="medium_term")
    data["top_medium"] = [
        {"id": t["id"], "name": t["name"], "artist": t["artists"][0]["name"],
         "popularity": t["popularity"]}
        for t in top_medium["items"]
    ]

    # Top artists
    top_artists = sp.current_user_top_artists(limit=20, time_range="short_term")
    data["top_artists"] = [
        {"name": a["name"], "genres": a["genres"]}
        for a in top_artists["items"]
    ]

    # Beğenilen şarkılar (son 50)
    saved = sp.current_user_saved_tracks(limit=50)
    data["saved_tracks"] = [
        {"id": item["track"]["id"], "name": item["track"]["name"],
         "artist": item["track"]["artists"][0]["name"]}
        for item in saved["items"]
    ]

    return data


def get_playlist_play_counts(sp: spotipy.Spotify, playlist_track_ids: list[str],
                              recent_tracks: list[dict]) -> dict:
    """Mevcut playlistteki şarkıların son 3 günde kaç kez çalındığını hesaplar."""
    counts = {tid: 0 for tid in playlist_track_ids}
    for t in recent_tracks:
        if t["id"] in counts:
            counts[t["id"]] += 1
    return counts


# ── AI Analizi ────────────────────────────────────────────────────────────

def ai_analyze_and_build(
    listening_data: dict,
    play_counts: dict,
    state: dict,
    candidate_ids: list[str]
) -> tuple[list[str], str]:
    """
    Claude'a dinleme verisini gönderir, playlist için şarkı ID listesi + notlar alır.
    """
    client = Anthropic(api_key=ANTHROPIC_API_KEY)

    # Önceki döngü feedback özeti
    feedback_summary = ""
    if state["feedback_history"]:
        last = state["feedback_history"][-1]
        feedback_summary = (
            f"Önceki döngü ({last['cycle']}): "
            f"playlist skoru {last['score']:.1f}/10, "
            f"ortalama çalma sayısı {last['avg_plays']:.1f}. "
            f"Notlar: {last.get('notes', '')}"
        )

    prompt = f"""Sen bir müzik küratörü ve dinleme alışkanlıkları analistinin.
Kullanıcının Spotify verilerini analiz edip en iyi {PLAYLIST_SIZE} şarkılık playlist önereceksin.

## Kullanıcı Profili
Döngü #{state['cycle'] + 1}
Önceki döngü notu: {state.get('ai_notes', 'İlk döngü.')}
{feedback_summary}

## Son 3 Günlük Dinleme ({len(listening_data['recent_tracks'])} şarkı)
{json.dumps(listening_data['recent_tracks'][:30], ensure_ascii=False, indent=2)}

## Kısa Dönem Favoriler (Top {len(listening_data['top_short'])} Şarkı)
{json.dumps(listening_data['top_short'][:20], ensure_ascii=False, indent=2)}

## Mevcut Playlist Çalma Sayıları (son 3 gün)
{json.dumps(play_counts, ensure_ascii=False)}

## Top Sanatçılar
{json.dumps(listening_data['top_artists'][:10], ensure_ascii=False, indent=2)}

## Aday Şarkılar (seçebilirsin)
{json.dumps(candidate_ids[:80], ensure_ascii=False)}

## Görev
1. Mevcut playlist skoru: Eğer play_counts verisinde şarkılar çok çalındıysa (>2) playlist başarılıydı, az çalındıysa değil.
2. Kullanıcı hangi şarkı türlerinden/sanatçılardan sıkılıyor mu değerlendir.
3. Aday listesinden VE top listelerden {PLAYLIST_SIZE} şarkı ID seç.
4. Seçtiğin şarkıları karıştır, sıkılmamak için çeşitlilik ekle.

Yanıtını SADECE JSON olarak ver, başka hiçbir şey yazma:
{{
  "track_ids": ["spotify_id_1", "spotify_id_2", ...],
  "score": <önceki playlist puanı 0-10>,
  "analysis": "<kısa Türkçe analiz, ne değişti, neden bu seçimler>",
  "notes": "<bir sonraki döngü için hatırlatma notları>"
}}"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text.strip()
    # JSON fence temizle
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    parsed = json.loads(raw.strip())

    track_ids = parsed.get("track_ids", [])[:PLAYLIST_SIZE]
    notes = parsed.get("notes", "")
    score = parsed.get("score", 5.0)
    analysis = parsed.get("analysis", "")

    log.info(f"AI analizi tamamlandı. Skor: {score}, Seçilen: {len(track_ids)} şarkı")
    log.info(f"AI Analiz: {analysis}")

    return track_ids, notes, float(score), analysis


# ── Playlist Güncelleme ────────────────────────────────────────────────────

def ensure_playlist(sp: spotipy.Spotify, state: dict) -> str:
    """Playlist yoksa oluştur, varsa ID'yi döndür."""
    if state.get("playlist_id"):
        try:
            sp.playlist(state["playlist_id"])
            return state["playlist_id"]
        except Exception:
            log.warning("Playlist bulunamadı, yeni oluşturuluyor.")

    user = sp.current_user()
    pl = sp.user_playlist_create(
        user=user["id"],
        name=PLAYLIST_NAME,
        public=False,
        description="🤖 Her 3 günde güncellenen AI playlist"
    )
    log.info(f"Yeni playlist oluşturuldu: {pl['id']}")
    return pl["id"]


def update_playlist(sp: spotipy.Spotify, playlist_id: str, track_ids: list[str]):
    """Mevcut playlistin şarkılarını tamamen değiştirir."""
    # Mevcut şarkıları temizle
    sp.playlist_replace_items(playlist_id, [])
    # 100'lük parçalar halinde ekle (Spotify limiti)
    for i in range(0, len(track_ids), 100):
        chunk = [f"spotify:track:{tid}" for tid in track_ids[i:i+100]]
        sp.playlist_add_items(playlist_id, chunk)
    log.info(f"Playlist güncellendi: {len(track_ids)} şarkı eklendi")


# ── Ana Döngü ─────────────────────────────────────────────────────────────

def run_cycle():
    log.info("═══════════ Yeni Döngü Başlıyor ═══════════")
    state = load_state()
    sp = get_spotify()

    # 1. Mevcut playlistin şarkılarını kaydet (önceki döngü için Excel)
    playlist_id = ensure_playlist(sp, state)
    state["playlist_id"] = playlist_id

    current_tracks = []
    if state.get("last_update"):
        try:
            items = sp.playlist_items(playlist_id, fields="items(track(id,name,artists,album))")
            for item in items["items"]:
                t = item["track"]
                if t and t.get("id"):
                    current_tracks.append({
                        "id": t["id"],
                        "name": t["name"],
                        "artist": t["artists"][0]["name"],
                        "album": t["album"]["name"]
                    })
        except Exception as e:
            log.warning(f"Mevcut playlist okunamadı: {e}")

    # 2. Dinleme verisi topla
    listening_data = get_listening_data(sp)

    # 3. Mevcut playlistin çalınma sayılarını hesapla
    play_counts = {}
    avg_plays = 0
    if current_tracks:
        current_ids = [t["id"] for t in current_tracks]
        play_counts = get_playlist_play_counts(sp, current_ids, listening_data["recent_tracks"])
        avg_plays = sum(play_counts.values()) / max(len(play_counts), 1)
        log.info(f"Mevcut playlist ort. çalma: {avg_plays:.1f}")

    # 4. Aday şarkı havuzu (top + saved + recent)
    candidate_ids = list({
        t["id"] for t in (
            listening_data["top_short"] +
            listening_data["top_medium"] +
            listening_data["saved_tracks"] +
            listening_data["recent_tracks"]
        ) if t.get("id")
    })

    # 5. AI ile analiz ve yeni şarkı listesi
    new_track_ids, ai_notes, score, analysis = ai_analyze_and_build(
        listening_data, play_counts, state, candidate_ids
    )

    # 6. Excel'e kaydet (önceki playlist)
    if current_tracks:
        save_to_excel(current_tracks, state["cycle"], score)

    # 7. Playlistı güncelle
    update_playlist(sp, playlist_id, new_track_ids)

    # 8. State güncelle
    state["cycle"] += 1
    state["last_update"] = datetime.datetime.utcnow().isoformat()
    state["ai_notes"] = ai_notes
    state["feedback_history"].append({
        "cycle": state["cycle"],
        "date": datetime.datetime.utcnow().isoformat(),
        "score": score,
        "avg_plays": avg_plays,
        "analysis": analysis,
        "notes": ai_notes
    })
    # Son 10 döngü feedback'i tut
    state["feedback_history"] = state["feedback_history"][-10:]
    save_state(state)

    log.info(f"═══════════ Döngü #{state['cycle']} Tamamlandı ═══════════")
    log.info(f"Sonraki güncelleme: 3 gün sonra")


def main():
    log.info("Spotify AI Bot başlatıldı.")
    
    # İlk çalıştırmada hemen bir döngü yap
    run_cycle()

    # Her 3 günde bir çalıştır (72 saat)
    schedule.every(72).hours.do(run_cycle)

    while True:
        schedule.run_pending()
        time.sleep(60)


if __name__ == "__main__":
    main()
