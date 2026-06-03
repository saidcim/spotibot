"""
Spotify AI Playlist Bot
- Her 3 günde bir dinleme geçmişini analiz eder, AI ile playlist günceller.
- Flask API ile Vercel dashboard'a endpoint sağlar.
- State Railway environment variable'da saklanır (redeploy'da kaybolmaz).
"""

import os
import json
import time
import logging
import datetime
import threading
import openpyxl
import schedule
import spotipy
import requests
from spotipy.oauth2 import SpotifyOAuth
from groq import Groq
from flask import Flask, jsonify, request
from flask_cors import CORS

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
log = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=False)

# ── Config ─────────────────────────────────────────────────────────────────
SPOTIFY_CLIENT_ID     = os.environ["SPOTIFY_CLIENT_ID"]
SPOTIFY_CLIENT_SECRET = os.environ["SPOTIFY_CLIENT_SECRET"]
SPOTIFY_REDIRECT_URI  = os.environ.get("SPOTIFY_REDIRECT_URI", "http://127.0.0.1:5000/callback")
SPOTIFY_REFRESH_TOKEN = os.environ["SPOTIFY_REFRESH_TOKEN"]
GROQ_API_KEY          = os.environ["GROQ_API_KEY"]
API_SECRET            = os.environ.get("API_SECRET", "said")
PLAYLIST_NAME         = os.environ.get("PLAYLIST_NAME", "🤖 AI Daily Mix")
PLAYLIST_SIZE         = int(os.environ.get("PLAYLIST_SIZE", "40"))

# Railway API — state kalıcılığı için
RAILWAY_API_TOKEN     = os.environ.get("RAILWAY_API_TOKEN", "")
RAILWAY_PROJECT_ID    = os.environ.get("RAILWAY_PROJECT_ID", "")
RAILWAY_SERVICE_ID    = os.environ.get("RAILWAY_SERVICE_ID", "")
RAILWAY_ENVIRONMENT_ID= os.environ.get("RAILWAY_ENVIRONMENT_ID", "")

HISTORY_FILE = "playlist_history.xlsx"
STATE_VAR    = "BOT_STATE"   # Railway'de saklanacak variable adı

is_running = False

# ── State: Railway variable'da sakla ──────────────────────────────────────

def _railway_headers():
    return {"Authorization": f"Bearer {RAILWAY_API_TOKEN}", "Content-Type": "application/json"}

def _can_use_railway():
    return all([RAILWAY_API_TOKEN, RAILWAY_PROJECT_ID, RAILWAY_SERVICE_ID, RAILWAY_ENVIRONMENT_ID])

def load_state() -> dict:
    default = {
        "playlist_id": os.environ.get("PLAYLIST_ID", ""),
        "last_update": None,
        "cycle": 0,
        "feedback_history": [],
        "mood_history": [],
        "ai_notes": ""
    }

    # Önce Railway variable'dan oku
    if _can_use_railway():
        try:
            query = """
            query Variables($projectId: String!, $serviceId: String!, $environmentId: String!) {
              variables(projectId: $projectId, serviceId: $serviceId, environmentId: $environmentId)
            }"""
            resp = requests.post(
                "https://backboard.railway.app/graphql/v2",
                headers=_railway_headers(),
                json={"query": query, "variables": {
                    "projectId": RAILWAY_PROJECT_ID,
                    "serviceId": RAILWAY_SERVICE_ID,
                    "environmentId": RAILWAY_ENVIRONMENT_ID
                }},
                timeout=10
            )
            data = resp.json()
            variables = data.get("data", {}).get("variables", {})
            if STATE_VAR in variables:
                state = json.loads(variables[STATE_VAR])
                log.info(f"State Railway'den yüklendi (döngü #{state.get('cycle', 0)})")
                return state
        except Exception as e:
            log.warning(f"Railway state okunamadı, dosyaya bakılıyor: {e}")

    # Fallback: lokal dosya
    if os.path.exists("bot_state.json"):
        with open("bot_state.json") as f:
            return json.load(f)

    return default


def save_state(state: dict):
    # Lokal dosyaya yaz (hızlı fallback)
    with open("bot_state.json", "w") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)

    # Railway variable'a yaz (kalıcı)
    if _can_use_railway():
        try:
            mutation = """
            mutation UpsertVariables($input: VariableCollectionUpsertInput!) {
              variableCollectionUpsert(input: $input)
            }"""
            requests.post(
                "https://backboard.railway.app/graphql/v2",
                headers=_railway_headers(),
                json={"query": mutation, "variables": {"input": {
                    "projectId": RAILWAY_PROJECT_ID,
                    "serviceId": RAILWAY_SERVICE_ID,
                    "environmentId": RAILWAY_ENVIRONMENT_ID,
                    "variables": {STATE_VAR: json.dumps(state, ensure_ascii=False)}
                }}},
                timeout=10
            )
            log.info("State Railway'e kaydedildi.")
        except Exception as e:
            log.warning(f"Railway state kaydedilemedi: {e}")


# ── Excel Geçmişi ──────────────────────────────────────────────────────────

def save_to_excel(tracks: list, cycle: int, score=None):
    wb = openpyxl.load_workbook(HISTORY_FILE) if os.path.exists(HISTORY_FILE) else openpyxl.Workbook()
    ws = wb.active
    ws.title = "Playlist History"
    if ws.max_row == 1 and ws.cell(1, 1).value is None:
        for col, h in enumerate(["Döngü","Tarih","Spotify ID","Şarkı","Sanatçı","Albüm","AI Skoru"], 1):
            ws.cell(1, col).value = h
    date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    for track in tracks:
        ws.append([cycle, date_str, track["id"], track["name"], track["artist"], track["album"], score or "—"])
    wb.save(HISTORY_FILE)
    log.info(f"Excel'e {len(tracks)} şarkı kaydedildi (döngü {cycle})")


# ── Spotify ────────────────────────────────────────────────────────────────

def get_spotify() -> spotipy.Spotify:
    auth = SpotifyOAuth(
        client_id=SPOTIFY_CLIENT_ID,
        client_secret=SPOTIFY_CLIENT_SECRET,
        redirect_uri=SPOTIFY_REDIRECT_URI,
        scope=(
            "user-read-recently-played user-top-read user-library-read "
            "playlist-modify-public playlist-modify-private playlist-read-private"
        ),
    )
    token_info = auth.refresh_access_token(SPOTIFY_REFRESH_TOKEN)
    return spotipy.Spotify(auth=token_info["access_token"])


def get_listening_data(sp: spotipy.Spotify) -> dict:
    data = {}
    recent = sp.current_user_recently_played(limit=50)
    recent_tracks = []
    cutoff = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=3)
    for item in recent["items"]:
        played_at = datetime.datetime.strptime(item["played_at"], "%Y-%m-%dT%H:%M:%S.%fZ").replace(tzinfo=datetime.timezone.utc)
        if played_at >= cutoff:
            t = item["track"]
            recent_tracks.append({"id": t["id"], "name": t["name"],
                "artist": t["artists"][0]["name"], "album": t["album"]["name"],
                "played_at": item["played_at"]})
    data["recent_tracks"] = recent_tracks

    top_short = sp.current_user_top_tracks(limit=50, time_range="short_term")
    data["top_short"] = [{"id": t["id"], "name": t["name"], "artist": t["artists"][0]["name"],
        "popularity": t["popularity"]} for t in top_short["items"]]

    top_medium = sp.current_user_top_tracks(limit=50, time_range="medium_term")
    data["top_medium"] = [{"id": t["id"], "name": t["name"], "artist": t["artists"][0]["name"],
        "popularity": t["popularity"]} for t in top_medium["items"]]

    top_artists = sp.current_user_top_artists(limit=20, time_range="short_term")
    data["top_artists"] = [{"name": a["name"], "genres": a["genres"]} for a in top_artists["items"]]

    saved = sp.current_user_saved_tracks(limit=50)
    data["saved_tracks"] = [{"id": item["track"]["id"], "name": item["track"]["name"],
        "artist": item["track"]["artists"][0]["name"]} for item in saved["items"]]

    return data


def get_playlist_play_counts(sp, playlist_track_ids, recent_tracks):
    counts = {tid: 0 for tid in playlist_track_ids}
    for t in recent_tracks:
        if t["id"] in counts:
            counts[t["id"]] += 1
    return counts


def find_or_create_playlist(sp: spotipy.Spotify, state: dict) -> str:
    user_id = sp.current_user()["id"]
    if state.get("playlist_id"):
        try:
            pl = sp.playlist(state["playlist_id"])
            log.info(f"Mevcut playlist: {pl['name']} ({pl['id']})")
            return state["playlist_id"]
        except Exception:
            log.warning("Playlist ID geçersiz, aranıyor...")

    matching = []
    offset = 0
    while True:
        results = sp.current_user_playlists(limit=50, offset=offset)
        for pl in results["items"]:
            if pl["name"] == PLAYLIST_NAME and pl["owner"]["id"] == user_id:
                matching.append(pl)
        if results["next"] is None:
            break
        offset += 50

    if matching:
        keeper = matching[0]
        for dup in matching[1:]:
            try:
                sp.current_user_unfollow_playlist(dup["id"])
                log.info(f"Kopya silindi: {dup['id']}")
            except Exception as e:
                log.warning(f"Kopya silinemedi: {e}")
        return keeper["id"]

    pl = sp.user_playlist_create(user=user_id, name=PLAYLIST_NAME, public=False,
        description="🤖 Her 3 günde güncellenen AI playlist")
    log.info(f"Yeni playlist: {pl['id']}")
    return pl["id"]


def update_playlist(sp, playlist_id, track_ids):
    sp.playlist_replace_items(playlist_id, [])
    for i in range(0, len(track_ids), 100):
        chunk = [f"spotify:track:{tid}" for tid in track_ids[i:i+100]]
        sp.playlist_add_items(playlist_id, chunk)
    log.info(f"Playlist güncellendi: {len(track_ids)} şarkı")


# ── AI ─────────────────────────────────────────────────────────────────────

def analyze_mood(listening_data, client):
    all_recent = listening_data.get("recent_tracks", [])
    prompt = f"""Sen bir müzik psikolojisi uzmanısın. Son 3 günlük dinleme verisine bakarak ruh halini analiz et.

## Son 3 Günde Dinlenen Şarkılar
{json.dumps(all_recent, ensure_ascii=False)}

## Kısa Dönem En Çok Dinlenenler
{json.dumps(listening_data.get('top_short', [])[:15], ensure_ascii=False)}

## Favori Sanatçılar
{json.dumps(listening_data.get('top_artists', [])[:8], ensure_ascii=False)}

SADECE JSON döndür:
{{"mood": "Ana ruh hali (Türkçe)", "mood_emoji": "emoji", "energy_level": "düşük/orta/yüksek",
"dominant_genres": ["tür1","tür2"], "top_artists_this_period": ["sanatçı1","sanatçı2","sanatçı3"],
"summary": "2-3 cümlelik Türkçe özet", "track_count": {len(all_recent)}}}"""

    try:
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile", max_tokens=600,
            messages=[{"role": "user", "content": prompt}])
        raw = resp.choices[0].message.content.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"): raw = raw[4:]
        return json.loads(raw.strip())
    except Exception as e:
        log.warning(f"Ruh hali analizi başarısız: {e}")
        return {"mood": "Bilinmiyor", "mood_emoji": "🎵", "energy_level": "orta",
                "dominant_genres": [], "top_artists_this_period": [],
                "summary": "Analiz yapılamadı.", "track_count": len(all_recent)}


MAX_CARRY_OVER = 5  # Eski playlistten yeni playliste geçebilecek max şarkı sayısı

def ai_analyze_and_build(listening_data, play_counts, state, candidate_ids, current_track_ids=None):
    client = Groq(api_key=GROQ_API_KEY)

    current_track_ids = current_track_ids or []

    # Eski playlistten en çok çalınan MAX_CARRY_OVER şarkıyı belirle
    if current_track_ids and play_counts:
        sorted_old = sorted(play_counts.items(), key=lambda x: x[1], reverse=True)
        carry_over_pool = [tid for tid, count in sorted_old if count > 0][:MAX_CARRY_OVER]
        banned_ids = [tid for tid in current_track_ids if tid not in carry_over_pool]
    else:
        carry_over_pool = []
        banned_ids = current_track_ids

    # Aday listesinden yasak şarkıları çıkar
    fresh_candidates = [tid for tid in candidate_ids if tid not in banned_ids]
    log.info(f"Aday: {len(fresh_candidates)} taze + {len(carry_over_pool)} taşınabilir eski şarkı")

    feedback_summary = ""
    if state["feedback_history"]:
        last = state["feedback_history"][-1]
        feedback_summary = (f"Önceki döngü ({last['cycle']}): skor {last['score']:.1f}/10, "
            f"ort. çalma {last['avg_plays']:.1f}. Notlar: {last.get('notes', '')}")

    mood_data = analyze_mood(listening_data, client)
    log.info(f"Ruh hali: {mood_data.get('mood')} {mood_data.get('mood_emoji')}")

    prompt = f"""Sen bir müzik küratörüsün. {PLAYLIST_SIZE} şarkılık playlist seç.

Döngü #{state['cycle'] + 1} | {feedback_summary}
Ruh hali: {json.dumps(mood_data, ensure_ascii=False)}
Son dinlenenler: {json.dumps(listening_data['recent_tracks'][:30], ensure_ascii=False)}
Top şarkılar: {json.dumps(listening_data['top_short'][:20], ensure_ascii=False)}
Çalma sayıları: {json.dumps(play_counts, ensure_ascii=False)}
Top sanatçılar: {json.dumps(listening_data['top_artists'][:10], ensure_ascii=False)}

## TAZE ADAY ŞARKILAR (öncelikli olarak bunlardan seç)
{json.dumps(fresh_candidates[:80], ensure_ascii=False)}

## TAŞINABİLİR ESKİ ŞARKILAR (geçen dönem çok çalındı, en fazla {MAX_CARRY_OVER} tane kullanabilirsin)
{json.dumps(carry_over_pool, ensure_ascii=False)}

## KESİNLİKLE YASAK (bu şarkıları ekleme)
{json.dumps(banned_ids[:50], ensure_ascii=False)}

Görev:
1. Taze aday şarkılardan en az {PLAYLIST_SIZE - MAX_CARRY_OVER} şarkı seç.
2. Taşınabilir eski şarkılardan en fazla {MAX_CARRY_OVER} tane ekleyebilirsin.
3. Yasak listesindeki şarkıları KESİNLİKLE ekleme.
4. Ruh haline uygun seç, çeşitlilik ekle, sıralamayı karıştır.

SADECE JSON:
{{"track_ids":["id1","id2"],"score":7.5,"analysis":"Türkçe analiz","notes":"sonraki döngü notu"}}"""

    resp = client.chat.completions.create(
        model="llama-3.3-70b-versatile", max_tokens=2000,
        messages=[{"role": "user", "content": prompt}])
    raw = resp.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"): raw = raw[4:]
    parsed = json.loads(raw.strip())

    track_ids = parsed.get("track_ids", [])[:PLAYLIST_SIZE]

    # Kod tarafında da zorla: yasak şarkıları çıkar, carry_over limitini uygula
    carry_over_used = [tid for tid in track_ids if tid in carry_over_pool]
    if len(carry_over_used) > MAX_CARRY_OVER:
        excess = set(carry_over_used[MAX_CARRY_OVER:])
        track_ids = [tid for tid in track_ids if tid not in excess]
        log.info(f"Carry-over limiti: {len(excess)} fazla şarkı çıkarıldı")
    track_ids = [tid for tid in track_ids if tid not in banned_ids]

    fresh_count = len([tid for tid in track_ids if tid not in current_track_ids])
    carry_count = len([tid for tid in track_ids if tid in carry_over_pool])
    log.info(f"AI tamamlandı. Skor: {parsed.get('score')}, Taze: {fresh_count}, Eski: {carry_count}, Toplam: {len(track_ids)}")
    return track_ids, parsed.get("notes",""), float(parsed.get("score",5.0)), parsed.get("analysis",""), mood_data


# ── Ana Döngü ─────────────────────────────────────────────────────────────

def run_cycle(manual=False):
    global is_running
    if is_running:
        log.warning("Zaten çalışıyor.")
        return {"status": "already_running"}
    is_running = True
    trigger = "Manuel" if manual else "Otomatik"
    log.info(f"═══════════ Döngü Başlıyor ({trigger}) ═══════════")

    try:
        state = load_state()
        sp = get_spotify()

        playlist_id = find_or_create_playlist(sp, state)
        state["playlist_id"] = playlist_id
        save_state(state)

        current_tracks = []
        if state.get("last_update"):
            try:
                items = sp.playlist_items(playlist_id, fields="items(track(id,name,artists,album))")
                for item in items["items"]:
                    t = item["track"]
                    if t and t.get("id"):
                        current_tracks.append({"id": t["id"], "name": t["name"],
                            "artist": t["artists"][0]["name"], "album": t["album"]["name"]})
            except Exception as e:
                log.warning(f"Playlist okunamadı: {e}")

        listening_data = get_listening_data(sp)

        play_counts = {}
        avg_plays = 0
        if current_tracks:
            current_ids = [t["id"] for t in current_tracks]
            play_counts = get_playlist_play_counts(sp, current_ids, listening_data["recent_tracks"])
            avg_plays = sum(play_counts.values()) / max(len(play_counts), 1)

        candidate_ids = list({t["id"] for t in (
            listening_data["top_short"] + listening_data["top_medium"] +
            listening_data["saved_tracks"] + listening_data["recent_tracks"]
        ) if t.get("id")})

        current_ids_for_ai = [t['id'] for t in current_tracks]
        new_track_ids, ai_notes, score, analysis, mood_data = ai_analyze_and_build(
            listening_data, play_counts, state, candidate_ids, current_ids_for_ai)

        if current_tracks:
            save_to_excel(current_tracks, state["cycle"], score)

        update_playlist(sp, playlist_id, new_track_ids)

        now = datetime.datetime.now(datetime.timezone.utc).isoformat()
        state["cycle"] += 1
        state["last_update"] = now
        state["ai_notes"] = ai_notes

        if "mood_history" not in state:
            state["mood_history"] = []
        state["mood_history"].append({"date": now, "cycle": state["cycle"],
            "trigger": trigger, **mood_data})
        state["mood_history"] = state["mood_history"][-30:]

        state["feedback_history"].append({"cycle": state["cycle"], "date": now,
            "score": score, "avg_plays": avg_plays, "analysis": analysis, "notes": ai_notes})
        state["feedback_history"] = state["feedback_history"][-10:]

        save_state(state)
        log.info(f"═══════════ Döngü #{state['cycle']} Tamamlandı ═══════════")
        return {"status": "ok", "cycle": state["cycle"], "tracks": len(new_track_ids), "mood": mood_data}

    except Exception as e:
        log.error(f"Döngü hatası: {e}", exc_info=True)
        return {"status": "error", "message": str(e)}
    finally:
        is_running = False


# ── Flask API ──────────────────────────────────────────────────────────────

def check_auth(req):
    secret = req.headers.get("X-API-Secret") or req.args.get("secret")
    return secret == API_SECRET

@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "bot": "Spotify AI Playlist Bot"})

@app.route("/status", methods=["GET"])
def status():
    if not check_auth(request):
        return jsonify({"error": "Unauthorized"}), 401
    state = load_state()
    return jsonify({
        "cycle": state.get("cycle", 0),
        "last_update": state.get("last_update"),
        "playlist_id": state.get("playlist_id"),
        "is_running": is_running,
        "feedback_history": state.get("feedback_history", []),
        "mood_history": state.get("mood_history", [])
    })

@app.route("/trigger", methods=["POST"])
def trigger():
    if not check_auth(request):
        return jsonify({"error": "Unauthorized"}), 401
    if is_running:
        return jsonify({"status": "already_running"}), 409
    t = threading.Thread(target=lambda: run_cycle(manual=True), daemon=True)
    t.start()
    return jsonify({"status": "started"})

@app.route("/mood-history", methods=["GET"])
def mood_history():
    if not check_auth(request):
        return jsonify({"error": "Unauthorized"}), 401
    state = load_state()
    return jsonify(state.get("mood_history", []))


# ── Başlat ─────────────────────────────────────────────────────────────────

def scheduler_loop():
    schedule.every(72).hours.do(run_cycle)
    while True:
        schedule.run_pending()
        time.sleep(60)

def main():
    log.info("Spotify AI Bot başlatıldı.")
    threading.Thread(target=run_cycle, daemon=True).start()
    threading.Thread(target=scheduler_loop, daemon=True).start()
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

if __name__ == "__main__":
    main()