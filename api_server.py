import os
import json
import base64
import subprocess
import sys
from datetime import datetime
from flask import Flask, request, jsonify
from flask_cors import CORS
from pymongo import MongoClient

app = Flask(__name__)
CORS(app)

MONGO_URI = os.environ.get("MONGODB_URI", "mongodb+srv://genshin:genshin123@cluster0.svtlvs0.mongodb.net/scraper_db?appName=Cluster0")

GCP_KEY_STATUS = "Not started"

# ─── 啟動時解碼 GCP Key ───────────────────────────────────────────────────────
def _setup_gcp_key():
    """將 GCP key 解碼成 gcp_key.json（支援三段拼接，避免 Railway 截斷）"""
    global GCP_KEY_STATUS
    # 優先用三段拼接（最穩定，每段 ~1000 字，Railway 不會截斷）
    p1 = os.environ.get("GCP_KEY_PART_1", "").strip()
    p2 = os.environ.get("GCP_KEY_PART_2", "").strip()
    p3 = os.environ.get("GCP_KEY_PART_3", "").strip()
    if p1:
        b64 = p1 + p2 + p3
    else:
        # fallback：舊版單段 B64
        b64 = os.environ.get("GCP_KEY_B64", "").strip()

    if not b64:
        GCP_KEY_STATUS = "WARNING: No GCP_KEY_PART_1 or GCP_KEY_B64"
        print("[API] 無 GCP_KEY_PART_1 也無 GCP_KEY_B64，略過 gcp_key.json 寫入")
        return

    # 補回 Railway 可能截掉 the base64 padding
    b64 += "=" * ((-len(b64)) % 4)
    key_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gcp_key.json")
    try:
        decoded = base64.b64decode(b64).decode("utf-8")
        
        # 修正私鑰中的真實換行符，避免 JSON 解析失敗
        import re
        def repl(match):
            val = match.group(1)
            val = val.replace("\r", "").replace("\n", "\\n")
            return f'"private_key": "{val}"'
        decoded = re.sub(r'"private_key"\s*:\s*"([^"]*)"', repl, decoded, flags=re.DOTALL)
        
        json.loads(decoded)  # 驗證 JSON 正確
        with open(key_path, "w", encoding="utf-8") as f:
            f.write(decoded)
        GCP_KEY_STATUS = f"SUCCESS: gcp_key.json written OK ({len(decoded)} chars)"
        print(f"[API] {GCP_KEY_STATUS}")
    except Exception as e:
        GCP_KEY_STATUS = f"ERROR: {e}"
        print(f"[API] ERROR writing gcp_key.json: {e}")

_setup_gcp_key()

def get_db():
    client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
    return client["scraper_db"]

# ─── Start background workers ───────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))

def start_workers():
    """Start scraper and Discord bot as background subprocesses."""
    procs = []
    
    # Create logs directory
    log_dir = os.path.join(BASE, "logs")
    os.makedirs(log_dir, exist_ok=True)
    
    discord_path = os.path.join(BASE, "discord_bot.py")
    if os.path.exists(discord_path):
        discord_log_path = os.path.join(log_dir, "discord_bot.log")
        with open(discord_log_path, "w", encoding="utf-8") as f:
            f.write(f"=== Discord Bot Started at {datetime.now()} ===\n")
        discord_log = open(discord_log_path, "a", encoding="utf-8")
        p = subprocess.Popen([sys.executable, discord_path],
                             stdout=discord_log, stderr=discord_log)
        procs.append(("discord_bot", p))
        print(f"[API] Discord bot started (pid={p.pid})")
    
    scraper_path = os.path.join(BASE, "genshin_scraper_original.py")
    if os.path.exists(scraper_path):
        scraper_log_path = os.path.join(log_dir, "scraper.log")
        with open(scraper_log_path, "w", encoding="utf-8") as f:
            f.write(f"=== Scraper Started at {datetime.now()} ===\n")
        scraper_log = open(scraper_log_path, "a", encoding="utf-8")
        p = subprocess.Popen([sys.executable, scraper_path],
                             stdout=scraper_log, stderr=scraper_log)
        procs.append(("scraper", p))
        print(f"[API] Scraper started (pid={p.pid})")
    
    return procs

# ─── API routes ─────────────────────────────────────────────────────────────
@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "alive"}), 200

@app.route('/api/targets', methods=['GET'])
def get_targets():
    try:
        db = get_db()
        targets = list(db["custom_targets"].find({}))
        # _id is the URL string, already serializable
        return jsonify({"status": "ok", "data": targets}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/targets', methods=['POST'])
def add_target():
    data = request.json
    if not data or 'url' not in data or 'target_price' not in data:
        return jsonify({"status": "error", "message": "Missing url or target_price"}), 400
    
    url = data['url']
    try:
        target_price = int(data['target_price'])
    except (ValueError, TypeError):
        return jsonify({"status": "error", "message": "target_price must be integer"}), 400
    
    title = data.get('title', 'Unknown Item')
    
    db = get_db()
    db["custom_targets"].update_one(
        {"_id": url},
        {"$set": {"target_price": target_price, "title": title, "alerted": False}},
        upsert=True
    )
    return jsonify({"status": "ok", "message": "Target saved"}), 200

@app.route('/api/targets/<path:url>', methods=['DELETE'])
def delete_target(url):
    db = get_db()
    db["custom_targets"].delete_one({"_id": url})
    return jsonify({"status": "ok"}), 200

@app.route('/api/logs/<name>', methods=['GET'])
def get_log(name):
    if name not in ["scraper", "discord_bot"]:
        return jsonify({"status": "error", "message": "Invalid log name"}), 400
    log_file = os.path.join(BASE, "logs", f"{name}.log")
    if not os.path.exists(log_file):
        return jsonify({"status": "ok", "content": "Log file empty or not found"}), 200
    try:
        with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
            content = "".join(lines[-200:])
        return jsonify({"status": "ok", "content": content}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/gcp_status', methods=['GET'])
def gcp_status():
    p1 = os.environ.get("GCP_KEY_PART_1", "")
    p2 = os.environ.get("GCP_KEY_PART_2", "")
    p3 = os.environ.get("GCP_KEY_PART_3", "")
    
    combined_b64 = p1.strip() + p2.strip() + p3.strip()
    padded_b64 = combined_b64 + "=" * ((-len(combined_b64)) % 4)
    
    status_info = {
        "status": GCP_KEY_STATUS,
        "p1_len": len(p1),
        "p2_len": len(p2),
        "p3_len": len(p3),
        "combined_b64_len": len(combined_b64),
        "padded_b64_len": len(padded_b64),
        "p1_preview": p1[:10] + "..." + p1[-10:] if p1 else "empty",
        "p2_preview": p2[:10] + "..." + p2[-10:] if p2 else "empty",
        "p3_preview": p3[:10] + "..." + p3[-10:] if p3 else "empty"
    }
    
    if combined_b64:
        try:
            decoded = base64.b64decode(padded_b64).decode("utf-8")
            status_info["raw_decoded_len"] = len(decoded)
            
            # 嘗試進行清洗
            import re
            def repl(match):
                val = match.group(1)
                val = val.replace("\r", "").replace("\n", "\\n")
                return f'"private_key": "{val}"'
            clean_decoded = re.sub(r'"private_key"\s*:\s*"([^"]*)"', repl, decoded, flags=re.DOTALL)
            
            status_info["decoded_len"] = len(clean_decoded)
            status_info["decoded_preview"] = clean_decoded[:30] + "..." + clean_decoded[-30:]
            json.loads(clean_decoded)
            status_info["json_valid"] = True
        except Exception as ex:
            status_info["json_valid"] = False
            status_info["json_error"] = str(ex)
            import re
            m = re.search(r"char (\d+)", str(ex))
            if m:
                pos = int(m.group(1))
                start = max(0, pos - 50)
                end = min(len(decoded), pos + 50)
                snippet = decoded[start:end]
                status_info["error_snippet"] = snippet
                status_info["error_snippet_hex"] = [hex(ord(c)) for c in snippet]
                if pos < len(decoded):
                    status_info["error_char"] = decoded[pos]
                    status_info["error_char_hex"] = hex(ord(decoded[pos]))
            
    return jsonify(status_info), 200

# ─── Entry point ────────────────────────────────────────────────────────────
# Called at module level so gunicorn --preload also triggers it
_workers_started = False
def _ensure_workers():
    global _workers_started
    if not _workers_started:
        start_workers()
        _workers_started = True

_ensure_workers()

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 31422))
    print(f"[API] Flask dev server on 0.0.0.0:{port}")
    app.run(host='0.0.0.0', port=port, threaded=True)
