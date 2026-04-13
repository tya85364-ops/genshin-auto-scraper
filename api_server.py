import os
import json
from flask import Flask, request, jsonify
from flask_cors import CORS
from pymongo import MongoClient

app = Flask(__name__)
# Enable CORS so the PWA GitHub Pages can hit this API
CORS(app)

# MongoDB Connection
MONGO_URI = os.environ.get("MONGODB_URI", "mongodb+srv://genshin:genshin123@cluster0.svtlvs0.mongodb.net/scraper_db?appName=Cluster0")

def get_db():
    client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
    try:
        db = client.get_default_database()
    except Exception:
        db = client["scraper_db"]
    return db

@app.route('/api/targets', methods=['GET'])
def get_targets():
    try:
        db = get_db()
        targets = list(db["custom_targets"].find({}))
        return jsonify({"status": "ok", "data": targets}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/targets', methods=['POST'])
def add_target():
    data = request.json
    if not data or 'url' not in data or 'target_price' not in data:
        return jsonify({"status": "error", "message": "Missing url or target_price"}), 400
    
    url = data['url']
    target_price = int(data['target_price'])
    title = data.get('title', 'Unknown Item')
    
    db = get_db()
    db["custom_targets"].update_one(
        {"_id": url},
        {"$set": {"target_price": target_price, "title": title, "alerted": False}},
        upsert=True
    )
    
    return jsonify({"status": "ok", "message": "Target saved successfully"}), 200

@app.route('/api/targets/<path:url>', methods=['DELETE'])
def delete_target(url):
    db = get_db()
    db["custom_targets"].delete_one({"_id": url})
    return jsonify({"status": "ok", "message": "Target deleted"}), 200

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "alive"}), 200

if __name__ == '__main__':
    # Railway passes the PORT via environment variable
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
