#!/bin/sh

echo "🔄 啟動 Discord 機器人 (背景服務)..."
python discord_bot.py &

echo "🔄 啟動核心市場分析爬蟲 (背景服務)..."
python genshin_scraper_original.py &

echo "🔄 啟動 API 伺服器 (主服務 - Railway PORT 監聽)..."
exec python api_server.py
