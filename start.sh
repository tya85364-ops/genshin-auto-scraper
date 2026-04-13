#!/bin/sh

echo "🔄 啟動 Discord 機器人 (背景服務)..."
python discord_bot.py &

echo "🔄 啟動 API 伺服器 (背景服務)..."
python api_server.py &

echo "🔄 啟動核心市場分析爬蟲 (主服務)..."
python genshin_scraper_original.py
