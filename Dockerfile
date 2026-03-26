# 使用 Microsoft 官方 Playwright Python 映像（已內建 Chromium binary）
FROM mcr.microsoft.com/playwright/python:v1.57.0-noble

WORKDIR /app

# 先安裝依賴（利用 Docker cache 層加速重建）
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 複製程式碼與 tier list JSON（靜態設定檔）
COPY genshin_scraper_original.py .
COPY genshin_tier_list.json .
COPY wutheringwaves_tier_list.json .
COPY hsr_tier_list.json .

# 確保 Python 輸出不 buffer（log 即時顯示）
ENV PYTHONUNBUFFERED=1

CMD ["python", "genshin_scraper_original.py"]
