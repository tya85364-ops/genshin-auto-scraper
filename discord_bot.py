import os
import requests
import discord
from discord.ext import commands
from discord import app_commands
from dotenv import load_dotenv
from typing import Optional

# 載入環境變數
load_dotenv()
TOKEN = os.getenv("DISCORD_BOT_TOKEN")

# 遊戲清單與 8591 對應 ID (依據爬蟲設定的亞服/繁中服)
GAMES = {
    "崩壞：星穹鐵道 (亞服)": ("53396", "53397"),
    "原神 (亞服)": ("34169", "34170"),
    "鳴潮 (繁中服)": ("44693", "53160")
}

API_URL = "https://www.8591.com.tw/v3/mall/search"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'application/json, text/plain, */*'
}

class RadarBot(discord.Client):
    def __init__(self):
        super().__init__(intents=discord.Intents.default())
        self.tree = app_commands.CommandTree(self)

    async def setup_hook(self):
        # 啟動時註冊 slash commands 到全域
        await self.tree.sync()
        print(f"✅ Discord 機器人上線：{self.user}")

bot = RadarBot()

@bot.tree.command(name="search", description="8591 市場快速盤查")
@app_commands.describe(
    game="選擇遊戲名稱",
    min_price="最低價",
    max_price="最高價",
    keyword="必須包含的關鍵字"
)
@app_commands.choices(game=[
    app_commands.Choice(name=k, value=k) for k in GAMES.keys()
])
async def search_8591(
    interaction: discord.Interaction,
    game: app_commands.Choice[str],
    min_price: Optional[int] = None,
    max_price: Optional[int] = None,
    keyword: Optional[str] = None
):
    # 立即回應，避免超時 (8591 API 有時較慢)
    await interaction.response.defer(thinking=True)
    
    game_id, server_id = GAMES[game.value]
    p_min = min_price if min_price is not None else 0
    p_max = max_price if max_price is not None else 999999
    kw    = keyword.strip() if keyword else ""
    params = {
        "game_id": game_id,
        "server_id": server_id,
        "type": "1",
        "firstRow": 0,
        "isLimitExt": 0,
        "keyword": kw
    }
    
    try:
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        r = requests.get(API_URL, params=params, headers=HEADERS, timeout=10, verify=False)
        data = r.json()
        
        if data["msg"] != "success":
            await interaction.followup.send(f"❌ 8591 查詢失敗: {data['msg']}")
            return
            
        records = data["data"]["list"]
        matched = []
        for rec in records:
            price = float(rec.get("price", 0))
            title_txt = rec.get("title", "")
            if p_min <= price <= p_max:
                if not kw or kw in title_txt:
                    matched.append(rec)
                
        if not matched:
            cond = f"${p_min}~${p_max}" + (f" 關鍵字:{kw}" if kw else "")
            await interaction.followup.send(f"🔍 找不到符合條件（{cond}）的【{game.name}】帳號哦。")
            return
            
        cond_str = f"${p_min}~${p_max}" + (f" | 關鍵字：`{kw}`" if kw else "")
        embed = discord.Embed(
            title=f"🎯 8591 雷達偵測結果 ({len(matched)} 筆)",
            description=f"**條件**：{game.name} | {cond_str}\n（僅顯示前 5 筆）",
            color=0x00FF00
        )
        
        for i, rec in enumerate(matched[:5]):
            sn = rec.get("goods_sn")
            url = f"https://www.8591.com.tw/v3/mall/detail/{sn}"
            title = rec.get("title", "無標題")
            price = rec.get("price", 0)
            score = rec.get("credit", "無")
            
            embed.add_field(
                name=f"💰 ${price:,} | 賣家分數: {score}",
                value=f"[{title}]({url})",
                inline=False
            )
            
        await interaction.followup.send(embed=embed)
        
    except Exception as e:
        await interaction.followup.send(f"⚠️ 查詢發生錯誤: {e}")

if __name__ == "__main__":
    if not TOKEN:
        print("❌ 未設定 DISCORD_BOT_TOKEN")
    else:
        bot.run(TOKEN)
