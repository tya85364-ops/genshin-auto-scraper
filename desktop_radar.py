import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import webbrowser
from dotenv import load_dotenv

# 嘗試載入 PyMongo
try:
    from pymongo import MongoClient
    HAS_PYMONGO = True
except ImportError:
    HAS_PYMONGO = False

load_dotenv()

# 遊戲檔案名稱對應
GAMES = {
    "崩壞：星穹鐵道": "sr_market_stats.json",
    "原神": "gs_market_stats.json",
    "鳴潮": "ww_market_stats.json"
}

class RadarApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🎯 8591 倒賣庫存查盤雷達 - 本機/雲端資料庫版")
        self.geometry("1000x600")
        self.configure(bg="#1E1E1E")
        
        # 建立 MongoDB 連線
        self.mongo_client = None
        self.db = None
        uri = os.getenv("MONGODB_URI")
        if HAS_PYMONGO and uri:
            try:
                self.mongo_client = MongoClient(uri, serverSelectionTimeoutMS=3000)
                self.mongo_client.admin.command('ping')
                self.db = self.mongo_client["genshin_scraper"]
                print("✅ 成功連線至 MongoDB")
            except Exception as e:
                print(f"⚠️ MongoDB 連線失敗，將退回使用本機 JSON: {e}")
                self.db = None
        
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TLabel", background="#1E1E1E", foreground="#FFFFFF", font=("微軟正黑體", 10))
        style.configure("TButton", font=("微軟正黑體", 10, "bold"), background="#007ACC", foreground="#FFFFFF")
        style.configure("Treeview", background="#2D2D2D", foreground="#FFFFFF", fieldbackground="#2D2D2D", rowheight=30)
        style.map("Treeview", background=[("selected", "#007ACC")])
        
        # --- 頂部控制面板 ---
        control_frame = tk.Frame(self, bg="#1E1E1E")
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(control_frame, text="遊戲:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.game_combo = ttk.Combobox(control_frame, values=list(GAMES.keys()), width=15, state="readonly")
        self.game_combo.current(0)
        self.game_combo.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(control_frame, text="關鍵字:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.keyword_entry = ttk.Entry(control_frame, width=15)
        self.keyword_entry.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(control_frame, text="CP最少:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.min_cp = ttk.Entry(control_frame, width=8)
        self.min_cp.insert(0, "0")
        self.min_cp.grid(row=0, column=5, padx=5, pady=5)
        
        ttk.Label(control_frame, text="預算($):").grid(row=0, column=6, padx=5, pady=5, sticky="w")
        self.min_price = ttk.Entry(control_frame, width=8)
        self.min_price.grid(row=0, column=7, padx=5, pady=5)
        
        ttk.Label(control_frame, text="~").grid(row=0, column=8, padx=2, pady=5, sticky="w")
        self.max_price = ttk.Entry(control_frame, width=8)
        self.max_price.grid(row=0, column=9, padx=5, pady=5)
        
        self.search_btn = ttk.Button(control_frame, text="資料庫檢索", command=self.do_search)
        self.search_btn.grid(row=0, column=10, padx=15, pady=5)
        
        # --- 表格 ---
        cols = ("日期", "標題", "價格", "純角CP", "金卡數", "賣家")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("日期", text="紀錄日期")
        self.tree.column("日期", width=90, anchor="center")
        self.tree.heading("標題", text="標題 (雙擊前往賣場)")
        self.tree.column("標題", width=450, anchor="w")
        self.tree.heading("價格", text="價格 ($)")
        self.tree.column("價格", width=80, anchor="center")
        self.tree.heading("純角CP", text="純角CP")
        self.tree.column("純角CP", width=60, anchor="center")
        self.tree.heading("金卡數", text="金數")
        self.tree.column("金卡數", width=60, anchor="center")
        self.tree.heading("賣家", text="賣家資訊")
        self.tree.column("賣家", width=120, anchor="center")
        
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # 隱藏的 URL 記錄器
        self.row_urls = {}
        
        self.status_var = tk.StringVar(value=f"就緒，資料來源：{'MongoDB雲端' if self.db is not None else '本機 JSON'}")
        self.status_bar = ttk.Label(self, textvariable=self.status_var)
        self.status_bar.pack(fill=tk.X, padx=10, pady=5)

    def do_search(self):
        game_name = self.game_combo.get()
        filename = GAMES[game_name]
        kw = self.keyword_entry.get().strip()
        
        try: p_min = float(self.min_price.get().strip() or str(0))
        except: p_min = 0
        
        try: p_max = float(self.max_price.get().strip() or str(999999))
        except: p_max = 999999
        
        try: min_cp_val = float(self.min_cp.get().strip() or str(0))
        except: min_cp_val = 0

        self.status_var.set("檢索中...請稍候")
        self.update_idletasks()
        
        records = []
        if self.db is not None:
            # 雲端模式
            doc = self.db["market_stats"].find_one({"_id": filename.replace('.json', '')})
            if doc and "records" in doc:
                records = doc["records"]
        else:
            # 本機模式
            if os.path.exists(filename):
                try:
                    with open(filename, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        records = data.get("records", [])
                except Exception as e:
                    self.status_var.set(f"讀取 {filename} 失敗：{e}")
                    return
            else:
                self.status_var.set(f"找不到資料庫檔案 {filename}")
                return
        
        # 清除舊表格
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.row_urls.clear()
        
        count = 0
        records.reverse()  # 從最新開始顯示
        
        for rec in records:
            title = rec.get("title", "")
            price = rec.get("price", 0)
            cp1 = rec.get("cp1", 0) or 0
            
            # 過濾條件
            if not (p_min <= price <= p_max):
                continue
            if min_cp_val > 0 and (cp1 < min_cp_val):
                continue
            if kw and (kw.lower() not in title.lower()):
                continue
                
            item_id = self.tree.insert("", "end", values=(
                rec.get("date", "未知"),
                title,
                f"${price:,.0f}",
                f"{cp1:.2f}",
                f"{rec.get('gold_char', 0)} / {rec.get('gold_weap', 0)}",
                rec.get("seller_str", "未知")
            ))
            self.row_urls[item_id] = rec.get("url", "")
            count += 1

        self.status_var.set(f"檢索完成！符合條件的獵物：{count} 筆 (來源: {'MongoDB雲端' if self.db is not None else '本機 JSON'})")

    def on_double_click(self, event):
        item = self.tree.selection()
        if not item: return
        url = self.row_urls.get(item[0])
        if url:
            webbrowser.open(url)

if __name__ == "__main__":
    app = RadarApp()
    app.mainloop()
