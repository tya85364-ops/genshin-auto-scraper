import tkinter as tk
from tkinter import ttk, messagebox
import requests
import webbrowser

API_URL = "https://www.8591.com.tw/v3/mall/search"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'application/json, text/plain, */*'
}

# 遊戲分類 ID 對照表 (依據爬蟲設定的亞服/繁中服)
GAMES = {
    "崩壞：星穹鐵道 (亞服)": ("53396", "53397"),
    "原神 (亞服)": ("34169", "34170"),
    "鳴潮 (繁中服)": ("44693", "53160")
}

class RadarApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🎯 8591 倒賣操盤雷達 - 本機直連工具")
        self.geometry("1000x600")
        self.configure(bg="#1E1E1E")
        
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TLabel", background="#1E1E1E", foreground="#FFFFFF", font=("微軟正黑體", 10))
        style.configure("TButton", font=("微軟正黑體", 10, "bold"), background="#007ACC", foreground="#FFFFFF")
        style.configure("Treeview", background="#2D2D2D", foreground="#FFFFFF", fieldbackground="#2D2D2D", rowheight=30)
        style.map("Treeview", background=[("selected", "#007ACC")])
        
        # --- 頂部控制面板 ---
        control_frame = tk.Frame(self, bg="#1E1E1E")
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(control_frame, text="遊戲/伺服器:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.game_combo = ttk.Combobox(control_frame, values=list(GAMES.keys()), width=25, state="readonly")
        self.game_combo.current(0)
        self.game_combo.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(control_frame, text="關鍵字:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.keyword_entry = ttk.Entry(control_frame, width=20)
        self.keyword_entry.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(control_frame, text="最低價:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.min_price = ttk.Entry(control_frame, width=10)
        self.min_price.grid(row=0, column=5, padx=5, pady=5)
        
        ttk.Label(control_frame, text="最高價:").grid(row=0, column=6, padx=5, pady=5, sticky="w")
        self.max_price = ttk.Entry(control_frame, width=10)
        self.max_price.grid(row=0, column=7, padx=5, pady=5)
        
        self.search_btn = ttk.Button(control_frame, text="開始掃描", command=self.do_search)
        self.search_btn.grid(row=0, column=8, padx=15, pady=5)
        
        # --- 表格 ---
        cols = ("ID", "標題", "價格", "賣家分數")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("ID", text="商品編號")
        self.tree.column("ID", width=80, anchor="center")
        self.tree.heading("標題", text="標題 (雙擊可開啟賣場)")
        self.tree.column("標題", width=600, anchor="w")
        self.tree.heading("價格", text="價格 ($)")
        self.tree.column("價格", width=100, anchor="center")
        self.tree.heading("賣家分數", text="賣家評價")
        self.tree.column("賣家分數", width=100, anchor="center")
        
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # --- 狀態列 ---
        self.status_var = tk.StringVar(value="就緒")
        self.status_bar = ttk.Label(self, textvariable=self.status_var)
        self.status_bar.pack(fill=tk.X, padx=10, pady=5)

    def do_search(self):
        game_name = self.game_combo.get()
        game_id, server_id = GAMES[game_name]
        kw = self.keyword_entry.get().strip()
        p_min = self.min_price.get().strip() or "0"
        p_max = self.min_price.get().strip() or "99999"
        
        params = {
            "game_id": game_id,
            "server_id": server_id,
            "type": "1", # 帳號
            "firstRow": 0,
            "isLimitExt": 0,
            "keyword": kw
        }
        
        self.status_var.set("掃描中...請稍候")
        self.update_idletasks()
        
        try:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            r = requests.get(API_URL, params=params, headers=HEADERS, timeout=10, verify=False)
            data = r.json()
            if data["msg"] == "success":
                records = data["data"]["list"]
                
                # 價格過濾
                try: p_min_f, p_max_f = float(p_min), float(p_max or 999999)
                except: p_min_f, p_max_f = 0, 999999
                
                # 清除舊表格
                for item in self.tree.get_children():
                    self.tree.delete(item)
                    
                count = 0
                for rec in records:
                    price = float(rec.get("price", 0))
                    if p_min_f <= price <= p_max_f:
                        self.tree.insert("", "end", values=(
                            rec.get("goods_sn"),
                            rec.get("title", ""),
                            f"${price:,.0f}",
                            f"⭐ {rec.get('credit', '無')}"
                        ))
                        count += 1
                self.status_var.set(f"掃描完成！符合條件的獵物：{count} 筆")
            else:
                self.status_var.set(f"API 回應錯誤：{data.get('msg')}")
        except Exception as e:
            messagebox.showerror("錯誤", f"無法連接至 8591:\n{e}")
            self.status_var.set("就緒")

    def on_double_click(self, event):
        item = self.tree.selection()
        if not item: return
        item_values = self.tree.item(item[0], "values")
        goods_sn = item_values[0]
        url = f"https://www.8591.com.tw/v3/mall/detail/{goods_sn}"
        webbrowser.open(url)

if __name__ == "__main__":
    app = RadarApp()
    app.mainloop()
