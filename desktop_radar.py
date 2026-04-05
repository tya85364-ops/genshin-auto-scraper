import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser
import os
import threading
from dotenv import load_dotenv

load_dotenv()

# ---- Google Sheets 設定 ----
SPREADSHEET_ID = "1SOt-2DwJVEcEgvuvQfAvW6ue6WcrnvywxPbKIJFEcYI"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GCP_KEY_FILE = os.path.join(BASE_DIR, "gcp_key.json")

# 工作表名稱對照 (在架 / 成交)
GAME_SHEETS = {
    "崩鐵 (在架)":  ("崩鐵",            "in_progress"),
    "崩鐵 (成交)":  ("崩鐵-成交紀錄",   "completed"),
    "原神 (在架)":  ("原神",            "in_progress"),
    "原神 (成交)":  ("原神-成交紀錄",   "completed"),
    "鳴潮 (在架)":  ("鳴潮",            "in_progress"),
    "鳴潮 (成交)":  ("鳴潮-成交紀錄",   "completed"),
}

# ---- 欄位定義 (依照實際工作表 header) ----
# 在架: 發現時間(0) 上架時間(1) 標題(2) 價格(3) 金角(4) 金武/專(5) 純角CP(6) ... 賣家ID(12) 連結(13)
# 成交: 成交發現日(0) 上架時間(1) 售出天數(2) 標題(3) 價格(4) 金角(5) ... 純角CP(9) ... 賣家ID(12) 連結(13)
COLS = {
    "in_progress": {"date": 0, "title": 2, "price": 3, "gold": 4, "cp1": 6, "seller": 12, "url": 13},
    "completed":   {"date": 0, "title": 3, "price": 4, "gold": 5, "cp1": 9, "seller": 12, "url": 13},
}


def get_gspread_client():
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_file(GCP_KEY_FILE, scopes=scopes)
    return gspread.authorize(creds)

def fetch_sheet_data(sheet_name, status_callback):
    """從試算表抓取指定工作表的所有資料列 (跳過標題列)"""
    try:
        status_callback("連線 Google Sheets 中...")
        gc = get_gspread_client()
        sh = gc.open_by_key(SPREADSHEET_ID)

        # 嘗試找到符合名稱的工作表
        ws = None
        for s in sh.worksheets():
            if sheet_name in s.title:
                ws = s
                break

        if ws is None:
            # 找不到就列出所有工作表供偵錯
            names = [s.title for s in sh.worksheets()]
            status_callback(f"找不到工作表 '{sheet_name}'。現有: {names}")
            return []

        status_callback(f"讀取工作表「{ws.title}」中...")
        rows = ws.get_all_values()
        return rows[1:]  # 跳過標題列
    except Exception as e:
        status_callback(f"讀取失敗：{e}")
        return []


class RadarApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🎯 8591 看盤雷達 - Google Sheets 直連版")
        self.geometry("1150x650")
        self.configure(bg="#1E1E1E")
        self._raw_rows = []   # 目前拉下來的原始資料
        self._row_urls = {}   # treeview item_id -> url

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TLabel",   background="#1E1E1E", foreground="#FFFFFF", font=("微軟正黑體", 10))
        style.configure("TButton",  font=("微軟正黑體", 10, "bold"), background="#007ACC", foreground="#FFFFFF")
        style.configure("TEntry",   fieldbackground="#2D2D2D", foreground="#FFFFFF")
        style.configure("TCombobox",fieldbackground="#2D2D2D", foreground="#FFFFFF")
        style.configure("Treeview", background="#2D2D2D", foreground="#FFFFFF",
                        fieldbackground="#2D2D2D", rowheight=28)
        style.map("Treeview", background=[("selected", "#007ACC")])

        # ---- 控制面板 ----
        ctrl = tk.Frame(self, bg="#1E1E1E")
        ctrl.pack(fill=tk.X, padx=10, pady=8)

        ttk.Label(ctrl, text="工作表:").grid(row=0, column=0, padx=4, sticky="w")
        self.sheet_combo = ttk.Combobox(ctrl, values=list(GAME_SHEETS.keys()), width=16, state="readonly")
        self.sheet_combo.current(0)
        self.sheet_combo.grid(row=0, column=1, padx=4)

        ttk.Label(ctrl, text="關鍵字:").grid(row=0, column=2, padx=4, sticky="w")
        self.keyword_entry = ttk.Entry(ctrl, width=14)
        self.keyword_entry.grid(row=0, column=3, padx=4)

        ttk.Label(ctrl, text="預算:").grid(row=0, column=4, padx=4, sticky="w")
        self.min_price = ttk.Entry(ctrl, width=8)
        self.min_price.grid(row=0, column=5, padx=2)
        ttk.Label(ctrl, text="~").grid(row=0, column=6, padx=1)
        self.max_price = ttk.Entry(ctrl, width=8)
        self.max_price.grid(row=0, column=7, padx=2)

        ttk.Label(ctrl, text="純角CP≤:").grid(row=0, column=8, padx=4, sticky="w")
        self.max_cp = ttk.Entry(ctrl, width=7)
        self.max_cp.grid(row=0, column=9, padx=2)

        self.fetch_btn = ttk.Button(ctrl, text="📡 拉取資料", command=self.async_fetch)
        self.fetch_btn.grid(row=0, column=10, padx=8)
        self.filter_btn = ttk.Button(ctrl, text="🔍 套用篩選", command=self.apply_filter)
        self.filter_btn.grid(row=0, column=11, padx=4)

        # ---- 資料表 ----
        cols = ("日期", "標題", "價格", "金角", "純角CP", "賣家")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        self.tree.heading("日期",  text="發現日期")
        self.tree.column("日期",   width=90,  anchor="center")
        self.tree.heading("標題",  text="標題 (雙擊前往賣場)")
        self.tree.column("標題",   width=540, anchor="w")
        self.tree.heading("價格",  text="價格($)")
        self.tree.column("價格",   width=80,  anchor="center")
        self.tree.heading("金角",  text="金角")
        self.tree.column("金角",   width=50,  anchor="center")
        self.tree.heading("純角CP",text="純角CP")
        self.tree.column("純角CP", width=70,  anchor="center")
        self.tree.heading("賣家",  text="賣家")
        self.tree.column("賣家",   width=120, anchor="center")

        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10,0), pady=4)
        sb.pack(side=tk.LEFT, fill=tk.Y, pady=4)

        self.tree.bind("<Double-1>", self.on_double_click)

        # ---- 狀態列 ----
        self.status_var = tk.StringVar(value="就緒｜按「拉取資料」連線 Google Sheets")
        ttk.Label(self, textvariable=self.status_var).pack(fill=tk.X, padx=10, pady=4)

    # ------------------------------------------------------------------
    def set_status(self, msg):
        """可從執行緒安全呼叫"""
        self.after(0, lambda: self.status_var.set(msg))

    def async_fetch(self):
        """背景執行緒拉資料，避免 UI 凍結"""
        self.fetch_btn.config(state="disabled")
        sheet_key = self.sheet_combo.get()
        sheet_name, _ = GAME_SHEETS[sheet_key]

        def worker():
            rows = fetch_sheet_data(sheet_name, self.set_status)
            self._raw_rows = rows
            self.after(0, self.apply_filter)
            self.after(0, lambda: self.fetch_btn.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def apply_filter(self):
        if not self._raw_rows:
            self.set_status("尚無資料，請先按「拉取資料」")
            return

        _, sheet_type = GAME_SHEETS[self.sheet_combo.get()]
        c = COLS[sheet_type]   # 欄位索引 dict

        kw = self.keyword_entry.get().strip().lower()
        try: p_min = float(self.min_price.get().strip() or "0")
        except: p_min = 0
        try: p_max = float(self.max_price.get().strip() or "999999")
        except: p_max = 999999
        try: max_cp_val = float(self.max_cp.get().strip() or "999")
        except: max_cp_val = 999

        # 清除
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._row_urls.clear()

        count = 0
        for row in self._raw_rows:
            def get(idx):
                return row[idx] if len(row) > idx else ""

            title   = get(c["title"])
            price_s = get(c["price"])
            gold_s  = get(c["gold"])
            cp1_s   = get(c["cp1"])
            seller  = get(c["seller"])
            url     = get(c["url"])
            date_s  = get(c["date"])

            try: price = float(price_s.replace(",", "").replace("$", ""))
            except: price = 0
            try: cp1 = float(cp1_s.replace(",", ""))
            except: cp1 = 999

            if price == 0: continue
            if not (p_min <= price <= p_max): continue
            if kw and kw not in title.lower(): continue
            if max_cp_val < 999 and cp1 > max_cp_val: continue

            iid = self.tree.insert("", "end", values=(
                date_s, title, f"${price:,.0f}", gold_s, cp1_s, seller
            ))
            self._row_urls[iid] = url
            count += 1

        self.set_status(f"顯示 {count} 筆 (共拉取 {len(self._raw_rows)} 筆 | 來源: Google Sheets)")


    def on_double_click(self, event):
        sel = self.tree.selection()
        if not sel: return
        url = self._row_urls.get(sel[0], "")
        if url:
            webbrowser.open(url)
        else:
            messagebox.showinfo("提示", "此列沒有連結資訊")


if __name__ == "__main__":
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        messagebox.showerror("缺少套件", "請先執行：pip install gspread google-auth")
        exit()
    app = RadarApp()
    app.mainloop()
