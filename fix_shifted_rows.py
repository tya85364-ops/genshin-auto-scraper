import os
import time
import gspread
from google.oauth2.service_account import Credentials

games = ["原神-成交紀錄", "鳴潮-成交紀錄", "崩鐵-成交紀錄"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GCP_KEY_FILE = os.path.join(BASE_DIR, "gcp_key.json")
G_SHEET_KEY = "1SOt-2DwJVEcEgvuvQfAvW6ue6WcrnvywxPbKIJFEcYI"

def main():
    print("啟動修正錯位腳本...")
    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(GCP_KEY_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    try:
        sh = gc.open_by_key(G_SHEET_KEY)
    except Exception as e:
        print("無法開啟 Spreadsheet", e)
        return
        
    for sheet_name in games:
        try:
            ws = sh.worksheet(sheet_name)
        except Exception as e:
            continue
            
        print(f"檢查分頁：{sheet_name} ...")
        data = ws.get_all_values()
        if not data: continue
        
        updates = []
        # row index API update format is 'range': 'A2:Q2'. 
        # But we can just write it column by column, or rewrite the entire row.
        for i, row in enumerate(data[1:], start=2):
            if len(row) > 14 and str(row[14]).startswith("http"):
                # 這行位移了！
                # 原始錯誤: ... [8] high, [9] high, [10] cp1, [11] cp2, [12] cpw, [13] seller, [14] url, [15] price, [16] price
                # 修正: 剔除 [9]
                print(f"  > 發現位移列 第 {i} 列，進行修復")
                
                # 取得原本內容
                corrected_row = row[:9] + row[10:17] 
                # 長度不足則補空字串以覆蓋原本格子的內容
                while len(corrected_row) < 16:
                    corrected_row.append("")
                    
                # 如果後面有超過 16 格的，也補空字串把它清掉
                if len(row) >= 17:
                    # 我們寫入 A:Q 範圍 (也就是 17 格)，第 17 格放空
                    while len(corrected_row) < len(row):
                        corrected_row.append("")
                
                # 轉成字母座標
                end_col = chr(ord('A') + len(corrected_row) - 1)
                if len(corrected_row) > 26:
                    # just arbitrary safe handling, but 17 is Q
                    pass
                updates.append({'range': f'A{i}:{end_col}{i}', 'values': [corrected_row]})
                
        if updates:
            print(f"  > 將寫入 {len(updates)} 筆修正...")
            # 分批
            for i in range(0, len(updates), 100):
                ws.batch_update(updates[i:i+100])
                time.sleep(1)
        else:
            print("  > 無錯位列")

if __name__ == '__main__':
    main()
