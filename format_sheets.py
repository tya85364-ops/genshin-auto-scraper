import os
import time
import gspread
from google.oauth2.service_account import Credentials

games = [
    ("原神", "原神-成交紀錄"),
    ("鳴潮", "鳴潮-成交紀錄"),
    ("崩鐵", "崩鐵-成交紀錄")
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GCP_KEY_FILE = os.path.join(BASE_DIR, "gcp_key.json")
G_SHEET_KEY = "1SOt-2DwJVEcEgvuvQfAvW6ue6WcrnvywxPbKIJFEcYI"

def main():
    print("啟動表格修復與回補腳本...")
    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(GCP_KEY_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    try:
        sh = gc.open_by_key(G_SHEET_KEY)
    except Exception as e:
        print("無法開啟 Spreadsheet", e)
        return
        
    formats = []
    
    for active, history in games:
        for sheet_name in [active, history]:
            try:
                ws = sh.worksheet(sheet_name)
            except Exception as e:
                print(f"找不到 {sheet_name}")
                continue
                
            print(f"處理分頁：{sheet_name} ...")
            data = ws.get_all_values()
            if not data or len(data) < 1: 
                continue
                
            header = data[0]
            price_idx = -1
            try:
                price_idx = header.index("價格")
            except ValueError:
                print("  找不到價格欄位，跳過。")
                continue
                
            # 填補高低價，統一在 index=14(O) 和 index=15(P)
            # 因為剛剛的 HEADERS 是把歷來低價和高價放在最後，也就是 14, 15。
            updates = []
            for i, row in enumerate(data[1:], start=2):
                if len(row) <= price_idx: continue
                price = row[price_idx]
                if not price.strip(): continue
                # 拿掉可能存在的逗號與 $
                price = price.replace(',', '').replace('$', '').strip()
                
                col_o = row[14] if len(row) > 14 else ""
                col_p = row[15] if len(row) > 15 else ""
                
                update_needed = False
                new_o = col_o.strip()
                new_p = col_p.strip()
                
                # 如果沒有高低價資料，使用現在目前的標價來當作初始歷史高低價
                if not new_o or new_o == "-":
                    new_o = price
                    update_needed = True
                if not new_p or new_p == "-":
                    new_p = price
                    update_needed = True
                    
                if update_needed:
                    updates.append({'range': f'O{i}:P{i}', 'values': [[new_o, new_p]]})
            
            if updates:
                print(f"  > 寫入 {len(updates)} 筆補齊資料...")
                for i in range(0, len(updates), 500):
                    batch = updates[i:i+500]
                    ws.batch_update(batch)
                    time.sleep(1)
            else:
                print("  > 無空缺需補齊")
                
            sheet_id = ws.id
            formats.append({
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "gridProperties": {
                            "frozenRowCount": 1
                        }
                    },
                    "fields": "gridProperties.frozenRowCount"
                }
            })
            formats.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "bold": True
                            },
                            "horizontalAlignment": "CENTER",
                            "backgroundColor": {
                                "red": 0.9, "green": 0.9, "blue": 0.9
                            }
                        }
                    },
                    "fields": "userEnteredFormat(textFormat,horizontalAlignment,backgroundColor)"
                }
            })
            
    if formats:
        print("套用外觀格式化...")
        try:
            sh.batch_update({"requests": formats})
            print("格式化成功！")
        except Exception as e:
            print("格式化失敗：", e)
            
    print("全部完成！")

if __name__ == '__main__':
    main()
