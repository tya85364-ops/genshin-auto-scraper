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
    print("啟動表格清洗與對齊腳本...")
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
                continue
                
            print(f"檢查分頁：{sheet_name} ...")
            data = ws.get_all_values()
            if not data: continue
            
            # 尋找異常夾雜的標題列 (例如在崩鐵-成交紀錄底下出現了重複的表頭)
            # 倒序尋找，這樣刪除列時才不會影響前面的 index
            bad_indices = []
            for i in range(len(data) - 1, 0, -1):  # 不檢查 index 0 (正常的第一列)
                row = data[i]
                if not row: continue
                # 如果遇到被縮排或被弄亂的標題列 (例如含有 '標題' 或 '成交發現日')
                if "標題" in row or "成交發現日" in row or "發現時間" in row or row[0] == "":
                    # 避免誤刪真的沒有第一欄但有其它有效資料的列，檢查這被視為標題列的特徵
                    if "標題" in row or row[0] in ["成交發現日", "發現時間"]:
                        print(f"  > 發現異常標題列於第 {i+1} 列，標記刪除")
                        bad_indices.append(i + 1)
            
            # 如果發現空行很多或者多餘標題，進行刪除
            for idx in bad_indices:
                ws.delete_rows(idx)
                print(f"  > 成功刪除第 {idx} 列的異常標題！")
                time.sleep(1)
                
            # 加入「自動對齊與自動欄寬」格式化請求
            sheet_id = ws.id
            formats.append({
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": 16
                    }
                }
            })
            formats.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1, 
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER", # 置中
                            "verticalAlignment": "MIDDLE",
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"
                }
            })
            
    if formats:
        print("套用整體表格自動對齊...")
        try:
            sh.batch_update({"requests": formats})
            print("對齊格式化成功！")
        except Exception as e:
            print("對齊失敗：", e)
            
    print("全部整理完成！")

if __name__ == '__main__':
    main()
