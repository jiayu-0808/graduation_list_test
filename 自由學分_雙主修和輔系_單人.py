import pdfplumber
import pandas as pd
import re

# ==========================================
# 1. 設定要分析的個人學分表檔案路徑
# ==========================================
pdf_file = "list_two/11123205_黃若雅.pdf" 

def get_personal_special_credits(pdf_path):
    all_rows = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table(table_settings={
                    "vertical_strategy": "lines", 
                    "horizontal_strategy": "lines"
                })
                if table: all_rows.extend(table)
    except Exception as e:
        print(f"❌ 讀取檔案時發生錯誤: {e}")
        return

    if not all_rows:
        print("❌ 在 PDF 中找不到任何表格數據。")
        return

    # 關鍵字設定
    targets = [
        {"key": "實際已修輔系課程總學分數", "label": "輔系"},
        {"key": "實際已修雙主修課程總學分數", "label": "雙主修"}
    ]

    found_any = False
    total_special = 0

    print("\n" + "="*60)
    
    print("="*60)
    print(f"{'類別':<15} | {'詳細內容':<30} | {'採計學分'}")
    print("-" * 60)

    for row in all_rows:
        # 將整列轉為字串（移除空格），檢查是否有關鍵字
        row_str = "".join([str(cell) for cell in row if cell]).replace(" ", "").replace("\n", "")
        
        for target in targets:
            if target['key'] in row_str:
                found_keyword_cell = False
                actual_val = 0
                
                # 精確定位：只找關鍵字「右邊」的數字
                for cell in row:
                    if not cell: continue
                    cell_clean = str(cell).replace(" ", "").replace("\n", "")
                    
                    # 狀況 A：數字跟關鍵字黏在同一個儲存格（且在關鍵字後面）
                    if target['key'] in cell_clean:
                        found_keyword_cell = True
                        after_text = cell_clean.split(target['key'])[-1]
                        nums = re.findall(r"(\d+\.?\d*)", after_text)
                        if nums:
                            actual_val = float(nums[0])
                            break
                    
                    # 狀況 B：關鍵字在左邊格，數字在右邊格
                    elif found_keyword_cell:
                        nums = re.findall(r"(\d+\.?\d*)", cell_clean)
                        if nums:
                            actual_val = float(nums[0])
                            break
                
                if actual_val > 0:
                    print(f"{target['label']:<15} | 偵測到實際已修畢紀錄            | {actual_val:>3} 學分")
                    total_special += actual_val
                    found_any = True

    if not found_any:
        print(f"{'查無紀錄':<15} | 本檔案無輔系或雙主修已修學分        | 0 學分")

    print("-" * 60)
    print(f"{'總計採計':<15} |                                | {total_special:>3} 學分")
    print("="*60)
    print("※ 備註：此學分將直接併入自由選修計算，不需檢查記號。")

if __name__ == "__main__":
    get_personal_special_credits(pdf_file)
