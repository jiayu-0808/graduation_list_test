import pdfplumber
import pandas as pd
import re

# ==========================================
# 1. 設定要分析的個人學分表路徑
# ==========================================
pdf_file = "list_two/10924160_戴敏峯.pdf" 

def get_free_elective_total_fixed(pdf_path):
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
        return 0

    if not all_rows:
        print("❌ 找不到表格數據。")
        return 0

    target_key = "自由選修："
    found_val = 0
    found_row_display = ""

    print("\n" + "="*60)
    
    for row in all_rows:
        # --- 修正點：使用空格 " " 來連結儲存格，避免數字黏在一起 ---
        # 濾掉 None 並轉為字串
        cells = [str(c).strip() for c in row if c is not None]
        row_str = " ".join(cells).replace("\n", " ")
        
        if target_key in row_str:
            # 找到標籤後，取標籤右邊的內容
            after_text = row_str.split(target_key)[-1]
            
            # 使用正則表達式找出所有數字 (會找到 ['14', '3'])
            nums = re.findall(r"(\d+\.?\d*)", after_text)
            
            if nums:
                # 實修學分通常在該列的最右邊 (最後一個數字)
                found_val = float(nums[-1])
                found_row_display = row_str
                break

    if found_val > 0:
        print(f"✅ 偵測成功！")
        print(f"偵測標籤：{target_key}")
        print("-" * 60)
        print(f"採計自由學分：{found_val} 學分")
    else:
        # 如果最後結果是 0，可能是因為該欄位確實是空的
        print(f"ℹ️ 在『{target_key}』欄位偵測到的數值為 0 或無紀錄。")

    print("="*60)
    return found_val

if __name__ == "__main__":
    get_free_elective_total_fixed(pdf_file)
