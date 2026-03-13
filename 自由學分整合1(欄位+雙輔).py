import pdfplumber
import re

pdf_file = "list_one/11123162_王凱民.pdf"

# --- 1. 完全保留 自由選修_欄位.py 邏輯 ---
def get_free_val(all_rows):
    target_key = "自由選修："
    for row in all_rows:
        cells = [str(c).strip() for c in row if c is not None]
        row_str = " ".join(cells).replace("\n", " ")
        if target_key in row_str:
            after_text = row_str.split(target_key)[-1]
            nums = re.findall(r"(\d+\.?\d*)", after_text)
            if nums:
                return float(nums[-1])
    return 0

# --- 2. 完全保留 自由學分_雙主修和輔系_單人.py 邏輯 ---
def get_special_val(all_rows):
    targets = [
        {"key": "實際已修輔系課程總學分數", "label": "輔系"},
        {"key": "實際已修雙主修課程總學分數", "label": "雙主修"}
    ]
    total_special = 0
    for row in all_rows:
        row_str = "".join([str(cell) for cell in row if cell]).replace(" ", "").replace("\n", "")
        for target in targets:
            if target['key'] in row_str:
                found_keyword_cell = False
                actual_val = 0
                for cell in row:
                    if not cell: continue
                    cell_clean = str(cell).replace(" ", "").replace("\n", "")
                    if target['key'] in cell_clean:
                        found_keyword_cell = True
                        after_text = cell_clean.split(target['key'])[-1]
                        nums = re.findall(r"(\d+\.?\d*)", after_text)
                        if nums:
                            actual_val = float(nums[0]); break
                    elif found_keyword_cell:
                        nums = re.findall(r"(\d+\.?\d*)", cell_clean)
                        if nums:
                            actual_val = float(nums[0]); break
                total_special += actual_val
    return total_special

# --- 3. 整合統計輸出 ---
with pdfplumber.open(pdf_file) as pdf:
    all_rows = []
    for page in pdf.pages:
        table = page.extract_table()
        if table: all_rows.extend(table)

v1 = get_free_val(all_rows)    # 抓取自由選修欄位
v2 = get_special_val(all_rows) # 抓取輔系雙修欄位

print(f"1. 官方自由選修欄位：{v1} 學分")
print(f"2. 輔系/雙主修已修學分：{v2} 學分")
print(f"【總計自由學分】：{v1 + v2} 學分")