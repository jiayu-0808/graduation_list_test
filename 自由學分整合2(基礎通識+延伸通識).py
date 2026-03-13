import pdfplumber
import pandas as pd
import re
import os

# ==========================================
# 1. 設定區 (更換檔案路徑即可分析不同學生)
# ==========================================
pdf_file = "list_two/11123215_賴奕嘉.pdf" 

FREE_MARKERS = ['(跨)', '(就)', '(微)', '(P)', '(M)']

# ==========================================
# 2. 核心輔助函式 (不改動邏輯)
# ==========================================
def count_markers(name):
    if not name or name == "None": return 0
    return sum(1 for m in FREE_MARKERS if m in str(name))

def has_score(val):
    v = str(val).strip().lower()
    return v not in ['', 'none', 'nan', 'pd.na', 'null']

# ==========================================
# 3. 整合主程式
# ==========================================
def audit_all_gen_ed_overflow(pdf_path):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
            if table: all_rows.extend(table)

    if not all_rows: return print("❌ 讀取失敗")

    df = pd.DataFrame(all_rows)
    if df.shape[1] > 11: df = df.iloc[:, :11]
    df.columns = ['類別', '科目_左', '課程名稱_左', '學分_左', '修課學期_右', '科目_右', '課程名稱_右', '性質_右', '學期_右', '學分_右', '分數_右']
    df['搜尋包'] = df.apply(lambda row: "".join([str(i) for i in row if i]).replace(" ", "").replace("\n", ""), axis=1)

    free_credit_detail = [] # 儲存採計明細

    print("\n" + "═"*90)
    print(f"║ {'通識溢出自由學分整合採計 (基礎 + 延伸)':^66} ║")
    print("═"*90)

    # ------------------------------------------
    # --- [第一部分：通識基礎溢出邏輯] ---
    # ------------------------------------------
    basic_buckets = {
        "天類": {"codes": ["GQ101", "GQ201"], "limit": 2},
        "人類_經政": {"codes": ["GQ392", "GQ393", "GQ394", "GQ395", "GQ396", "GQ397"], "limit": 1},
        "人類_歷史": {"codes": ["GQ456", "GQ457"], "limit": 1},
        "物類": {"codes": ["GE726", "GQ000"], "limit": 2},
        "我類": {"codes": ["GQ701", "GQ801"], "limit": 2}
    }
    bucket_contents = {name: [] for name in basic_buckets.keys()}

    for _, row in df.iterrows():
        if not has_score(row['分數_右']): continue
        for b_name, b_info in basic_buckets.items():
            if any(code in str(row['搜尋包']) for code in b_info["codes"]):
                name = str(row['課程名稱_右']).replace("\n", "").strip()
                bucket_contents[b_name].append({"name": name, "markers": count_markers(name), "credits": 2.0})
                break

    print(f"【模組一：基礎通識填坑與溢出】")
    for b_name, courses in bucket_contents.items():
        limit = basic_buckets[b_name]["limit"]
        courses.sort(key=lambda x: x['markers']) # 記號少者填門檻
        for i, c in enumerate(courses):
            if i >= limit and c['markers'] > 0: # 超過定額且具記號
                free_credit_detail.append({"source": "基礎溢出", "name": c['name'], "credits": 2.0})
                print(f" ✅ [採計] {c['name']:<30} | {b_name} 溢出")

    # ------------------------------------------
    # --- [第二部分：通識延伸溢出邏輯] ---
    # ------------------------------------------
    print("\n【模組二：延伸通識動態扣除】")
    try:
        # 定位與門檻偵測
        threshold_row = df[df['搜尋包'].str.contains("通識延伸選修：", na=False)].iloc[0]
        left_text = "".join([str(threshold_row.iloc[i]) for i in range(5) if threshold_row.iloc[i]])
        req_credits = float(re.search(r"(\d+)", left_text).group(1))
        
        start_idx = df[df['搜尋包'].str.contains("通識基礎必修：", na=False)].index[0]
        end_idx = df[df['搜尋包'].str.contains("通識延伸選修：", na=False)].index[0]
        target_df = df.iloc[start_idx + 1 : end_idx]

        ext_cats = ['天學', '人學', '物學', '我學']
        ext_pool = []
        curr_cat = ""
        for _, row in target_df.iterrows():
            cat_cell = str(row['類別']).strip()
            if any(c in cat_cell for c in ext_cats): curr_cat = next(c for c in ext_cats if c in cat_cell)
            if curr_cat and has_score(row['分數_右']) and not re.search("GQ101|GQ201|GE726|GQ000|GQ701|GQ801", str(row['搜尋包'])):
                name = str(row['課程名稱_右']).replace("\n", "").strip()
                if name and name != "None" and not any(k in name for k in ["課程名稱", "科目"]):
                    c_val = float(row['學分_右']) if str(row['學分_右']).replace('.','').isdigit() else 2.0
                    ext_pool.append({"name": name, "cat": curr_cat, "markers": count_markers(name), "credits": c_val})

        ext_total = sum(c['credits'] for c in ext_pool)
        overflow_limit = max(0, ext_total - req_credits)
        ext_pool.sort(key=lambda x: (-x['markers'], -x['credits'])) # 記號多者優先撥出

        allocated_ext_sum = 0
        allocated_ext_items = []
        for c in ext_pool:
            if c['markers'] > 0 and (allocated_ext_sum + c['credits'] <= overflow_limit):
                temp_rem = [item for item in ext_pool if item not in allocated_ext_items and item != c]
                if all(cat in set(i['cat'] for i in temp_rem) for cat in ext_cats):
                    allocated_ext_items.append(c)
                    allocated_ext_sum += c['credits']
                    free_credit_detail.append({"source": "延伸溢出", "name": c['name'], "credits": c['credits']})
                    print(f" ✅ [採計] {c['name']:<30} | 溢出撥入")
        print(f" >> 延伸通識偵測門檻：{req_credits} 學分 (目前實得 {ext_total})")
    except Exception as e:
        print(f" ⚠️ 延伸通識解析中斷: {e}")

    # ------------------------------------------
    # --- [結果總計輸出] ---
    # ------------------------------------------
    print("\n" + "─"*90)
    print(f"{'來源項目':<15} | {'課程名稱':<40} | {'採計學分'}")
    print("-" * 90)
    total_sum = 0
    for item in free_credit_detail:
        print(f"{item['source']:<15} | {item['name']:<40} | {item['credits']:>5.1f}")
        total_sum += item['credits']
    print("-" * 90)
    print(f"{'自由學分 (通識溢出部分) 總計':>57} : {total_sum:>5.1f} 學分")
    print("═"*90 + "\n")

if __name__ == "__main__":
    if os.path.exists(pdf_file):
        audit_all_gen_ed_overflow(pdf_file)