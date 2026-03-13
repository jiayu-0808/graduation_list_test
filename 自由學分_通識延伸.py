import pdfplumber
import pandas as pd
import re

# 1. 設定檔案路徑
pdf_file = "list_one/10823124_邱煒育.pdf" 
FREE_MARKERS = ['(跨)', '(就)', '(微)', '(P)', '(M)']

def count_markers(name):
    if not name or name == "None": return 0
    return sum(1 for m in FREE_MARKERS if m in str(name))

def has_score(val):
    v = str(val).strip().lower()
    return v not in ['', 'none', 'nan', 'pd.na', 'null']

def audit_extension_dynamic_threshold(pdf_path):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
            if table: all_rows.extend(table)

    df = pd.DataFrame(all_rows)
    if df.shape[1] > 11: df = df.iloc[:, :11]
    df.columns = ['類別', '科目_左', '課程名稱_左', '學分_左', '修課學期_右', '科目_右', '課程名稱_右', '性質_右', '學期_右', '學分_右', '分數_右']
    df['搜尋包'] = df.apply(lambda row: "".join([str(i) for i in row if i]).replace(" ", "").replace("\n", ""), axis=1)

    # ==========================================
    # --- [關鍵修改：動態抓取應修門檻] ---
    # ==========================================
    try:
        # 定位「通識延伸選修：」這一列
        threshold_row = df[df['搜尋包'].str.contains("通識延伸選修：", na=False)].iloc[0]
        
        # 1. 抓取左側應修學分 (通常在第 3 欄，索引為 3 或 4)
        # 我們掃描該列左半部，找第一個出現的數字
        left_text = "".join([str(threshold_row[i]) for i in range(5) if threshold_row[i]])
        required_credits_match = re.search(r"(\d+)", left_text)
        required_credits = float(required_credits_match.group(1)) if required_credits_match else 12.0
        
        # 2. 確定區間範圍以利後續計算
        start_idx = df[df['搜尋包'].str.contains("通識基礎必修：", na=False)].index[0]
        end_idx = df[df['搜尋包'].str.contains("通識延伸選修：", na=False)].index[0]
        target_df = df.iloc[start_idx + 1 : end_idx]
    except Exception as e:
        return print(f"❌ 無法定位門檻或區間標籤: {e}")

    # --- [資料收集] ---
    ext_cats = ['天學', '人學', '物學', '我學']
    basic_codes = "GQ101|GQ201|GQ39|GQ45|GE726|GQ000|GQ701|GQ801"
    all_ext_pool = []
    curr_cat = ""

    for _, row in target_df.iterrows():
        cat_cell = str(row['類別']).strip()
        if any(c in cat_cell for c in ext_cats): curr_cat = next(c for c in ext_cats if c in cat_cell)
        if curr_cat and has_score(row['分數_右']) and not re.search(basic_codes, str(row['搜尋包'])):
            name = str(row['課程名稱_右']).replace("\n", "").strip()
            if name and name != "None" and not any(k in name for k in ["課程名稱", "科目"]):
                credit_val = float(row['學分_右']) if str(row['學分_右']).replace('.','').isdigit() else 2.0
                all_ext_pool.append({
                    "name": name, "cat": curr_cat, "markers": count_markers(name), "credits": credit_val
                })

    # ==========================================
    # --- [逆向扣除算法：使用動態門檻] ---
    # ==========================================
    total_earned = sum(c['credits'] for c in all_ext_pool)
    overflow_limit = max(0, total_earned - required_credits) # 使用抓到的 required_credits
    
    # 排序：記號越多、學分大的優先撥出 (極大化)
    all_ext_pool.sort(key=lambda x: (-x['markers'], -x['credits']))
    
    allocated_B = [] # 自由學分撥出區
    current_overflow_sum = 0
    
    for c in all_ext_pool:
        # 條件：有記號、撥出後不超額、且不破壞四類保底
        if c['markers'] > 0 and (current_overflow_sum + c['credits'] <= overflow_limit):
            temp_remaining = [item for item in all_ext_pool if item not in allocated_B and item != c]
            remaining_cats = set(item['cat'] for item in temp_remaining)
            
            # 必須確保剩餘課程仍包含天、人、物、我
            if all(cat in remaining_cats for cat in ext_cats):
                allocated_B.append(c)
                current_overflow_sum += c['credits']

    reserved_A = [c for c in all_ext_pool if c not in allocated_B]

    # --- [結果輸出] ---
    print("\n" + "="*80)
    print(f"{'通識延伸：動態門檻逆向核算報告':^70}")
    print("="*80)
    print(f" 偵測應修門檻：{required_credits} 學分")
    print(f" 目前實得總計：{total_earned} 學分")
    print(f" 可撥出溢出額：{overflow_limit} 學分")
    print("-" * 80)
    
    print(f"【A. 保留在延伸通識】(滿足 {required_credits} 學分門檻)")
    for r in reserved_A:
        print(f" - {r['name']:<35} | {r['cat']:<5} | {r['credits']}學分 | 記號: {r['markers']}")
    print(f" >> A區合計保留：{sum(i['credits'] for i in reserved_A)} 學分數")

    print("\n【B. 撥入自由選修】(具標記之最高價值組合)")
    if allocated_B:
        for b in allocated_B:
            print(f" ✅ {b['name']:<35} | {b['credits']}學分 | 記號: {b['markers']}")
        print(f" >> B區合計撥入：{sum(i['credits'] for i in allocated_B)} 學分數")
    else: print(" (無溢出學分或無符合標記課程)")
    print("="*80)

if __name__ == "__main__":
    audit_extension_dynamic_threshold(pdf_file)
