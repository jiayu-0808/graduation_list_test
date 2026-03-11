#20260131
import pdfplumber
import pandas as pd
import re

# 1. 設定檔案路徑
pdf_file = "11123215.pdf" 

def check_graduation_progress(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_rows = []
        for page in pdf.pages:
            table = page.extract_table(table_settings={
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines"
            })
            if table:
                all_rows.extend(table)

    if not all_rows:
        print("錯誤：無法從 PDF 中讀取表格。")
        return

    # --- [建立與清理表格] ---
    df = pd.DataFrame(all_rows)
    if df.shape[1] > 11:
        df = df.iloc[:, :11]
    df.columns = ['類別', '科目_左', '課程名稱_左', '學分_左', '修課學期_右', '科目_右', '課程名稱_右', '性質_右', '學期_右', '學分_右', '分數_右']
    
    df['搜尋包'] = df.apply(lambda row: "".join([str(i) for i in row if i]).replace(" ", "").replace("\n", ""), axis=1)

    def has_score(val):
        v = str(val).strip().lower()
        return v not in ['', 'none', 'nan', 'pd.na']

    # ==========================================
    # --- [第一部分：基本知能檢查] ---
    # ==========================================
    print("\n--- [分析結果：基本知能檢查] ---")
    df['學分_左_數字'] = pd.to_numeric(df['學分_左'], errors='coerce')
    eng_condition = (df['搜尋包'].str.contains('ZA', na=False)) & (df['學分_左_數字'] == 1)
    eng_rows = df[eng_condition]
    if eng_rows.empty:
        print("never [英文]：找不到符合學分為 1 且代碼含 ZA 的英文課程")
    else:
        for _, row in eng_rows.iterrows():
            if has_score(row['分數_右']):
                print(f"ok [英文]：{row['課程名稱_左']} (分數：{row['分數_右']})")
            else:
                print(f"never [英文]：{row['課程名稱_左']} (無分數)")

    pe_condition = (df['搜尋包'].str.contains('體育', na=False)) | (df['搜尋包'].str.contains('GR', na=False))
    pe_rows = df[pe_condition]
    pe_done = pe_rows[pe_rows['分數_右'].apply(has_score)]
    pe_count = len(pe_done)
    print(f"{'ok' if pe_count >= 4 else 'never'} [體育]：已達標 (目前完成：{pe_count}/4 門)")

    # ==========================================
    # --- [第二部分：通識基礎必修檢查] ---
    # ==========================================
    print("\n--- [分析結果：通識基礎必修檢查] ---")
    for code, name, label in [('GQ101', '宗教哲學', '天類'), ('GQ201', '人生哲學', '天類'), 
                               ('GE726', '運算思維', '物類'), ('GQ000', '自然科學', '物類'),
                               ('GQ701', '文學經典', '我類'), ('GQ801', '語文修辭', '我類')]:
        row = df[df['搜尋包'].str.contains(code, na=False)]
        if not row.empty:
            score = str(row.iloc[0]['分數_右']).strip()
            print(f"{'ok' if has_score(score) else 'never'} [{label}]：{code} {name} (分數：{score if has_score(score) else '無'})")
        else:
            print(f"never [{label}]：找不到課程代碼 {code}")

    # --- 人類 (六選一與二選一) ---
    ren1_match = df[df['搜尋包'].str.contains('GQ392|GQ393|GQ394|GQ395|GQ396|GQ397', na=False)]
    ren1_done = ren1_match[ren1_match['分數_右'].apply(has_score)]
    if not ren1_done.empty:
        print(f"ok [人類-1]：已完成 (分數：{ren1_done.iloc[0]['分數_右']})")
    else:
        print("never [人類-1]：六選一尚未完成")

    ren2_match = df[df['搜尋包'].str.contains('GQ456|GQ457', na=False)]
    ren2_done = ren2_match[ren2_match['分數_右'].apply(has_score)]
    if not ren2_done.empty:
        print(f"ok [人類-2]：已完成 (分數：{ren2_done.iloc[0]['分數_右']})")
    else:
        print("never [人類-2]：二選一尚未完成")
    # ==========================================
    # --- [第三部分：通識延伸選修與工程倫理] ---
    # ==========================================
    print("\n--- [分析結果：通識延伸選修檢查] ---")

    # 1. 四大延伸類別檢查
    ext_categories = {'天學': '天延伸', '人學': '人延伸', '物學': '物延伸', '我學': '我延伸'}
    for key, label in ext_categories.items():
        cat_match = df[df['類別'].str.contains(key, na=False) & ~df['搜尋包'].str.contains('GQ101|GQ201|GE726|GQ000|GQ701|GQ801', na=False)]
        cat_done = cat_match[cat_match['分數_右'].apply(has_score)]
        if not cat_done.empty:
            print(f"ok [{label}]：已修習 {cat_done.iloc[0]['課程名稱_右']}")
        else:
            print(f"never [{label}]：尚未修習完成延伸課程")

    # 2. 工程倫理判定
    ethics_match = df[df['課程名稱_右'].str.contains('工程倫理', na=False) & df['分數_右'].apply(has_score)]
    print(f"{'ok' if not ethics_match.empty else 'never'} [工程倫理]：判定結果")

    # 3. 通識延伸學分統計 (關鍵新增)
    try:
        sum_row_ext = df[df['類別'].str.contains('通識延伸選修', na=False)]
        if not sum_row_ext.empty:
            nums_ext = []
            for cell in sum_row_ext.iloc[0]:
                if cell: nums_ext.extend(re.findall(r"(\d+\.?\d*)", str(cell)))
            
            if len(nums_ext) >= 2:
                f_num_ext = nums_ext[0]
                # 處理黏連邏輯
                if len(f_num_ext) >= 4 and "." in f_num_ext:
                    clean_ext = f_num_ext.split('.')[0]
                    left_ext, right_ext = float(clean_ext[:2]), float(clean_ext[2:])
                else:
                    left_ext, right_ext = float(nums_ext[0]), float(nums_ext[1])
                
                diff_ext = left_ext - right_ext
                print(f"\n>> 通識延伸學分統計：")
                print(f"   應修(左)：{left_ext} / 已修(右)：{right_ext}")
                if diff_ext <= 0:
                    print("   結果：✅ 已修畢")
                else:
                    print(f"   結果：❌ 未修畢，尚缺 {int(diff_ext)} 學分")
    except Exception as e:
        print(f"⚠️ 通識延伸學分解析失敗: {e}")

    # ==========================================
    # --- [第四部分：學系必修科目與學分統計] ---
    # ==========================================
    print("\n--- [分析結果：學系必修檢查] ---")
    try:
        start_idx = df[df['搜尋包'].str.contains("通識延伸選修", na=False)].index[0]
        target_indices = df[(df.index > start_idx) & (df['搜尋包'].str.contains("學系必修：|學系必修:", na=False))].index
        
        if not target_indices.empty:
            end_idx = target_indices[0]
            for i in range(start_idx + 1, end_idx):
                row = df.iloc[i]
                name_raw = str(row['課程名稱_左']) if row['課程名稱_左'] else ""
                score_raw = str(row['分數_右']) if row['分數_右'] else ""
                if not name_raw or name_raw == "None" or any(k in name_raw for k in ["課程名稱", "科目"]): continue
                
                names = [n.strip() for n in name_raw.split('\n') if n.strip() and n.strip() != "None"]
                scores = [s.strip() for s in score_raw.split('\n') if s.strip()]
                actual_names = [n for n in names if not n.isalnum()] or names
                for idx, c_name in enumerate(actual_names):
                    s_val = scores[idx] if idx < len(scores) else ""
                    if not has_score(s_val): print(f"[-] {c_name} 未修過")

            summary_row = df.iloc[end_idx]
            nums = []
            for cell in summary_row:
                if cell: nums.extend(re.findall(r"(\d+\.?\d*)", str(cell)))

            if len(nums) >= 2:
                f_num = nums[0]
                if len(f_num) >= 4 and "." in f_num:
                    clean = f_num.split('.')[0]
                    left_val, right_val = float(clean[:2]), float(clean[2:])
                else:
                    left_val, right_val = float(nums[0]), float(nums[1])

                diff = left_val - right_val
                print(f"\n>> 學系必修學分統計：")
                print(f"   應修(左)：{left_val} / 已修(右)：{right_val}")
                if diff <= 0: print("   結果：✅ 已修畢")
                else: print(f"   結果：❌ 未修畢，尚缺 {int(diff)} 學分")

    except Exception as e:
        print(f"⚠️ 學系必修解析失敗: {e}")

if __name__ == "__main__":
    check_graduation_progress(pdf_file)