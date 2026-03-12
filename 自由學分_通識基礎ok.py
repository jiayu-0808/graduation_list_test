import pdfplumber
import pandas as pd
import re

# 1. 設定檔案路徑
pdf_file = "list_two/11123205_黃若雅.pdf" 

FREE_MARKERS = ['(跨)', '(就)', '(微)', '(P)', '(M)']

def count_markers(name):
    """精確計算單一課程具備的標記數量"""
    if not name or name == "None": return 0
    return sum(1 for m in FREE_MARKERS if m in str(name))

def has_score(val):
    """判定課程是否確實修畢"""
    v = str(val).strip().lower()
    return v not in ['', 'none', 'nan', 'pd.na', 'null']

def audit_basic_gen_ed(pdf_path):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
            if table: all_rows.extend(table)

    df = pd.DataFrame(all_rows)
    if df.shape[1] > 11: df = df.iloc[:, :11]
    df.columns = ['類別', '科目_左', '課程名稱_左', '學分_左', '修課學期_右', '科目_右', '課程名稱_右', '性質_右', '學期_右', '學分_右', '分數_右']
    
    # 建立搜尋包以利代碼比對
    df['搜尋包'] = df.apply(lambda row: "".join([str(i) for i in row if i]).replace(" ", "").replace("\n", ""), axis=1)

    # ==========================================
    # --- [核心配置：定義基礎通識的五個坑位] ---
    # ==========================================
    # 根據 image_b54d40.png 與您的說明設定定額 (limit)
    basic_buckets = {
        "天類 (應修2門)": {"codes": ["GQ101", "GQ201"], "limit": 2},
        "人類_經濟政治項 (應修1門)": {"codes": ["GQ392", "GQ393", "GQ394", "GQ395", "GQ396", "GQ397"], "limit": 1},
        "人類_歷史項 (應修1門)": {"codes": ["GQ456", "GQ457"], "limit": 1},
        "物類 (應修2門)": {"codes": ["GE726", "GQ000"], "limit": 2},
        "我類 (應修2門)": {"codes": ["GQ701", "GQ801"], "limit": 2}
    }

    # 用於存放各坑位的修課狀況
    bucket_contents = {name: [] for name in basic_buckets.keys()}

    # 遍歷表格，將課程塞入對應的坑位
    for _, row in df.iterrows():
        search_str = str(row['搜尋包'])
        score = row['分數_右']
        
        if not has_score(score): continue

        for b_name, b_info in basic_buckets.items():
            if any(code in search_str for code in b_info["codes"]):
                course_name = str(row['課程名稱_右']).replace("\n", "").strip()
                bucket_contents[b_name].append({
                    "name": course_name,
                    "markers": count_markers(course_name),
                    "credits": 2.0
                })
                break

    # ==========================================
    # --- [判定邏輯：組內排序與溢出撥入] ---
    # ==========================================
    print("\n" + "="*70)
    print(f"{'基礎通識必修：各類別坑位填補與自由學分撥入報告':^60}")
    print("="*70)

    total_free_from_basic = 0

    for b_name, courses in bucket_contents.items():
        limit = basic_buckets[b_name]["limit"]
        print(f"【{b_name}】")
        
        if not courses:
            print("   ⚠️ 尚未修習任何此類課程")
        else:
            # 關鍵細節：記號最少的排前面，優先拿去抵門檻
            courses.sort(key=lambda x: x['markers'])
            
            # 填坑部分
            for i, c in enumerate(courses):
                if i < limit:
                    print(f"   - [門檻抵扣] {c['name']:<30} | 記號: {c['markers']}")
                else:
                    # 溢出部分：必須具備標記才採計為自由學分
                    if c['markers'] > 0:
                        print(f"   ✅ [撥入自由] {c['name']:<30} | 記號: {c['markers']} | +{c['credits']}學分")
                        total_free_from_basic += c['credits']
                    else:
                        print(f"   ❌ [溢出排除] {c['name']:<30} | 無標記，不採計")
        print("-" * 70)

    print(f"【基礎通識區塊產出之自由學分總計】： {total_free_from_basic} 學分")
    print("="*70)

if __name__ == "__main__":
    audit_basic_gen_ed(pdf_file)