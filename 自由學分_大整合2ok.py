import pdfplumber
import pandas as pd
import re
import os
from itertools import combinations

# ==========================================
# 1. 設定區
# ==========================================
pdf_file = "list_one/11123106_莊詩涵.pdf" 
FREE_MARKERS = ['(跨)', '(就)', '(微)', '(P)', '(M)']
PROGRAMS = {
    "固力與機設": {
        "基礎": ["高等材料力學", "破壞力學", "機械振動學", "電腦輔助機構分析", "電腦輔助工程一", "電腦輔助工程"],
        "選修、橫跨": ["高等振動學","金屬疲勞分析","複合材料力學","微系統封裝設計與力學","有限元素法","軸承潤滑分析","機電整合彈性體機構模擬","力學程式分析","軸承及潤滑專題","機械振動專題","轉子動力學專題","精密機械設計專題","FORTRAN程式語言","電腦輔助模具設計","共通核心職能課程",
                      "系統動態分析", "噪音振動分析與實作", "系統鑑別", "骨科力學與臨床急診", "機器視覺", "伺服系統", "工程機率與統計", "精密量測", "參數化電腦輔助設計ProE", "參數化設計ProE", "飛行動力學與自動控制", "汽車工程", "高等工程數學", "工程聲學", "機電控制與設計實務", "無人飛行器系統設計", "工業大數據", "創意機構設計與實作", "機構設計中的美學", "儲能創新設計與實作", "機電產業行銷與媒體企劃", "玻璃蝕刻技術", "參數化電腦輔助設計Inventor", "平行與移動計算導論", 
                      "CAD/CAM", "微分方程數值解", "風力機系統設計與分析", "精密模具設計與應用", "電腦輔助設計幾何建模", "塑性加工", "工程應用進階資訊科技", "顯示器製造與自動化技術", "科學應用軟體", "最佳化機械設計", "製程自動化專案設計實務", "Python程式語言", "產品設計與分析", "參數化設計SolidWorks", "Python人工智慧", "半導體及顯示器製造自動化技術", "自動化系統設計與案例分析", "自動化分析", "產業及學術國際領袖體驗講座",
                      "半導體元件特論", "工具機系統設計分析", "精密工具機技術專論", "智慧模具設計/製造", "智慧能源專題", "智慧電網與儲能創新設計與實作", "智慧能源設計與實作", "淨零永續能源創新設計與實作", "淨零永續能源設計實作",
                      "校外工業講座", "工業安全與衛生", "工業物聯網", "機器學習應用", "擴增實境微型顯示導論", "液壓傳動與氣動技術", "軟性電子與顯示器", "智慧模具專題", "智慧自動化專題", "數值方法", "科技英文報告與寫作", "創意性工程設計", "可靠度工程", "工程經濟學", "工業4.0概論一", "工業4.0概論二", "液壓控制", "氣壓控制", "實驗設計與統計分析", "科技英文", "工程英文", "工業日文", "工程德文", "工具機機械設計", "視覺化Android行動應用程式設計", "視覺化APP程式設計", "機械工業講座", "積體電路製程整合與設計最佳化", "積體電路製程整合設計", "奈米材料與元件", "機械工業之人因工程", "碳管理概論與實務", "自主學習課程","實習課程(累計)"]
    },
    "量測與機電控制": {
        "基礎": ["近代控制", "線性代數", "微處理機原理", "數位邏輯設計與控制", "應用電子學", "機電整合應用與實習"],
        "選修、橫跨": ["自動化光電檢測", "非線性控制", "微致動器與感測器原理與應用", "微致動器與感測器設計", "機器人學", "嵌入式系統設計", "感測原理與應用", "磁力軸承系統設計與控制", "變結構控制", "微機電系統", "機器學習", "數位訊號處理", "數位控制", "電腦輔助檢測", "C/C++程式語言", "交流電機控制", "交換式電源供應器設計", "生醫微機電系統", "生醫微機實作", "工具機工程", "進階線性代數工程應用", "智能控制", "微處理機應用", "配線實務一", "機電整合", "配線實務二", "嵌入式系統設計與實作", "機電系統控制與設計實務", "移動式機器人", "智慧機器人實務控制技術", "智慧機器人", "電腦輔助電路模擬設計及PCB佈線", "電路模擬及PCB佈線", "物聯網程式設計實作", "企業物聯網", "PLC及Macro程式設計", "急診與醫療檢測晶片微系統", "急診與醫療檢測微系統", "急診與醫療晶片", "急診與醫療檢測晶片微系統二", "急診與醫療檢測微系統二", "急診與醫療晶片二", "成型整廠聯網技術",
                      "系統動態分析", "網際網路應用", "噪音振動分析與實作", "系統鑑別", "機器視覺", "伺服系統", "工程機率與統計", "精密量測", "參數化電腦輔助設計ProE", "參數化設計ProE", "飛行動力學與自動控制", "積體電路晶片封裝實務", "雷射工程應用", "綠色能源系統設計應用", "半導體元件與量測", "電路板機械加工技術", "熱流實驗方法", "機電控制與設計實務", "無人飛行器系統設計", "射出成型原理與製程", "機構設計中的美學", "電路板智慧講座一", "電路板智慧製造講座一", "電路板智慧講座二", "電路板智慧製造講座二", "參數化電腦輔助設計Inventor", "平行與移動計算導論", "電路板智慧製造技術", "半導體製程與控制", 
                      "微分方程數值解", "風力機系統設計與分析", "工程應用進階資訊科技", "顯示器製造與自動化技術", "科學應用軟體", "最佳化機械設計", "工程分析軟體的應用", "綠色成型設備技術講座", "先進模造成型技術", "節能綠色射出機講座", "Python程式語言", "產品設計與分析", "參數化設計SolidWorks", "半導體及顯示器製造自動化技術", "自動化系統設計與案例分析", "自動化分析", "產業及學術國際領袖體驗講座",
                      "半導體元件特論", "工具機系統設計分析", "精密工具機技術專論", "智慧模具設計/製造", "智慧能源專題", "智慧電網與儲能創新設計與實作", "智慧能源設計與實作", "淨零永續能源創新設計與實作", "淨零永續能源設計實作",
                      "校外工業講座", "工業安全與衛生", "工業物聯網", "機器學習應用", "擴增實境微型顯示導論", "液壓傳動與氣動技術", "軟性電子與顯示器", "智慧模具專題", "智慧自動化專題", "數值方法", "科技英文報告與寫作", "創意性工程設計", "可靠度工程", "工程經濟學", "工業4.0概論一", "工業4.0概論二", "液壓控制", "氣壓控制", "實驗設計與統計分析", "科技英文", "工程英文", "工業日文", "工程德文", "工具機機械設計", "視覺化Android行動應用程式設計", "視覺化APP程式設計", "機械工業講座", "積體電路製程整合與設計最佳化", "積體電路製程整合設計", "奈米材料與元件", "機械工業之人因工程", "碳管理概論與實務", "自主學習課程","實習課程(累計)"]
    },
    "材料與製造": {
        "基礎": ["工程材料二", "高等材料力學", "電腦輔助模具設計", "射出成型原理與製程", "電腦輔助工程一"],
        "選修、橫跨": ["薄膜技術", "光電半導體製程與設備", "VLSI可靠度工程", "有機半導體材料與元件", "微電子工程與整合技術", "微影技術概論", "積層製造", "生物力學與材料運用", "儲能技術實務", "線性代數", "空污處理設備理論與實務", "固態照明技術與原理", "太陽能電池", "3D IC與先進封裝", 
                      "低維度半導體材料與元件", "網際網路應用", "骨科力學與臨床急診", "積體電路晶片封裝實務", "雷射工程應用", "半導體元件與量測", "電路板機械加工技術", "綠色製造科技", "電腦輔助分析與應用", "綠色塑膠產品製程概論", "模具概論", "模具製造", "模具材料熱處理", "塑膠加工模具設計&CAE", "模具設計實務", "智慧型模具生產技術", "精密製造", "人工智慧的工業應用", "創意機構設計與實作", "電路板智慧講座一", "電路板智慧製造講座一", "電路板智慧講座二", "電路板智慧製造講座二", "儲能創新設計與實作", "機電產業行銷與媒體企劃", "玻璃蝕刻技術", "電路板智慧製造技術", "電半導體製程與控制", 
                      "CAD/CAM", "精密模具設計與應用", "電腦輔助設計幾何建模", "塑性加工", "工程應用進階資訊科技", "顯示器製造與自動化技術", "科學應用軟體", "製程自動化專案設計實務", "工程分析軟體的應用", "綠色成型設備技術講座", "先進模造成型技術", "節能綠色射出機講座", "半導體及顯示器製造自動化技術", "自動化系統設計與案例分析", "自動化分析", 
                      "半導體元件特論", "工具機系統設計分析", "精密工具機技術專論", "智慧模具設計/製造", "智慧能源專題", "智慧電網與儲能創新設計與實作", "智慧能源設計與實作", "淨零永續能源創新設計與實作", "淨零永續能源設計實作",
                      "校外工業講座", "工業安全與衛生", "工業物聯網", "機器學習應用", "擴增實境微型顯示導論", "液壓傳動與氣動技術", "軟性電子與顯示器", "智慧模具專題", "智慧自動化專題", "數值方法", "科技英文報告與寫作", "創意性工程設計", "可靠度工程", "工程經濟學", "工業4.0概論一", "工業4.0概論二", "液壓控制", "氣壓控制", "實驗設計與統計分析", "科技英文", "工程英文", "工業日文", "工程德文", "工具機機械設計", "視覺化Android行動應用程式設計", "視覺化APP程式設計", "機械工業講座", "積體電路製程整合與設計最佳化", "積體電路製程整合設計", "奈米材料與元件", "機械工業之人因工程", "碳管理概論與實務", "自主學習課程","實習課程(累計)"]
    },
    "熱流與能源": {
        "基礎": ["工程數學三", "流體力學導論", "能源工程", "冷凍空調", "熱對流", "進階熱力學"],
        "選修、橫跨": ["熱流特論", "黏滯性流", "質量傳遞", "空氣動力學", "熱交換系統", "電腦輔助工程二", "燃料電池", "燃燒學原理與應用", "電子冷卻", "流體機械", "航空英文", "再生能源技術", "儲能系統熱交換原理", "綠色文明與永續發展", "永續能源跨域應用實務", "奈米流體理論與應用", "熱管理技術", 
                      "汽車工程", "高等工程數學", "工程聲學", "綠色能源系統設計應用", "熱流實驗方法", "計算流體力學", 
                      "微分方程數值解", "風力機系統設計與分析", "製程自動化專案設計實務", "工程分析軟體的應用", "Python人工智慧", 
                      "半導體元件特論", "智慧電網與儲能創新設計與實作", "智慧能源設計與實作", "淨零永續能源創新設計與實作", "淨零永續能源設計實作", 
                      "校外工業講座", "工業安全與衛生", "工業物聯網", "機器學習應用", "擴增實境微型顯示導論", "液壓傳動與氣動技術", "軟性電子與顯示器", "智慧模具專題", "智慧自動化專題", "數值方法", "科技英文報告與寫作", "創意性工程設計", "可靠度工程", "工程經濟學", "工業4.0概論一", "工業4.0概論二", "液壓控制", "氣壓控制", "實驗設計與統計分析", "科技英文", "工程英文", "工業日文", "工程德文", "工具機機械設計", "視覺化Android行動應用程式設計", "視覺化APP程式設計", "機械工業講座", "積體電路製程整合與設計最佳化", "積體電路製程整合設計", "奈米材料與元件", "機械工業之人因工程", "碳管理概論與實務", "自主學習課程","實習課程(累計)"]
    },
    "模具與成型": {
        "基礎": ["工程材料二", "高等材料力學", "參數化電腦輔助設計ProE", "參數化設計ProE", "電腦輔助模具設計", "工程數學三", "流體力學導論", "近代控制", "線性代數", "微處理機原理", "數位邏輯設計與控制"],
        "選修、橫跨": ["參數化設計 NX", "模具加工製造學", "射出成型機實務講座", "先進成型模具設計", "成型導引與試模", "知識化模具設計導引", "射出成型整廠聯網整合技術", "RL增強式學習在智慧製造的實作應用", "增強式學習應用實作", "生產自動化排程", "計算熱流", 
                      "低維度半導體材料與元件", "電腦輔助分析與應用", "綠色塑膠產品製程概論", "模具概論", "模具製造", "模具材料熱處理", "塑膠加工模具設計&CAE", "模具設計實務", "智慧型模具生產技術", "計算流體力學", "精密製造", "人工智慧的工業應用", "工業大數據", "射出成型原理與製程", 
                      "CAD/CAM", "精密模具設計與應用", "電腦輔助設計幾何建模", "塑性加工", "最佳化機械設計", "綠色成型設備技術講座", "先進模造成型技術", "節能綠色射出機講座", "Python程式語言", "產品設計與分析", "參數化設計SolidWorks", "Python人工智慧", "產業及學術國際領袖體驗講座", 
                      "工具機系統設計分析", "精密工具機技術專論", "智慧模具設計/製造", "智慧能源專題", 
                      "校外工業講座", "工業安全與衛生", "工業物聯網", "機器學習應用", "擴增實境微型顯示導論", "液壓傳動與氣動技術", "軟性電子與顯示器", "智慧模具專題", "智慧自動化專題", "數值方法", "科技英文報告與寫作", "創意性工程設計", "可靠度工程", "工程經濟學", "工業4.0概論一", "工業4.0概論二", "液壓控制", "氣壓控制", "實驗設計與統計分析", "科技英文", "工程英文", "工業日文", "工程德文", "工具機機械設計", "視覺化Android行動應用程式設計", "視覺化APP程式設計", "機械工業講座", "積體電路製程整合與設計最佳化", "積體電路製程整合設計", "奈米材料與元件", "機械工業之人因工程", "碳管理概論與實務", "自主學習課程","實習課程(累計)"]
    }
}

INTERNSHIP_KEYWORDS = ["校外工業實習(一)", "校外工業實習(二)","校外工業實習(三)","校外實習", "暑期實習"]


# ==========================================
# 2. 輔助運算函式
# ==========================================
def count_markers(name):
    if not name or name == "None": return 0
    return sum(1 for m in FREE_MARKERS if m in str(name))

def has_score(val):
    v = str(val).strip().lower()
    return v not in ['', 'none', 'nan', 'pd.na', 'null']

def smart_normalize(name):
    temp = str(name).replace('（', '(').replace('）', ')').replace(' ', '').replace('\n', '')
    if any(k in temp for k in ["工程數學", "電腦輔助工程", "工業4.0概論", "配線實務"]):
        matches = re.findall(r'\((.*?)\)', temp)
        for content in matches:
            if all(char in "一二三123" for char in content):
                temp = temp.replace(f"({content})", content)
            else:
                temp = temp.replace(f"({content})", "")
    return re.sub(r'\(.*?\)', '', temp).strip()

def find_best_pit_filler(candidates, base_list, target=10.0):
    best_combo = []; min_marked_sum = float('inf'); best_total = 0
    no_marker = [c for c in candidates if c['markers'] == 0]
    has_marker = [c for c in candidates if c['markers'] > 0]
    has_marker.sort(key=lambda x: -x['credits'])
    
    for r in range(0, min(len(has_marker) + 1, 6)):
        for combo in combinations(has_marker, r):
            test_combo = no_marker + list(combo)
            test_sum = sum(c['credits'] for c in test_combo)
            if test_sum >= target:
                if any(c['clean'] in base_list for c in test_combo):
                    marked_val = sum(c['credits'] for c in combo)
                    if marked_val < min_marked_sum:
                        min_marked_sum = marked_val; best_combo = test_combo; best_total = test_sum
                    elif marked_val == min_marked_sum:
                        if test_sum < best_total or best_total == 0:
                            best_combo = test_combo; best_total = test_sum
    return best_combo, best_total

# ==========================================
# 3. 整合核算主程式
# ==========================================
def run_ultimate_integrated_audit(pdf_path):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table: all_rows.extend(table)

    df = pd.DataFrame(all_rows)
    if df.shape[1] > 11: df = df.iloc[:, :11]
    df.columns = ['類別', '科目_左', '課程名稱_左', '學分_左', '修課學期_右', '科目_右', '課程名稱_右', '性質_右', '學期_右', '學分_右', '分數_右']
    df['搜尋包'] = df.apply(lambda row: "".join([str(i) for i in row if i]).replace(" ", "").replace("\n", ""), axis=1)

    free_results = []

    print("\n" + "═"*95)
    print(f"║ {'全方位自由學分整合核算報告 (官方標籤 + 學程尋優 + 通識溢出)':^66} ║")
    print("═"*95)

    # --- [第一部分：通識基礎與延伸溢出] (依據第二個檔案邏輯) ---
    # 基礎溢出
    basic_buckets = {"天類": ["GQ101", "GQ201"], "人類_政": ["GQ39"], "人類_歷": ["GQ45"], "物類": ["GE726", "GQ000"], "我類": ["GQ701", "GQ801"]}
    for b_name, codes in basic_buckets.items():
        limit = 1 if "政" in b_name or "歷" in b_name else 2
        pool = []
        for _, row in df.iterrows():
            if any(c in str(row['搜尋包']) for c in codes) and has_score(row['分數_右']):
                name = str(row['課程名稱_右']).strip()
                pool.append({"name": name, "markers": count_markers(name), "credits": 2.0})
        pool.sort(key=lambda x: x['markers'])
        if len(pool) > limit:
            for extra in pool[limit:]:
                if extra['markers'] > 0: free_results.append({"module": "基礎溢出", "name": extra['name'], "credits": 2.0})

    # 延伸溢出
    try:
        t_row_ext = df[df['搜尋包'].str.contains("通識延伸選修：", na=False)].iloc[0]
        ext_req = float(re.search(r"(\d+)", "".join([str(t_row_ext.iloc[i]) for i in range(5) if t_row_ext.iloc[i]])).group(1))
        e_start = df[df['搜尋包'].str.contains("通識基礎必修：", na=False)].index[0]; e_end = df[df['搜尋包'].str.contains("通識延伸選修：", na=False)].index[0]
        ext_pool = []
        curr_cat = ""
        for _, row in df.iloc[e_start+1 : e_end].iterrows():
            if any(c in str(row['類別']) for c in ['天學','人學','物學','我學']): curr_cat = str(row['類別']).strip()[:2]
            if curr_cat and has_score(row['分數_右']) and not re.search("GQ101|GQ201|GQ39|GQ45|GE726|GQ000|GQ701|GQ801", str(row['搜尋包'])):
                name = str(row['課程名稱_右']).replace("\n","").strip()
                ext_pool.append({"name": name, "cat": curr_cat, "markers": count_markers(name), "credits": float(row['學分_右']) if str(row['學分_右']).replace('.','').isdigit() else 2.0})
        ext_overflow = max(0, sum(c['credits'] for c in ext_pool) - ext_req)
        ext_pool.sort(key=lambda x: (-x['markers'], -x['credits']))
        ext_sum = 0; ext_allocated = []
        for c in ext_pool:
            if c['markers'] > 0 and (ext_sum + c['credits'] <= ext_overflow):
                temp_rem = [item for item in ext_pool if item not in ext_allocated and item != c]
                if all(cat in set(i['cat'] for i in temp_rem) for cat in ['天學','人學','物學','我學']):
                    ext_allocated.append(c); ext_sum += c['credits']; free_results.append({"module": "延伸溢出", "name": c['name'], "credits": c['credits']})
    except: pass

    # --- [第二部分：學系選修學程尋優] (依據第二個檔案邏輯) ---
    try:
        t_row_m = df[df['搜尋包'].str.contains("學系選修：", na=False)].iloc[0]
        major_req = float(re.search(r"(\d+)", "".join([str(t_row_m.iloc[i]) for i in range(5) if t_row_m.iloc[i]])).group(1))
        m_start = df[df['搜尋包'].str.contains("學系必修：", na=False)].index[0]; m_end = df[df['搜尋包'].str.contains("學系選修：", na=False)].index[0]
        major_pool = []
        intern_total = 0
        for _, row in df.iloc[m_start+1 : m_end].iterrows():
            if has_score(row['分數_右']):
                raw_n = str(row['課程名稱_右']).strip(); clean_n = smart_normalize(raw_n)
                actual_c = float(row['學分_右'])
                if any(k in clean_n for k in INTERNSHIP_KEYWORDS): intern_total += actual_c
                else: major_pool.append({"raw": raw_n, "clean": clean_n, "credits": actual_c, "markers": count_markers(raw_n)})
        if intern_total > 0: major_pool.append({"raw": "實習課程(累計)", "clean": "實習課程(累計)", "credits": min(intern_total, 4.0), "markers": 0})

        sims = []
        for p_name, p_data in PROGRAMS.items():
            base_list = [smart_normalize(b) for b in p_data["基礎"]]; ext_list = [smart_normalize(e) for e in p_data["選修、橫跨"]]
            candidates = [c for c in major_pool if c['clean'] in base_list or c['clean'] in ext_list]
            locked, locked_sum = find_best_pit_filler(candidates, base_list)
            if locked_sum >= 10.0:
                m_overflow = max(0, sum(c['credits'] for c in major_pool) - major_req)
                rem = [c for c in major_pool if c not in locked]; rem.sort(key=lambda x: (-x['markers'], -x['credits']))
                s_free = 0
                for c in rem:
                    if c['markers'] > 0 and (s_free + c['credits'] <= m_overflow): s_free += c['credits']
                sims.append({"name": p_name, "locked": locked, "free_sum": s_free})
        if sims:
            sims.sort(key=lambda x: -x['free_sum']); best_p = sims[0]
            final_m_rem = [c for c in major_pool if c not in best_p['locked']]
            final_m_rem.sort(key=lambda x: (-x['markers'], -x['credits']))
            m_allocated = 0
            for c in final_m_rem:
                if c['markers'] > 0 and (m_allocated + c['credits'] <= max(0, sum(i['credits'] for i in major_pool) - major_req)):
                    m_allocated += c['credits']; free_results.append({"module": f"學系溢出({best_p['name'][:2]})", "name": c['raw'], "credits": c['credits']})
    except: pass

    # --- [第三部分：抓取自由選修與雙輔學分] (依據第一個檔案邏輯) ---
    # 1. 自由選修欄位抓取
    for row in all_rows:
        cells = [str(c).strip() for c in row if c is not None]
        row_str = " ".join(cells).replace("\n", " ")
        if "自由選修：" in row_str:
            after_text = row_str.split("自由選修：")[-1]
            nums = re.findall(r"(\d+\.?\d*)", after_text)
            if nums:
                free_results.append({"module": "官方標籤", "name": "自由選修欄位", "credits": float(nums[-1])})
                break

    # 2. 輔系雙修欄位抓取
    targets = [{"key": "實際已修輔系課程總學分數", "label": "輔系"}, {"key": "實際已修雙主修課程總學分數", "label": "雙主修"}]
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
                        if nums: actual_val = float(nums[0]); break
                    elif found_keyword_cell:
                        nums = re.findall(r"(\d+\.?\d*)", cell_clean)
                        if nums: actual_val = float(nums[0]); break
                if actual_val > 0: free_results.append({"module": f"官方標籤({target['label']})", "name": target['key'], "credits": actual_val})

    # --- [報表總結輸出] ---
    print(f"\n{'來源模組':<30} | {'項目名稱':<45} | {'學分'}")
    print("-" * 95)
    total_sum = sum(item['credits'] for item in free_results)
    for item in free_results:
        print(f"{item['module']:<30} | {item['name']:<45} | {item['credits']:>5.1f}")
    print("-" * 95)
    print(f"{'自由學分總核計 (全方位極大化結果)':>78} : {total_sum:>5.1f} 學分")
    print("═"*95 + "\n")

if __name__ == "__main__":
    if os.path.exists(pdf_file): run_ultimate_integrated_audit(pdf_file)
    else: print(f"❌ 找不到檔案")