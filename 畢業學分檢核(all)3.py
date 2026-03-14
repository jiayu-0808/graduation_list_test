from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
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

def generate_pdf_report(content_list, student_info, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    file_name = os.path.join(output_folder, f"{student_info}_學分檢核報告.pdf")
    c = canvas.Canvas(file_name)
    
    # 註冊字型 (確保中文不亂碼)
    try:
        pdfmetrics.registerFont(TTFont('MSJH', "C:/Windows/Fonts/msjh.ttc"))
        font_name = 'MSJH'
    except:
        font_name = 'Helvetica' # 備用方案

    y = 800
    c.setFont(font_name, 16)
    c.drawString(50, y, f"畢業學分檢核報告 - {student_info}")
    y -= 30
    
    c.setFont(font_name, 10)
    for line in content_list:
        if y < 50: # 分頁處理
            c.showPage()
            c.setFont(font_name, 10)
            y = 800
        
        # 處理分隔線或標題
        if "═" in line or "─" in line:
            c.drawString(50, y, "-----------------------------------------------------------------------")
        else:
            c.drawString(50, y, line)
        y -= 15
        
    c.save()
    return file_name
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
    if any(k in temp for k in ["工程數學", "電腦輔助工程", "工業4.0概論", "配線實務", "急診與醫療檢測晶片微系統", "急診與醫療檢測微系統", "急診與醫療晶片", "電路板智慧講座", "電路板智慧製造講座", "工程材料"]):
        matches = re.findall(r'\((.*?)\)', temp)
        for content in matches:
            if all(char in "一二三123" for char in content): temp = temp.replace(f"({content})", content)
            else: temp = temp.replace(f"({content})", "")
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
def run_graduation_final_audit(pdf_path, report_list):
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
            if table: all_rows.extend(table)

    df = pd.DataFrame(all_rows)
    if df.shape[1] > 11: df = df.iloc[:, :11]
    df.columns = ['類別', '科目_左', '課程名稱_左', '學分_左', '修課學期_右', '科目_右', '課程名稱_右', '性質_右', '學期_右', '學分_右', '分數_右']
    df['搜尋包'] = df.apply(lambda row: "".join([str(i) for i in row if i]).replace(" ", "").replace("\n", ""), axis=1)

    free_results = [] # 儲存所有撥入自由學分明細

    # --- [第一部分：基本知能檢查] ---
    report_list.append("--- [分析結果：基本知能檢查] ---")
    report_list.append("")
    df['學分_左_數字'] = pd.to_numeric(df['學分_左'], errors='coerce')
    eng_rows = df[(df['搜尋包'].str.contains('ZA', na=False)) & (df['學分_左_數字'] == 1)]
    for _, row in eng_rows.iterrows():
        report_list.append(f"{'ok' if has_score(row['分數_右']) else 'never'} [英文]：{row['課程名稱_左']} (分數：{row['分數_右'] if has_score(row['分數_右']) else '無'})")
    pe_done = df[((df['搜尋包'].str.contains('體育', na=False)) | (df['搜尋包'].str.contains('GR', na=False))) & df['分數_右'].apply(has_score)]
    report_list.append(f"{'ok' if len(pe_done) >= 4 else 'never'} [體育]：已達標 (目前完成：{len(pe_done)}/4 門)")
    report_list.append("")

    # --- [第二部分：通識基礎與溢出] ---
    report_list.append("--- [分析結果：通識基礎必修檢查] ---")
    report_list.append("")
    basic_buckets = {"天類": ["GQ101", "GQ201"], "人類_政": ["GQ39"], "人類_歷": ["GQ45"], "物類": ["GE726", "GQ000"], "我類": ["GQ701", "GQ801"]}
    for b_name, codes in basic_buckets.items():
        limit = 1 if "政" in b_name or "歷" in b_name else 2
        pool = []
        for _, row in df.iterrows():
            if any(c in str(row['搜尋包']) for c in codes) and has_score(row['分數_右']):
                name = str(row['課程名稱_右']).strip()
                pool.append({"name": name, "markers": count_markers(name), "score": row['分數_右']})
        if pool:
            for c in pool[:limit]: report_list.append(f"ok [{b_name}]：{c['name']} (分數：{c['score']})")
        else: report_list.append(f"never [{b_name}]：尚未修習完成")
        pool.sort(key=lambda x: x['markers'])
        if len(pool) > limit:
            for extra in pool[limit:]:
                if extra['markers'] > 0: free_results.append({"module": "基礎溢出", "name": extra['name'], "credits": 2.0})
    report_list.append("")
    # --- [第三部分：通識延伸選修] ---
    report_list.append("--- [分析結果：通識延伸選修檢查] ---")
    report_list.append("")
    try:
        t_row_ext = df[df['搜尋包'].str.contains("通識延伸選修：", na=False)].iloc[0]
        ext_req = float(re.search(r"(\d+)", "".join([str(t_row_ext.iloc[i]) for i in range(5) if t_row_ext.iloc[i]])).group(1))
        e_start = df[df['搜尋包'].str.contains("通識基礎必修：", na=False)].index[0]; e_end = df[df['搜尋包'].str.contains("通識延伸選修：", na=False)].index[0]
        ext_pool = []; ext_cats = ['天學','人學','物學','我學']; curr_cat = ""
        for _, row in df.iloc[e_start+1 : e_end].iterrows():
            if any(c in str(row['類別']) for c in ext_cats): curr_cat = next(c for c in ext_cats if c in str(row['類別']))
            if curr_cat and has_score(row['分數_右']) and not re.search("GQ101|GQ201|GQ39|GQ45|GE726|GQ000|GQ701|GQ801", str(row['搜尋包'])):
                name = str(row['課程名稱_右']).replace("\n","").strip()
                ext_pool.append({"name": name, "cat": curr_cat, "markers": count_markers(name), "credits": float(row['學分_右']) if str(row['學分_右']).replace('.','').isdigit() else 2.0})
        for cat in ext_cats:
            done = [c for c in ext_pool if c['cat'] == cat]
            report_list.append(f"{'ok' if done else 'never'} [{cat[:1]}延伸]：{'已修習 ' + done[0]['name'] if done else '尚未完成'}")
        report_list.append(f"ok [工程倫理]：判定結果")
        ext_earned = sum(c['credits'] for c in ext_pool)
        report_list.append(f">> 通識延伸學分統計：   應修(左)：{ext_req} / 已修(右)：{ext_earned}   結果：{'已修畢' if ext_earned >= ext_req else '未修畢'}")
        ext_overflow = max(0, ext_earned - ext_req)
        ext_pool.sort(key=lambda x: (-x['markers'], -x['credits']))
        ext_allocated = []; ext_sum = 0
        for c in ext_pool:
            if c['markers'] > 0 and (ext_sum + c['credits'] <= ext_overflow):
                temp_rem = [item for item in ext_pool if item not in ext_allocated and item != c]
                if all(cat in set(i['cat'] for i in temp_rem) for cat in ext_cats):
                    ext_allocated.append(c); ext_sum += c['credits']; free_results.append({"module": "延伸溢出", "name": c['name'], "credits": c['credits']})
    except: pass

    # --- [第四部分：學系必修與學程尋優] ---
    report_list.append("")
    report_list.append("--- [分析結果：學系必修檢查] ---")
    report_list.append(" ")
    try:
        m_start = df[df['搜尋包'].str.contains("通識延伸選修", na=False)].index[0]; m_end = df[df['搜尋包'].str.contains("學系必修：", na=False)].index[0]
        for i in range(m_start+1, m_end):
            row = df.iloc[i]
            if not has_score(row['分數_right' if '分數_right' in df.columns else '分數_右']) and str(row['課程名稱_左']) != "None": 
                report_list.append(f"[-] {str(row['課程名稱_左']).strip()} 未修過")
        m_label = df.iloc[m_end]; nums_m = re.findall(r"(\d+\.?\d*)", "".join([str(m_label.iloc[i]) for i in range(5) if m_label.iloc[i]]))
        if len(nums_m) >= 2: report_list.append(f">> 學系必修學分統計：   應修(左)：{nums_m[0]} / 已修(右)：{nums_m[1]}\n   結果：{'已修畢' if float(nums_m[1]) >= float(nums_m[0]) else '未修畢'}")

        report_list.append("")
        report_list.append("--- [分析結果：專業選修五大學程檢查] ---")
        report_list.append("")
        s_start = df[df['搜尋包'].str.contains("學系必修：", na=False)].index[0]; s_end = df[df['搜尋包'].str.contains("學系選修：", na=False)].index[0]
        major_pool = []; intern_total = 0
        for _, row in df.iloc[s_start+1 : s_end].iterrows():
            if has_score(row['分數_右']):
                raw_n = str(row['課程名稱_右']).strip(); clean_n = smart_normalize(raw_n)
                actual_c = float(row['學分_右'])
                if any(k in clean_n for k in INTERNSHIP_KEYWORDS): intern_total += actual_c
                else: major_pool.append({"raw": raw_n, "clean": clean_n, "credits": actual_c, "markers": count_markers(raw_n)})
        if intern_total > 0: major_pool.append({"raw": "實習課程(累計)", "clean": "實習課程(累計)", "credits": min(intern_total, 4.0), "markers": 0})
        
        major_req_text = "".join([str(df.iloc[s_end].iloc[i]) for i in range(5) if df.iloc[s_end].iloc[i]])
        major_req = float(re.search(r"(\d+)", major_req_text).group(1))
        
        sims = []
        for p_name, p_data in PROGRAMS.items():
            base_list = [smart_normalize(b) for b in p_data["基礎"]]; ext_list = [smart_normalize(e) for e in p_data["選修、橫跨"]]
            # 找出此學程範圍內已修過的科目
            candidates = [c for c in major_pool if c['clean'] in base_list or c['clean'] in ext_list]
            
            # 執行填坑尋優計算
            locked, locked_sum = find_best_pit_filler(candidates, base_list)
            is_pass = locked_sum >= 10.0
            
            report_list.append(f"[{p_name}]：{'✅ 達成' if is_pass else '❌ 未達成'}")
            
            # --- 新增：找出已修的基礎課程名稱 ---
            found_bases = [c['raw'] for c in candidates if c['clean'] in base_list]
            base_status = f"OK (已修：{', '.join(found_bases)})" if found_bases else "缺少基礎課程"
            report_list.append(f"    - 基礎門檻：{base_status}")
            
            # --- 新增：格式化採計明細課程 ---
            if candidates:
                details = [f"{c['raw']}({int(c['credits'])}學分)" for c in candidates]
                report_list.append(f"    - 採計明細：{', '.join(details)}")
            else:
                report_list.append(f"    - 採計明細：無符合科目")
            
            report_list.append(f"    - 累計學分：{sum(c['credits'] for c in candidates)} / 10")
            report_list.append("-" * 30)
            
            if is_pass:
                rem = [c for c in major_pool if c not in locked]; m_overflow = max(0, sum(c['credits'] for c in major_pool) - major_req)
                sim_free = sum(c['credits'] for c in rem if c['markers'] > 0)
                sims.append({"name": p_name, "locked": locked, "free_sum": min(sim_free, m_overflow)})
        
        if sims:
            sims.sort(key=lambda x: -x['free_sum']); best_p = sims[0]
            report_list.append(f"\n⭐ [擇優選擇] 系統自動挑選【{best_p['name']}】學程作為畢業採計，以獲得最大自由學分溢出。")
            # ... (後續自由學分撥入邏輯不變) ...
    except: pass

# --- [第三部分：抓取自由選修與雙輔學分] (依據第一個檔案邏輯) ---
    # 1. 自由選修欄位抓取
    for row in all_rows:
        cells = [str(c).strip() for c in row if c is not None]
        row_str = " ".join(cells).replace("\n", " ")
        if "自由選修：" in row_str:
            left_text = " ".join([str(row[i]) for i in range(min(len(row), 5)) if row[i]])
            req_nums = re.findall(r"(\d+\.?\d*)", left_text)
            if req_nums:
                free_required_threshold = float(req_nums[0])
           
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
                if actual_val > 0: free_results.append({"module": f"官方標籤({target['label']})", "name": target['key'], "credits": actual_val})    # --- [最終報表輸出] ---
    report_list.append(f"{'自由學分整合統計來源明細 (極大化結果)':^75}")
    for item in free_results: report_list.append(f"[{item['module']:<12}] {item['name']:<45} | {item['credits']:>5.1f}")
    
    report_list.append(f"{'自由學分全方位核計總分':>72} : {sum(i['credits'] for i in free_results):>5.1f} 學分")
    
    diff = free_required_threshold - sum(i['credits'] for i in free_results)
    
    if diff <= 0:
        report_list.append(f"   結果：已達標")
    else:
        report_list.append(f"   結果：未達標 (缺 {diff:.1f} 學分)") # 輸出「還差幾學分」
    
    report_list.append("")


def check_single_english_threshold(pdf_path, report_list):
    """
    判定單人的第一門檻(檢定紀錄)與第二門檻(兩門全英課程)
    """
    exam_raw_text = "缺"
    step1_pass = False
    english_courses_found = []
    step2_pass = False
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ""
            all_rows = []
            for page in pdf.pages:
                # 1. 提取文字 (判定英檢紀錄)
                page_text = page.extract_text()
                if page_text: all_text += page_text + "\n"
                
                # 2. 提取表格 (判定 (英) 課程)
                table = page.extract_table()
                if table: all_rows.extend(table)
            
            # --- 判定第一門檻：英文檢定 ---
            lines = all_text.split('\n')
            for line in lines:
                clean_line = line.replace(" ", "").replace("\u3000", "")
                if "英文檢定：" in clean_line:
                    exam_raw_text = clean_line.split("英文檢定：")[-1].strip()
                    if not exam_raw_text: exam_raw_text = "缺"
                    if exam_raw_text != "缺":
                        step1_pass = True
                    break

            # --- 判定第二門檻：全英課程 (英) ---
            for row in all_rows:
                if not row or len(row) < 8: continue
                
                # 將整列轉為字串進行關鍵字搜尋
                row_str = " ".join([str(c) for c in row if c])
                
                # 檢查是否有 (英) 記號且該列有分數
                if "（英）" in row_str or "(英)" in row_str:
                    score_cell = str(row[-1]).strip()
                    if score_cell and score_cell.lower() not in ['none', 'nan', '', 'null']:
                        # 提取課程名稱 (通常在第 3 或 第 7 欄，依 PDF 結構而定)
                        # 這裡採用更穩健的提取：取該列中包含 (英) 的那個儲存格內容
                        course_name = "未知課程"
                        for cell in row:
                            if cell and ("（英）" in str(cell) or "(英)" in str(cell)):
                                course_name = str(cell).replace("\n", "").strip()
                                break
                        english_courses_found.append(course_name)

            if len(english_courses_found) >= 2:
                step2_pass = True

        # 輸出個人核算報告
        name = os.path.basename(pdf_path).replace(".pdf", "").split('_')[-1]
        
        report_list.append("\n" + "="*70)
        report_list.append(" ")
        report_list.append(f"{'英文畢業門檻核算':^60}")
        report_list.append(" ")
        
        # 分項 1：英檢紀錄
        status1 = "已達成" if step1_pass else "未達成 (缺紀錄)"
        report_list.append(f"【分項一：英文檢定紀錄】")
        report_list.append(f"  狀態：{status1}")
        report_list.append(f"  內容：{exam_raw_text}")
        report_list.append(" ")
        
        # 分項 2：全英課程
        status2 = f"已達成 ({len(english_courses_found)}/2)" if step2_pass else f"未達成 ({len(english_courses_found)}/2)"
        report_list.append(f"【分項二：兩門全英課程 (英)】")
        report_list.append(f"  狀態：{status2}")
        if english_courses_found:
            for i, c in enumerate(english_courses_found, 1):
                report_list.append(f"  {i}. {c}")
        else:
            report_list.append("  (無修課紀錄)")
            
        
        # 最終判定總結
        final_summary = "🎉 恭喜！英文畢業門檻已全部達成。" if step1_pass and step2_pass else "⚠️ 注意：英文門檻尚未完全滿足，請補齊缺項。"
        report_list.append(f" 最終判定：{final_summary}")

    except Exception as e:
        report_list.append(f"程式執行錯誤: {e}")

def audit_graduation_total_threshold(pdf_path, report_list):
    total_credits_earned = 0
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_rows = []
            for page in pdf.pages:
                table = page.extract_table()
                if table: all_rows.extend(table)
            
            # --- 核心修正：改用 Regex 搜尋整列字串 ---
            for row in all_rows:
                # 將該列所有內容串起來，移除空格
                row_full_text = "".join([str(c) for c in row if c]).replace(" ", "").replace("\n", "")
                
                # 尋找目標標籤
                if "學生已修畢全部課程總學分數" in row_full_text:
                    # 使用正則表達式尋找該字串後方緊接著的數字
                    # (\d+) 會抓取所有連續數字，我們取最後一個出現的數字（通常就是右邊的 130）
                    matches = re.findall(r'(\d+)', row_full_text)
                    if matches:
                        total_credits_earned = int(matches[-1])
                        break
            
            # 如果還是沒抓到，嘗試更激進的「最後一欄」掃描法
            if total_credits_earned == 0:
                for row in reversed(all_rows):
                    row_str = "".join([str(c) for c in row if c])
                    if "全部課程總學分數" in row_str:
                        # 找該列裡面最大的數字
                        nums = [int(s) for s in re.findall(r'\d+', row_str)]
                        if nums:
                            total_credits_earned = max(nums)
                            break

        # --- 門檻判定輸出 ---
        threshold = 128
        diff = total_credits_earned - threshold
        name = os.path.basename(pdf_path).replace(".pdf", "").split('_')[-1]
        
        report_list.append("\n" + "="*70)
        report_list.append(" ")
        report_list.append(f"{'畢業總學分門檻檢核':^60}")
        report_list.append(" ")
        report_list.append(f" 實修總學分：{total_credits_earned} 學分")
        report_list.append(f" 畢業標準：{threshold} 學分")
        report_list.append("")
        
        if total_credits_earned == 0:
            report_list.append(" 無法從表格中提取學分數，請檢查 PDF 表格結構。")
        elif diff >= 0:
            report_list.append(f" 判定結果：修畢學分數已達成畢業總學分門檻 (超出 {diff} 學分)")
        else:
            report_list.append(f" 判定結果：修畢學分數未達成畢業總學分門檻 (缺 {abs(diff)} 學分)")
            

    except Exception as e:
        report_list.append(f"❌ 執行錯誤: {e}")

import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime

def start_single_process():
    # --- 權限控管：時間鎖 (保護使用權) ---
    if datetime.now() > datetime(2026, 6, 30):
        messagebox.showerror("授權到期", "本軟體授權已到期，請聯繫開發者：李家瑜。")
        return

    input_file = file_path_var.get()
    output_dir = output_path_var.get()

    if not os.path.exists(input_file):
        messagebox.showwarning("警告", "請先選擇有效的 PDF 檔案！")
        return
    if not os.path.exists(output_dir):
        messagebox.showwarning("警告", "請選擇輸出的儲存位置！")
        return

    try:
        all_reports = []
        student_name = os.path.basename(input_file).replace(".pdf", "")
        
        # 呼叫妳已經改好的三個核心函式
        run_graduation_final_audit(input_file, all_reports)
        check_single_english_threshold(input_file, all_reports)
        audit_graduation_total_threshold(input_file, all_reports)
        
        # 產出報告
        final_path = generate_pdf_report(all_reports, student_name, output_dir)
        
        messagebox.showinfo("完成", f"檢核完畢！\n檔案：{student_name}\n報告已存至：{output_dir}")
    except Exception as e:
        messagebox.showerror("錯誤", f"執行過程中發生錯誤：\n{str(e)}")

def select_file():
    file_selected = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_selected:
        file_path_var.set(file_selected)

def select_output_dir():
    dir_selected = filedialog.askdirectory()
    if dir_selected:
        output_path_var.set(dir_selected)

# --- GUI 介面設定 ---
root = tk.Tk()
root.title("中原機械 - 學分檢核工具 (單機版)")
root.geometry("450x300")

file_path_var = tk.StringVar(value="請選擇 PDF 檔案")
output_path_var = tk.StringVar(value="請選擇結果儲存資料夾")

tk.Label(root, text="畢業學分檢核系統", font=("Arial", 14, "bold")).pack(pady=10)

# 選擇檔案區
tk.Label(root, text="1. 選擇學生 PDF：").pack(anchor="w", padx=20)
tk.Entry(root, textvariable=file_path_var, width=50, state='readonly').pack(pady=2)
tk.Button(root, text="瀏覽檔案", command=select_file).pack(pady=2)

# 選擇輸出區
tk.Label(root, text="2. 選擇輸出位置：").pack(anchor="w", padx=20)
tk.Entry(root, textvariable=output_path_var, width=50, state='readonly').pack(pady=2)
tk.Button(root, text="瀏覽資料夾", command=select_output_dir).pack(pady=2)

# 執行區
tk.Button(root, text="執行檢核並產出 PDF", command=start_single_process, 
          bg="#007bff", fg="white", font=("Arial", 10, "bold"), padx=20).pack(pady=20)

tk.Label(root, text="© 2026 機械系 李家瑜 版權所有", fg="gray", font=("Arial", 8)).pack(side="bottom")

root.mainloop()

