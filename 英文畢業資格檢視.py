import pdfplumber
import os
import re

# ==========================================
# 1. 設定單一檔案路徑
# ==========================================
pdf_file = "list_two/11123255_黃義凱.pdf" 

def check_single_english_threshold(pdf_path):
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
        
        print("\n" + "="*70)
        print(f"{'英文畢業門檻個人核算報告':^60}")
        print("="*70)
        print(f" 學生姓名：{name}")
        print("-" * 70)
        
        # 分項 1：英檢紀錄
        status1 = "✅ 已達成" if step1_pass else "❌ 未達成 (缺紀錄)"
        print(f"【分項一：英文檢定紀錄】")
        print(f"  狀態：{status1}")
        print(f"  內容：{exam_raw_text}")
        
        print("-" * 70)
        
        # 分項 2：全英課程
        status2 = f"✅ 已達成 ({len(english_courses_found)}/2)" if step2_pass else f"❌ 未達成 ({len(english_courses_found)}/2)"
        print(f"【分項二：兩門全英課程 (英)】")
        print(f"  狀態：{status2}")
        if english_courses_found:
            for i, c in enumerate(english_courses_found, 1):
                print(f"  {i}. {c}")
        else:
            print("  (無修課紀錄)")
            
        print("-" * 70)
        
        # 最終判定總結
        final_summary = "🎉 恭喜！英文畢業門檻已全部達成。" if step1_pass and step2_pass else "⚠️ 注意：英文門檻尚未完全滿足，請補齊缺項。"
        print(f" 最終判定：{final_summary}")
        print("="*70 + "\n")

    except Exception as e:
        print(f"❌ 程式執行錯誤: {e}")

# ==========================================
# 3. 執行執行
# ==========================================
if __name__ == "__main__":
    if os.path.exists(pdf_file):
        check_single_english_threshold(pdf_file)
    else:
        print(f"❌ 找不到檔案：{pdf_file}")