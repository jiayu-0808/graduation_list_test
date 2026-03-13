import pdfplumber
import os
import re

# ==========================================
# 1. 設定檔案路徑
# ==========================================
pdf_file = "list_two/11123255_黃義凱.pdf" 

def audit_graduation_total_threshold(pdf_path):
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
        
        print("\n" + "="*70)
        print(f"{'畢業總學分門檻檢核 (修正版)':^60}")
        print("="*70)
        print(f" 學生姓名：{name}")
        print(f" 實修總學分：{total_credits_earned} 學分")
        print(f" 畢業標準：{threshold} 學分")
        print("-" * 70)
        
        if total_credits_earned == 0:
            print(" ⚠️ 警告：無法從表格中提取學分數，請檢查 PDF 表格結構。")
        elif diff >= 0:
            print(f" 判定結果：修畢學分數已達成畢業總學分門檻 (超出 {diff} 學分)")
        else:
            print(f" 判定結果：修畢學分數未達成畢業總學分門檻 (尚缺 {abs(diff)} 學分)")
            
        print("="*70 + "\n")

    except Exception as e:
        print(f"❌ 執行錯誤: {e}")

if __name__ == "__main__":
    if os.path.exists(pdf_file):
        audit_graduation_total_threshold(pdf_file)
    else:
        print(f"❌ 找不到檔案")