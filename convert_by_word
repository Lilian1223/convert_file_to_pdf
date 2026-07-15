import os
import sys
import ctypes
import time

# 註冊 Windows 編碼環境，保護拖曳中文字元路徑
if sys.platform == 'win32':
    try:
        import locale
        if sys.stdout is not None:
            sys.stdout.reconfigure(encoding='utf-8')
        if sys.stderr is not None:
            sys.stderr.reconfigure(encoding='utf-8')
        ctypes.windll.kernel32.SetConsoleCP(65001)
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    except Exception:
        pass

def convert_to_pdf_via_word(docx_path):
    """
    調用微軟 Word 引擎轉換 docx 為 pdf，保證 100% 不跑格、無亂碼
    """
    import win32com.client
    
    # 取得絕對路徑（Word API 必須要用絕對路徑才不易出錯）
    abs_docx_path = os.path.abspath(docx_path)
    base_name, _ = os.path.splitext(abs_docx_path)
    pdf_path = base_name + ".pdf"
    
    word = None
    doc = None
    try:
        print(f"\n正在調用 Word 引擎轉換: {os.path.basename(docx_path)}...")
        
        # 啟動背景 Word 程序
        # CoInitialize 預防多線程問題
        import pythoncom
        pythoncom.CoInitialize()
        
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False  # 不顯示 Word 視窗
        word.DisplayAlerts = 0  # 關閉 Word 的警告彈窗（如：巨集警告、唯讀提示）
        
        # 開啟 Word 檔
        doc = word.Documents.Open(abs_docx_path, ReadOnly=1)
        
        # Word 儲存為 PDF 的格式碼是 17
        wdFormatPDF = 17
        doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
        
        print(f"🎉 轉換成功！ PDF 已儲存至: {pdf_path}")
        return True
    except Exception as e:
        print(f"❌ 轉換失敗，錯誤原因: {e}")
        return False
    finally:
        # 確保不管成功或失敗，都會把背景 Word 程式與檔案關閉，釋放記憶體
        if doc is not None:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

def main():
    # 支援拖曳多個檔案
    files = sys.argv[1:]
    
    if not files:
        print("="*60)
        print("💡 使用說明：")
        print("請直接將一個或多個 Word 檔案 (.docx / .doc) 拖曳到本程式圖示上！")
        print("="*60)
        input("\n請按 Enter 鍵結束...")
        return

    print("="*60)
    print("🚀 微軟 Word 官方引擎 PDF 轉檔工具 🚀")
    print("="*60)
    
    success_count = 0
    fail_count = 0
    
    for file_path in files:
        if not os.path.exists(file_path):
            continue
            
        _, ext = os.path.splitext(file_path.lower())
        if ext in ['.docx', '.doc']:
            success = convert_to_pdf_via_word(file_path)
            if success:
                success_count += 1
            else:
                fail_count += 1
        else:
            print(f"⚠️ 跳過不支援的檔案格式: {os.path.basename(file_path)} (僅支援 .docx 與 .doc)")
            fail_count += 1
            
    print("\n" + "="*60)
    print(f"📊 轉檔任務結束！成功: {success_count} 個 | 失敗: {fail_count} 個")
    print("="*60)
    
    # 倒數 3 秒後自動關閉視窗，方便連續工作
    for i in range(3, 0, -1):
        print(f"\r視窗將在 {i} 秒後自動關閉...", end="", flush=True)
        time.sleep(1)

if __name__ == "__main__":
    main()
