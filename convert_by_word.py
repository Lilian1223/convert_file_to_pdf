import os
import sys
import ctypes
import time

# Register a Windows encoding environment to protect dragged Chinese character paths.
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
    It uses the Microsoft Word engine to convert docx files to pdf, guaranteeing 100% no formatting issues and no garbled characters.
    """
    import win32com.client
    
    # Obtain the absolute path (the Word API requires absolute paths to avoid errors).
    abs_docx_path = os.path.abspath(docx_path)
    base_name, _ = os.path.splitext(abs_docx_path)
    pdf_path = base_name + ".pdf"
    
    word = None
    doc = None
    try:
        print(f"\nCalling the Word engine for conversion: {os.path.basename(docx_path)}...")
        
        # Startup background: Word program
        # CoInitialize prevents multithreading issues.
        import pythoncom
        pythoncom.CoInitialize()
        
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False  # Word window not displayed
        word.DisplayAlerts = 0  # Disable Word's warning pop-ups (such as macro warnings and read-only warnings).
        
        # Open Word document
        doc = word.Documents.Open(abs_docx_path, ReadOnly=1)
        
        # The format code for saving a Word document as a PDF is 17.
        wdFormatPDF = 17
        doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
        
        print(f"🎉 Conversion successful! PDF saved to: {pdf_path}")
        return True
    except Exception as e:
        print(f"❌ Conversion failed, error reason: {e}")
        return False
    finally:
        # Ensure that the background Word program and files are closed regardless of success or failure to free up memory.
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
    # Supports dragging and dropping multiple files
    files = sys.argv[1:]
    
    if not files:
        print("="*60)
        print("💡 Instructions for use：")
        print("Please drag and drop one or more Word files (.docx / .doc) directly onto this program icon!")
        print("="*60)
        input("\nPlease press Enter to finish....")
        return

    print("="*60)
    print("🚀 Microsoft Word official engine PDF conversion tool 🚀")
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
            print(f"⚠️ Skip unsupported file formats: {os.path.basename(file_path)} (Supported only .docx 與 .doc)")
            fail_count += 1
            
    print("\n" + "="*60)
    print(f"📊 File transfer mission complete! Success!: {success_count} items | Fail: {fail_count} items")
    print("="*60)
    
    # The window will automatically close after a 3-second countdown for convenient continuous work.
    for i in range(3, 0, -1):
        print(f"\rThe window will close automatically in {i} seconds...", end="", flush=True)
        time.sleep(1)

if __name__ == "__main__":
    main()
