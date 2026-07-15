# How to use

## If you don't have Python on your device...

1. Download and install Python:
   https://www.python.org/downloads/
   *(Make sure to check "Add Python to PATH" during installation)*

2. Open cmd and verify installation:
   ```cmd
   python --version
   ```
## If you already have Python, then...

1. Open cmd and install the required libraries:
   ```cmd
   pip install python-docx openpyxl python-pptx img2pdf reportlab pillow
   ```
(If pip doesn't work, try py -m pip install ...)

3. Add shortcut

Put the Python file on the desktop

Make your own choice：

1: A completion message will pop up.
Right-click on the desktop → New → Shortcut, enter the target field：<br>
pythonw.exe "C:\Users\YOUR USER NAME\Desktop\convert_to_pdf_nolibre.py" --notify<br>
After you drag the file up, a small window will pop up telling you which ones succeeded and which ones failed.

2：Complete silence, do not jump anything<br>
Right-click on the desktop → New → Shortcut, enter the target field：<br>
pythonw.exe "C:\Users\YOUR USER NAME\Desktop\convert_to_pdf_nolibre.py" --silent
