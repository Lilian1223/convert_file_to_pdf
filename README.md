# Guide: Turn `convert_by_word.py` into a Portable Desktop App (.exe)

This guide will show you how to set up, install the necessary tools, and package the `convert_by_word.py` script into a single executable (`.exe`) file that you can easily use on any Windows PC.

---

## Prerequisite: Install Python & Add to PATH

If you don't have Python installed, or if your Command Prompt doesn't recognize Python commands:

1. Download the latest Python installer from the official website:
   https://www.python.org/downloads/
2. **Crucial Step:** Double-click the installer, and at the very bottom of the setup window, check the box that says **"Add python.exe to PATH"** (or "Add Python to PATH") before clicking "Install Now".

---

## Step 1: Download & Place the Script

1. Copy the code for `convert_by_word.py`.
2. Save or move the `convert_by_word.py` file directly onto your **Windows Desktop**.
   *(Placing it on the desktop makes it much easier to target in the following steps!)*

---

## Step 2: Open Command Prompt (CMD) & Navigate to Desktop

We need to tell your computer's terminal to look at your Desktop folder where the script is located.

1. Press the **`Win + R`** keys on your keyboard, type **`cmd`**, and press **Enter** to open the Command Prompt (the black window).
2. Type the following command to switch your directory to the Desktop, then press **Enter**:
   ```cmd
   cd Desktop
   ```
   *(You should now see `C:\Users\YOUR_NAME\Desktop>` on the left side of the cursor).*

---

## Step 3: Install the Required Tools & Libraries

Before packaging, we need to install the libraries that allow Python to talk to Microsoft Word, as well as the packaging tool (`PyInstaller`).

Type the following commands in the CMD window one by one, pressing **Enter** after each:

```cmd
py -m pip install pywin32
```

```cmd
py -m pip install pyinstaller
```

*(Note: If the `py` command doesn't work, replace `py` with `python` in the commands above).*

---

## Step 4: Generate the Executable (.exe) File

Now, run the packaging tool to compile your script. In the same CMD window, type this command and press **Enter**:

```cmd
py -m PyInstaller --onefile convert_by_word.py
```

* The packaging process will start. You will see a lot of text scrolling by.
* Wait until you see a success message at the end, such as: 
  `Building EXE from EXE-00.toc completed successfully.`

---

## Step 5: Where to Find Your App

Once the process is complete, look at your **Desktop**. You will notice a few new folders:

* 📂 **`dist` Folder (The Important One):** Double-click to open it. Inside, you will find your compiled **`convert_by_word.exe`** app!
* 📂 **`build` Folder & `.spec` File:** These are temporary files used during compilation. You can safely delete them from your desktop.

---

## How to Use Your New Tool

1. Drag **`convert_by_word.exe`** out of the `dist` folder and place it directly on your desktop.
2. Drag and drop any Microsoft Word file (`.docx` or `.doc`) onto the **`convert_by_word.exe`** icon.
3. A black window will pop up showing the conversion progress, and a perfect, beautifully formatted PDF with **zero font glitches** will be created right next to your original file!
   
