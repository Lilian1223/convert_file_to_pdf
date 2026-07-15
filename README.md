# Guide: Turn `convert_by_word.py` into a Portable Desktop App (.exe)

This guide will show you how to set up, install the necessary tools, and package the `convert_by_word.py` script into a single executable (`.exe`) file that you can easily use on any Windows PC.

---

## Prerequisite: Install Python & Add to PATH

If you don't have Python installed, or if your Command Prompt doesn't recognize Python commands:

1. Download the latest Python installer from the official website:
   https://www.python.org/downloads/
2. **Crucial Step:** Double-click the installer, and at the very bottom of the setup window, check the box that says **"Add python.exe to PATH"** (or "Add Python to PATH") before clicking "Install Now"[cite: 1].

---

## Step 1: Download & Place the Script

1. Copy the code for `convert_by_word.py`.
2. Save/move the `convert_by_word.py` file directly onto your **Windows Desktop**.
   *(Placing it on the desktop makes it much easier to target in the following steps!)*

---

## Step 2: Open Command Prompt (CMD) & Navigate to Desktop

We need to tell your computer's terminal to look at your Desktop folder where the script is located.

1. Press the **`Win + R`** keys on your keyboard, type **`cmd`**, and press **Enter** to open the Command Prompt (the black window).
2. Type the following command to switch your directory to the Desktop, then press **Enter**:
   ```cmd
   cd Desktop
   ```
   
