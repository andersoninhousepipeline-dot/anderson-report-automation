# Using PGT-A Report Generator on Windows (Native)

Since you need a native Windows experience, I have provided two ways to run the software without WSL.

## Option 1: Run via Batch Script (Fastest)
This is the easiest way if you have Python installed.

1.  **Install Python**: Download from [python.org](https://www.python.org/) and ensure you check **"Add Python to PATH"** during installation.
2.  **Double-click `launch.bat`**: This will automatically check for dependencies, install them if missing, and start the app.

---

## Option 2: Build a Standalone `.exe` (Best for distribution)
If you want a single file that you can move around or give to others, you can create a Windows `.exe`.

1.  **Open Command Prompt** in the project folder.
2.  **Run the build script**:
    ```cmd
    python build_windows_exe.py
    ```
3.  **Find your app**: Once finished, a new folder called `dist` will appear. Inside, you will find `PGTA_Report_Generator.exe`.
4.  **Important**: Keep the `wkhtmltopdf` installer handy, as the PDF engine still requires it to be installed on the system (see below).

---

## Vital Requirements for Both Options

### 1. Assets
Ensure the `assets/` folder and `pgta_styles.css` are in the same directory as the script/executable when running.

## Troubleshooting
- **"Python was not found" error**: 
  1. This means Python is not in your system's "PATH". 
  2. To fix: Uninstall Python and reinstall it, but **crucially** check the box that says **"Add Python to PATH"** in the first step of the installer.
  3. Alternatively, you can search for "Edit the system environment variables" in your start menu, click "Environment Variables", find "Path" under User Variables, and add the folder where you installed Python.
- **"PyQt6 failed to install" (Nuclear Option)**: 
  1. Open Command Prompt as **Administrator**.
  2. Run this "Clean Install" command:
     `python -m pip uninstall PyQt6 PyQt6-Qt6 PyQt6-sip -y && python -m pip install PyQt6 --user --force-reinstall`
  3. If it still fails, download and install the **Microsoft Visual C++ Redistributable** (specifically the X64 version) from [here](https://aka.ms/vs/17/release/vc_redist.x64.exe).
- **Black Screen/Not Starting**: If the `.exe` doesn't start, try running `launch.bat` instead to see if any errors are printed to the console.
- **Missing DLLs**: If you get a "PyQt6" error, ensure your Windows has the latest "Microsoft Visual C++ Redistributable" installed.
