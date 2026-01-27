import os
import subprocess
import sys

def build_exe():
    print("===================================")
    print("PGT-A Report Generator - EXE Builder")
    print("===================================")
    
    # Install PyInstaller if missing
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # Define the command
    # --onefile: Create a single executable
    # --windowed: No console window (since it's a GUI app)
    # --add-data: Include assets and styles
    # Adjust asset path syntax for Windows/Linux compatibility in build script
    
    assets_sep = ";" if os.name == 'nt' else ":"
    
    cmd = [
        "pyinstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        f"--add-data=pgta_styles.css{assets_sep}.",
        f"--add-data=assets{assets_sep}assets",
        f"--add-data=core{assets_sep}core",
        "--name=PGTA_Report_Generator",
        "pgta_report_generator.py"
    ]

    print(f"Running: {' '.join(cmd)}")
    subprocess.check_call(cmd)

    print("\n===================================")
    print("Build Complete!")
    print("Your executable is in the 'dist' folder.")
    print("===================================")

if __name__ == "__main__":
    build_exe()
