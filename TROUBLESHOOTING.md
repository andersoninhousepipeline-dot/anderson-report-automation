# Windows Troubleshooting Guide

## Common Startup Issues

### Issue 1: Application Closes Immediately (1 second)

**Symptoms:**
- Window appears briefly then closes
- No error message shown
- Works fine on Linux

**Causes:**
1. Missing or corrupted PyQt6 installation
2. 32-bit Python (PyQt6 requires 64-bit)
3. Missing Visual C++ Redistributable
4. Corrupted virtual environment

**Solutions:**

1. **Check startup.log**
   ```
   Open the file: startup.log
   Look for the last successful step before crash
   ```

2. **Delete and recreate virtual environment**
   ```
   1. Delete the .venv folder
   2. Run launch.bat again
   3. Wait for dependencies to reinstall
   ```

3. **Verify Python is 64-bit**
   ```
   python -c "import struct; print(struct.calcsize('P') * 8, 'bit')"
   ```
   Should show "64 bit". If it shows "32 bit", uninstall and install 64-bit Python.

4. **Install Visual C++ Redistributable**
   - Download from: https://aka.ms/vs/17/release/vc_redist.x64.exe
   - Install and restart computer

---

### Issue 2: Crashes When Uploading Excel File

**Symptoms:**
- Application starts fine
- Crashes immediately when selecting Excel file
- No error dialog shown

**Causes:**
1. Excel file format issues
2. Missing openpyxl library
3. Encoding problems in Excel file
4. Memory issues with large files

**Solutions:**

1. **Check the error logs**
   ```
   Open: startup.log
   Look for "ERROR in parse_bulk_excel"
   ```

2. **Verify Excel file format**
   - Must have "Details" and "summary" sheets
   - Save as .xlsx (not .xls)
   - Ensure no special characters in sheet names

3. **Test with template file**
   ```
   1. Use the "Download Template" button
   2. Fill in one patient
   3. Try uploading the template
   ```

4. **Check openpyxl installation**
   ```
   .venv\Scripts\python.exe -c "import openpyxl; print('OK')"
   ```

---

### Issue 3: PyQt6 Import Errors

**Symptoms:**
- Error: "PyQt6 is not installed or failed to load"
- Application won't start

**Solutions:**

1. **Reinstall PyQt6**
   ```
   .venv\Scripts\python.exe -m pip uninstall PyQt6 PyQt6-Qt6 -y
   .venv\Scripts\python.exe -m pip install PyQt6>=6.6.0 PyQt6-Qt6>=6.6.0
   ```

2. **Check for conflicting installations**
   ```
   where python
   ```
   Should only show one Python installation

3. **Use setup.bat for clean install**
   ```
   1. Delete .venv folder
   2. Run setup.bat
   3. Wait for completion
   4. Run launch.bat
   ```

---

### Issue 4: Assets Directory Missing

**Symptoms:**
- Error: "Assets directory is missing!"
- Application won't start

**Solutions:**

1. **Verify folder structure**
   ```
   Your folder should contain:
   - pgta_report_generator.py
   - launch.bat
   - assets/
     - pgta/
       - image_page1_0.png
       - image_page1_1.png
       - etc.
   ```

2. **Re-extract from ZIP**
   - Ensure you extracted ALL files
   - Don't move files around

---

## Diagnostic Tools

### Run Diagnostic Script

```
diagnose.bat
```

This will check:
- Python installation
- Virtual environment
- Dependencies
- File structure
- Common issues

Save the output and share if you need support.

### Check Log Files

The application creates several log files:

1. **startup.log** - Detailed application startup log
   - Shows each initialization step
   - Contains full error tracebacks
   - Created by Python application

2. **launcher.log** - Batch file execution log
   - Shows launcher script steps
   - Python detection results
   - Dependency installation status

3. **run_error.log** - Python stderr output
   - Contains Python exceptions
   - Shows import errors
   - Created by launcher script

### Manual Dependency Check

```batch
.venv\Scripts\python.exe -c "import PyQt6; print('PyQt6: OK')"
.venv\Scripts\python.exe -c "import reportlab; print('ReportLab: OK')"
.venv\Scripts\python.exe -c "import pandas; print('Pandas: OK')"
.venv\Scripts\python.exe -c "import openpyxl; print('openpyxl: OK')"
.venv\Scripts\python.exe -c "import docx; print('python-docx: OK')"
```

---

## Getting Help

If none of the above solutions work:

1. Run `diagnose.bat` and save the output
2. Check all three log files:
   - startup.log
   - launcher.log
   - run_error.log
3. Note the exact error message
4. Provide:
   - Windows version
   - Python version
   - When the crash occurs (startup, Excel upload, etc.)
   - Contents of startup.log

---

## Prevention

### Best Practices

1. **Always use 64-bit Python**
2. **Install Visual C++ Redistributable**
3. **Keep Python updated** (3.10 or newer)
4. **Don't modify the .venv folder manually**
5. **Use the provided launcher scripts**

### Regular Maintenance

1. **Update dependencies periodically**
   ```
   .venv\Scripts\python.exe -m pip install --upgrade pip
   .venv\Scripts\python.exe -m pip install --upgrade PyQt6 reportlab pandas openpyxl
   ```

2. **Clean reinstall if issues persist**
   ```
   1. Delete .venv folder
   2. Run launch.bat
   3. Let it reinstall everything
   ```
