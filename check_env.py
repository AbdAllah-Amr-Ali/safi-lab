import sys
import os

print("Checking dependencies...")
try:
    import PyQt5
    print("✅ PyQt5 installed")
except ImportError:
    print("❌ PyQt5 NOT installed")

try:
    import openpyxl
    print("✅ openpyxl installed")
except ImportError:
    print("❌ openpyxl NOT installed")

try:
    import win32com.client
    print("✅ pywin32 installed")
except ImportError:
    print("❌ pywin32 NOT installed")

try:
    import qrcode
    print("✅ qrcode installed")
except ImportError:
    print("❌ qrcode NOT installed")

print("\nChecking Excel file access...")
excel_path = os.path.abspath("Patients.xlsm")
if os.path.exists(excel_path):
    print(f"✅ Excel file found at: {excel_path}")
    try:
        from openpyxl import load_workbook
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        print("✅ Excel file opened successfully with openpyxl")
        wb.close()
    except Exception as e:
        print(f"❌ Failed to open Excel file: {e}")
else:
    print(f"❌ Excel file NOT found at: {excel_path}")
