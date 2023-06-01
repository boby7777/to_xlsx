import os
import sys
import win32com.client as win32

def convert_xls_to_xlsx(input_file, output_file):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(os.path.abspath(input_file))

    # Check if the output file already exists
    if os.path.isfile(output_file):
        while True:
            response = input("The output file already exists. Do you want to replace it? (Y/N): ")
            if response.upper() == 'Y':
                os.remove(output_file)  # Delete the existing file
                break
            elif response.upper() == 'N':
                print("Conversion aborted.")
                return
            else:
                print("Invalid input. Please enter Y or N.")

    wb.SaveAs(os.path.abspath(output_file), FileFormat=51)  # FileFormat=51 represents .xlsx format
    wb.Close()
    excel.Quit()

# 執行檔案轉換
if len(sys.argv) == 3:
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    convert_xls_to_xlsx(input_file, output_file)
else:
    print("請提供輸入檔案和輸出檔案的路徑。")
    print("範例: python convert_xls_to_xlsx.py input.xls output.xlsx")
