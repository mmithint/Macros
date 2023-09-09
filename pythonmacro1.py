import os
import win32com.client as win32
import xlsxwriter
# def run_excel_macro():
#     # "D:\\home\\ProcessingUnit\\TempSingleFile\\helloworld.xlsm
#     print("0"*40)
#     xl = win32.Dispatch('Excel.Application')
#     print("1"*40)
#     xl.Application.visible = False
#     file_path = "D:\\home\\ProcessingUnit\\TempSingleFile\\resultBook.xlsx"
#     print("2"*40)
#     separator_char = os.sep
#     try:
#         wb = xl.Workbooks.Open(os.path.abspath(file_path))
#         print("3"*40)
#         xl.Application.run("helloworld.xlsm!Module1.helloworld")
#         print("4"*40)
#         wb.Save()
#         wb.Close()
#         print("5"*40)
#         xl.Application.Quit()
#         del xl
#     except Exception as ex:
#         print("6"*40)
#         template = "An exception of type {0} occurred. Arguments:\n{1!r}"
#         message = template.format(type(ex).__name__, ex.args)
#         print(message)
    

# run_excel_macro()
if os.path.exists("C:\\Users\\Admin\\Downloads\\testing.xlsm"):
    print("0"*40)
    local_path = "D:\\home\\ProcessingUnit\\TempSingleFile"
    filename1 = 'output.xlsx'
    path = local_path + "\\" +filename1
    wb = xlsxwriter.Workbook(path)
    wb.close()
    print("1"*40)
    xl = win32.Dispatch('Excel.Application')
    print("2"*40)
    wb = xl.Workbooks.Open(os.path.abspath(path))
    print("3"*40)
    xl.Application.run("testing.xlsm!Module1.Macro2")
    print("4"*40)
    wb.Save()
    wb.Close()
    xl.Application.Quit()
    del xl

