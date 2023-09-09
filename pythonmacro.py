import os, os.path
import win32com.client

if os.path.exists("C:/Users/Admin/Downloads/testing.xlsm"):
    print(os.path)
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(os.path.abspath("C:/Users/Admin/Downloads/testing.xlsm"))
    xl.Application.Run("Module1.Macro2")
    #xl.Application.Save("C:/Users/Admin/Downloads/testing.xlsm")
    # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    xl.Application.Quit() # Comment this out if your excel script closes
    del xl