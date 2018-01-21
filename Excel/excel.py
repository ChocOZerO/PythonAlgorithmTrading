import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('d:\\IT\\ProgrammingStudy\\PythonAlgorithmTrading\\Excel\\test.xlsx')
ws = wb.ActiveSheet
ws.Cells(2, 1).Value = "Python"
ws.Cells(2, 2).Value = "is"
ws.Range("C2").Value = "good"
ws.Range("C2").Interior.ColorIndex = 10
excel.Quit()
