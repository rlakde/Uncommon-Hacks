Dim args, objExcel
Set args = WScript.Arguments
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False
objExcel.Workbooks.Open "C:\Users\rohit\Desktop\UncommonHacks\Book1.xltm"
objExcel.visible = True
objExcel.Run "helloworld"
objExcel.Activeworkbook.SaveAs "C:\Users\rohit\Desktop\UncommonHacks\" & Date &".xls"
objExcel.ActiveWorkbook.close "C:\Users\rohit\Desktop\UncommonHacks\Book1.xltm"
objExcel.Quit