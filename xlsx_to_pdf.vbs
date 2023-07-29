'Save Excel as PDF'

Dim Excel
Dim ExcelDoc

ExcelFile = WScript.Arguments(0)
PdfFile = WScript.Arguments(1)

'Opens the Excel file'
Set Excel = CreateObject("Excel.Application")
Set ExcelDoc = Excel.Workbooks.open(ExcelFile)

'Creates the pdf file'
Excel.ActiveSheet.ExportAsFixedFormat 0, PdfFile ,0, 1, 0,,,0

'Closes the Excel file'
Excel.ActiveWorkbook.Close
Excel.Application.Quit