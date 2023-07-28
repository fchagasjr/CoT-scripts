'Save Excel as PDF'

Dim Excel
Dim ExcelDoc

ExcelFile = "H:\Scripts\misc\PR_form.xlsx"
PdfFile = WScript.Arguments(0)

'Opens the Excel file'
Set Excel = CreateObject("Excel.Application")
Set ExcelDoc = Excel.Workbooks.open(ExcelFile)

'Creates the pdf file'
Excel.ActiveSheet.ExportAsFixedFormat 0, PdfFile ,0, 1, 0,,,0

'Closes the Excel file'
Excel.ActiveWorkbook.Close
Excel.Application.Quit