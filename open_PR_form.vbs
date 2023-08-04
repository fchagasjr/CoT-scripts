'Save Excel as PDF'

Dim Excel
Dim ExcelDoc
Dim ExcelSheet

ExcelFile = WScript.Arguments(0)

'Opens the Excel file'
Set Excel = CreateObject("Excel.Application")
Set ExcelDoc = Excel.Workbooks.open(ExcelFile)

Excel.Visible = True

Set ExcelSheet = ExcelDoc.Sheets("Purchase Order request form")

'Auto Fill Date'
ExcelSheet.Cells(5,8).Value = Date()

'Auto Fill Company'
ExcelSheet.Cells(7,3).Value = WScript.Arguments(1)

'Auto Fill Description'
ExcelSheet.Cells(17,4).Value = WScript.Arguments(2)

'Cache out Excel files'
MsgBox("Update the purchase request form, save it and click [OK]")
ExcelDoc.Saved = True
ExcelDoc.Close
Excel = Null