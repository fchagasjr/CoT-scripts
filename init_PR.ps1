Add-Type -AssemblyName PresentationFramework

# Setup variables
while (!$supplier -or !$description) { 
$supplier, $description = .\PR_form.ps1

  if (!$supplier -or !$description) {
   [System.Windows.MessageBox]::Show("Please select supplier and provide description", "Error", "OK", "None")
  }
}
$drive = "$pwd".substring(0,2)
$creation_date = date -format yyyyMMdd
$month = $creation_date.substring(4,2)
$day = $creation_date.substring(6,2)
$year = $creation_date.substring(0,4)
$order_info = "PR $creation_date - $description"

.$drive\Scripts\new_folder.ps1 "$drive\Documents\Purchase Orders"
.$drive\Scripts\new_folder.ps1 "$year"
.$drive\Scripts\new_folder.ps1 "$supplier"
.$drive\Scripts\new_folder.ps1 "$order_info"

mkdir "Receipts"

explorer .

[System.Windows.MessageBox]::Show("Drag and drop quote file to new directory", "Rename Quote file", "OK", "None")


move *.pdf "QT $creation_date.pdf"
".\QT $creation_date.pdf"

$pr_form_path = "$drive\Scripts\misc\PR_form.xlsx"
cscript.exe $drive\Scripts\open_PR_form.vbs $pr_form_path "$supplier" "$description"
[System.Windows.MessageBox]::Show("Update the purchase request form and save it", "Create Request Form", "OK", "None")
cscript.exe $drive\Scripts\xlsx_to_pdf.vbs "$drive\Scripts\misc\PR_form.xlsx" "$pwd\PR $creation_date.pdf"

[System.Windows.MessageBox]::Show("$order_info was created successfully!", "Request Created", "OK", "None")

<#
:focus_directory
IF NOT EXIST "\%~1" mkdir %1
cd %1
EXIT /B

:confirm
call :script_head
echo %1 was selected as %~2.
SET /p input="Press [ENTER] to continue."
EXIT /B

:rename_quote
CALL :script_head
echo Move quote file to folder.
SET /p input="Press [ENTER] to continue."
move *.pdf "QT %creation_date%.pdf"
start "" /max "QT %creation_date%.pdf"
EXIT /B

:create_pr
CALL :script_head
$pr_form_path=$drive/Scripts/misc/PR_form.xlsx
cscript.exe $drive/Scripts/open_PR_form.vbs $pr_form_path "$supplier" "$description"
echo Update the purchase request sheet and save it.
SET /p input="Press [ENTER] to continue."
cscript.exe %drive%/Scripts/xlsx_to_pdf.vbs "$drive\Scripts\misc\PR_form.xlsx" "$pwd\PR $creation_date.pdf"
EXIT /B

:exit
::CALL :script_head
SET /p input="%company% - %order_info% was created. Press [ENTER] to close."
GOTO :EOF
#>