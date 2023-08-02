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

# Create folders

.$drive\Scripts\new_folder.ps1 "$drive\Documents\Purchase Orders\$year\$supplier\$order_info"

mkdir "Receipts"

explorer .

[System.Windows.MessageBox]::Show("Drag and drop quote file to new directory", "Rename Quote file", "OK", "None")

move *.pdf "QT $creation_date.pdf"
Invoke-Item "QT $creation_date.pdf"

$pr_form_path = "$drive\Scripts\misc\PR_form.xlsx"
cscript.exe $drive\Scripts\open_PR_form.vbs $pr_form_path "$supplier" "$description"
[System.Windows.MessageBox]::Show("Update the purchase request form and save it", "Create Request Form", "OK", "None")
cscript.exe $drive\Scripts\xlsx_to_pdf.vbs "$drive\Scripts\misc\PR_form.xlsx" "$pwd\PR $creation_date.pdf"

[System.Windows.MessageBox]::Show("$order_info was created successfully!", "Request Created", "OK", "None")
