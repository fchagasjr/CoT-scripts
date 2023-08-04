Add-Type -AssemblyName PresentationFramework

# Collect Info

mkdir TEMP

$supplier, $description, $documents = .\PR_form.ps1 TEMP

 if (!$supplier -or !$description -or !$documents) {
	[System.Windows.MessageBox]::Show("Please select supplier, provide description and upload documents", "Error", "OK", "None")
	rm -r TEMP
	exit
 }
 
# Setup variables

$drive = "$pwd".substring(0,2)
$creation_date = date -format yyyyMMdd
$month = $creation_date.substring(4,2)
$day = $creation_date.substring(6,2)
$year = $creation_date.substring(0,4)
$order_info = "PR $creation_date - $description"

# Create folders

$new_folder = "$drive\Documents\Purchase Orders\$year\$supplier\$order_info"

.$drive\Scripts\new_folder.ps1 "$new_folder"

mkdir "Receipts"

# Copy documents to folder

foreach ($document in $documents) {
 Copy-Item "$document" -Destination "$new_folder"
}

Invoke-Item *.pdf

rm -r "$drive\Scripts\TEMP"

# Create the request form

$pr_form_path = "$drive\Scripts\misc\PR_form.xlsx"
cscript.exe $drive\Scripts\open_PR_form.vbs $pr_form_path "$supplier" "$description"
cscript.exe $drive\Scripts\xlsx_to_pdf.vbs "$drive\Scripts\misc\PR_form.xlsx" "$pwd\PR $creation_date.pdf"

[System.Windows.MessageBox]::Show("$order_info was created successfully!", "Request Created", "OK", "None")

explorer .