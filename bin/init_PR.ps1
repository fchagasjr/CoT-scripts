Add-Type -AssemblyName PresentationFramework

# Collect Info
$personal_drive = "$pwd".substring(0,2)
$scripts_folder = "$personal_drive\Scripts\bin"
$temp_folder = "$scripts_folder\TEMP"

mkdir "$temp_folder"

$supplier, $description, $documents = powershell.exe $scripts_folder\PR_form.ps1 "$temp_folder"

 if (!$supplier -or !$description -or !$documents) {
	if ($supplier -ne "Cancel") {
		[System.Windows.MessageBox]::Show("Please select supplier, provide description and upload documents", "Error", "OK", "None")
	}
	rm -r "$temp_folder"
	exit
 }
 
# Setup variables

$creation_date = date -format yyyyMMdd
$month = $creation_date.substring(4,2)
$day = $creation_date.substring(6,2)
$year = $creation_date.substring(0,4)
$order_info = "PR $creation_date - $description"

# Create folders

$new_folder = "$personal_drive\Documents\Purchase Orders\$year\$supplier\$order_info"

powershell.exe $scripts_folder\new_folder.ps1 "$new_folder"

mkdir "$new_folder\Receipts"

# Copy documents to folder

foreach ($document in $documents) {
 Copy-Item "$document" -Destination "$new_folder"
}

Invoke-Item "$new_folder\*.pdf"

rm -r "$temp_folder"

# Create the request form

$pr_form_path = "$personal_drive\Scripts\misc\PR_form.xlsx"
cscript.exe $scripts_folder\open_PR_form.vbs $pr_form_path "$supplier" "$description"
cscript.exe $scripts_folder\xlsx_to_pdf.vbs "$scripts_folder\..\misc\PR_form.xlsx" "$new_folder\PR $creation_date.pdf"

[System.Windows.MessageBox]::Show("$order_info was created successfully!", "Request Created", "OK", "None")

explorer "$new_folder"