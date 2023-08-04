Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Form Object
$tempFolder = $Args[0]

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Create Purchase Request'
$form.Size = New-Object System.Drawing.Size(300,400)
$form.StartPosition = 'CenterScreen'

# Add OK and Cancel buttons

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(112,320)
$okButton.Size = New-Object System.Drawing.Size(75,25)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,320)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
$form.CancelButton = $cancelButton

# Add Supplier List Box

$supplierLabel = New-Object System.Windows.Forms.Label
$supplierLabel.Location = New-Object System.Drawing.Point(10,20)
$supplierLabel.Size = New-Object System.Drawing.Size(280,20)
$supplierLabel.Text = 'Please select a supplier:'

$supplierListBox = New-Object System.Windows.Forms.ListBox
$supplierListBox.Location = New-Object System.Drawing.Point(10,40)
$supplierListBox.Size = New-Object System.Drawing.Size(260,20)
$supplierListBox.Height = 80

# Add Description Text Box 

$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Location = New-Object System.Drawing.Point(10,130)
$descriptionLabel.Size = New-Object System.Drawing.Size(280,20)
$descriptionLabel.Text = 'Please enter the order description below:'


$descriptionTextBox = New-Object System.Windows.Forms.TextBox
$descriptionTextBox.Location = New-Object System.Drawing.Point(10,150)
$descriptionTextBox.Size = New-Object System.Drawing.Size(260,20)


# Add Documents List Box

$documentsLabel = New-Object System.Windows.Forms.Label
$documentsLabel.Location = New-Object System.Drawing.Point(10,190)
$documentsLabel.Size = New-Object System.Drawing.Size(280,20)
$documentsLabel.Text = 'Please drag and drop documents below:'

$documentsListBox = New-Object Windows.Forms.ListBox
$documentsListBox.Location = New-Object System.Drawing.Point(10,210)
$documentsListBox.Size = New-Object System.Drawing.Size(260,20)
$documentsListBox.Height = 80
$documentsListBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top)
$documentsListBox.IntegralHeight = $False
$documentsListBox.AllowDrop = $True

# Darg and Drop Handler

$documentsListBox.Add_DragOver([System.Windows.Forms.DragEventHandler]{
  if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
  {
      $_.Effect = 'Copy'
  }
  else
  {
	$outlook = New-Object -ComObject Outlook.Application
    $s = $outlook.ActiveExplorer().Selection
	if ($s) {
		$_.Effect = 'Copy'
	}
  }
})

$documentsListBox.Add_DragDrop([System.Windows.Forms.DragEventHandler]{
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        foreach ($file in $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)) {
            $documentsListBox.Items.Add($file)
        }
    }
    else {
		$outlook = New-Object -ComObject Outlook.Application
		$s = $outlook.ActiveExplorer().Selection
        foreach ($item in $s){
            foreach($a in $item.Attachments){
				$file = Join-Path -Path "$pwd\$tempFolder" -ChildPath $a.filename
				$extension = [System.IO.Path]::GetExtension("$file")
				if ($extension -eq ".pdf") {
					$a.SaveAsFile($file)
					$documentsListBox.Items.Add($file)
				}
			}
		}
    }
})


# Render Form

$form.Controls.Add($okButton)
Import-CSV -path $pwd\db\suppliers.csv | ForEach-Object {
	[void] $supplierListBox.Items.Add($_.SupplierName)  
}
$form.Controls.Add($supplierLabel)
$form.Controls.Add($supplierListBox)
$form.Controls.Add($descriptionLabel)
$form.Controls.Add($descriptionTextBox)
$form.Controls.Add($documentsLabel)
$form.Controls.Add($documentsListBox)
$form.Topmost = $true

# Process result

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
	$supplier = $supplierListBox.SelectedItem
	$description = $descriptionTextBox.Text
	$documents = $documentsListBox.Items	
}

$supplier, $description, $documents