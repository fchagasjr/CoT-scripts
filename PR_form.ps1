Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Create Form Object

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Create Purchase Request'
$form.Size = New-Object System.Drawing.Size(300,300)
$form.StartPosition = 'CenterScreen'

# Add OK and Cancel buttons

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,220)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,220)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

# Add Supplier List Box

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select a supplier:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

Import-CSV -path $pwd\db\suppliers.csv | ForEach-Object {

  [void] $listBox.Items.Add($_.SupplierName)
  
}
$form.Controls.Add($listBox)

# Add Description Text Box 

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,130)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter the order description below:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,150)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
	$x = $listBox.SelectedItem
	$y = $textBox.Text
    $x, $y
}
