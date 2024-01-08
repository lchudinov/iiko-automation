Import-Module ImportExcel
Add-Type -AssemblyName System.Windows.Forms


# Specify the path to your Excel file
$excelFilePath = "7.01.xlsx"
$worksheetName = "ТЮМЕНЬ"
$worksheetName = "ЗАЯВКА"
# $worksheetName = "РЦ"
# $worksheetName = "РЦ1"
# $worksheetName = "ЧЕЛЯБИНСК"

# Use Import-Excel to read the Excel file
#$excelData = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName -NoHeader 
$excelData = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName
# $excelData
# $excelData | ForEach-Object {"Item: [$PSItem]"}

# foreach ( $node in $excelData )
# {
#     $node = $node | Select-Object @{Name='Название'; Expression={$_.'Интервал'}}
# }

$nameCol = @($excelData[0].PSObject.Properties)[0].Name
$nameCol

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true
$form.Text = "Ввод накладной"
$form.Size = New-Object System.Drawing.Size(400,200)

# Create a ComboBox
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10,30)
$comboBox.Size = New-Object System.Drawing.Size(350,30)

# Populate ComboBox with items
foreach ($item in $excelData) {
    $null = $comboBox.Items.Add($item.$nameCol)
}

# Create a Label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,70)
$label.Size = New-Object System.Drawing.Size(250,40)

$valuesLabel = New-Object System.Windows.Forms.Label
$valuesLabel.Location = New-Object System.Drawing.Point(10,110)
$valuesLabel.Size = New-Object System.Drawing.Size(400,20)

# Add ComboBox and Label to the form
$form.Controls.Add($comboBox)
$form.Controls.Add($label)
$form.Controls.Add($valuesLabel)



# Add an event handler for ComboBox selection change
$comboBox.Add_SelectedIndexChanged({
  $selectedItem = $comboBox.SelectedItem
  $selectedIndex = $comboBox.FindStringExact($selectedItem);
  $selectedObject = $excelData | Where-Object { $_.$nameCol -eq $selectedItem }
  $label.Text = "$($selectedIndex): $($selectedObject.$nameCol)" # -replace "`r`n|`r|`n", " "
  $count = @($selectedObject.PSObject.Properties).count;
  $values = @($selectedObject.PSObject.Properties)[1..$count] | ForEach-Object {$_.Value}
  for ($i = 0; $i -lt $values.Count; $i++) {
    if ($values[$i] -eq $null) {
        $values[$i] = 0
    }
  }
  $valuesWithTabs = $values -join "`t"
  $valuesLabel.Text = $values
  $valuesWithTabs | Set-Clipboard
})

# Show the form
$form.ShowDialog()


