Import-Module ImportExcel
Add-Type -AssemblyName System.Windows.Forms

function getCurrentExcelFile {
  $tomorrowDate = (Get-Date).AddDays(1)
  $formattedDate = "{0:d.MM}" -f $tomorrowDate
  return "$($formattedDate).xlsx"
}

function getTommorowtDateForInput {
  $tomorrowDate = (Get-Date).AddDays(1)
  $formattedDate = "{0:ddMM} 0830" -f $tomorrowDate
  return $formattedDate
}

# Specify the path to your Excel file
$excelFilePath = getCurrentExcelFile # "20.01.xlsx"
$excelFilePath
# $worksheetName = "ТЮМЕНЬ"
$worksheetName = "ЗАЯВКА"
# $worksheetName = "РЦ"
# $worksheetName = "РЦ1"
# $worksheetName = "ЧЕЛЯБИНСК"
  
$date = getTommorowtDateForInput

# Use Import-Excel to read the Excel file
$excelData = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName

function getValues($selectedObject) {
  $count = @($selectedObject.PSObject.Properties).count;
  $values = @($selectedObject.PSObject.Properties)[1..$count] | ForEach-Object {$_.Value}
  for ($i = 0; $i -lt $values.Count; $i++) {
    if ($null -eq $values[$i]) {
        $values[$i] = 0
    }
  }
  return $values
}

$nameCol = @($excelData[0].PSObject.Properties)[0].Name
"Навание колонки: $($nameCol)"

$countBefore = $excelData.count
$excelData = $excelData | Where-Object { $_.$nameCol -ne $null }
$countAfter = $excelData.count

if ($countAfter -ne $countBefore) {
  Write-Host "Отфильтровано $($countBefore - $countAfter) строк с пустым именем"
}

$rowCount = $excelData.count
"Число строк: $($rowCount)"

$productCount = @($excelData[0].PSObject.Properties).count - 1
"Число продуктов: $($productCount)"

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true
$form.Text = "Ввод накладнных из файла $($excelFilePath), $($worksheetName), количество продуктов: $($productCount)"
$form.Size = New-Object System.Drawing.Size(700,350)

# Create a ComboBox
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10,30)
$comboBox.Size = New-Object System.Drawing.Size(350,30)

# Populate ComboBox with items
foreach ($item in $excelData) {
    $null = $comboBox.Items.Add($item.$nameCol)
}

$currentItemIndex = 0 

# Create a Label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,70)
$label.Size = New-Object System.Drawing.Size(250,40)

$valuesLabel = New-Object System.Windows.Forms.Label
$valuesLabel.Location = New-Object System.Drawing.Point(10,110)
$valuesLabel.Size = New-Object System.Drawing.Size(1000,20)

$forwardButton = New-Object System.Windows.Forms.Button
$forwardButton.Text = "Следующий"
$forwardButton.Location = New-Object System.Drawing.Point(350,150)
$forwardButton.Size = New-Object System.Drawing.Size(100,30)

$backwardButton = New-Object System.Windows.Forms.Button
$backwardButton.Text = "Предыдущий"
$backwardButton.Location = New-Object System.Drawing.Point(10,150)
$backwardButton.Size = New-Object System.Drawing.Size(100,30)

$inputButton = New-Object System.Windows.Forms.Button
$inputButton.Text = "Ввести количество"
$inputButton.Location = New-Object System.Drawing.Point(150,150)
$inputButton.Size = New-Object System.Drawing.Size(100,40)

$inputWithDateButton = New-Object System.Windows.Forms.Button
$inputWithDateButton.Text = "Ввести количество с датой"
$inputWithDateButton.Location = New-Object System.Drawing.Point(150,250)
$inputWithDateButton.Size = New-Object System.Drawing.Size(100,50)

$inputHeader = New-Object System.Windows.Forms.Button
$inputHeader.Text = "Ввести шапку с даты"
$inputHeader.Location = New-Object System.Drawing.Point(150,200)
$inputHeader.Size = New-Object System.Drawing.Size(100,50)

# Add ComboBox and Label to the form
$form.Controls.Add($comboBox)
$form.Controls.Add($label)
$form.Controls.Add($valuesLabel)
$form.Controls.Add($forwardButton)
$form.Controls.Add($backwardButton)
$form.Controls.Add($inputButton)
$form.Controls.Add($inputHeader)
# $form.Controls.Add($inputWithDateButton)

function selectByIndex($selectedIndex) {
  if ($selectedIndex -lt 0) {
    return
  }
  if ($selectedIndex -ge $rowCount) {
    return
  }
  $global:currentItemIndex = $selectedIndex
  $comboBox.SelectedIndex = $selectedIndex
  $selectedObject = $excelData[$selectedIndex]
  $label.Text = "$($selectedIndex + 1): $($selectedObject.$nameCol)" # -replace "`r`n|`r|`n", " "
  $values = getValues($selectedObject)
  $valuesLabel.Text = $values -join " - "
  # Copy to clipboard
  # $valuesWithTabs = $values -join "`t"
  # $valuesWithTabs | Set-Clipboard
}

function inputNumbersFromClipboard () {
  # Get data from clipboard
  $data = Get-Clipboard
  
  # Split the data into an array of numbers
  $numbers = $data -split "`t"
  
  # Start a loop to process each number
  foreach ($number in $numbers) {
      # Output the current number (you can replace this with your desired action)
      Write-Host "Processing number: $number"
  
      # Send number if not zero using SendKeys
      if ($number -gt 0) {
        [System.Windows.Forms.SendKeys]::SendWait("$($number)")
      }

      # Send down arrow key press using SendKeys
      [System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
  
      # Sleep for a short duration to give time for the key press to take effect
      Start-Sleep -Milliseconds 50
  }
}

function inputDate() {
  Write-Host "Воодим дату $($date)"
  [System.Windows.Forms.SendKeys]::SendWait("$($date)")
}

function goToTableFromDate() {
  Write-Host "Переходим к таблице от даты"
  [System.Windows.Forms.SendKeys]::SendWait("{TAB 10}")
}

function goToTableFromName() {
  Write-Host "Переходим к таблице от имени"
  [System.Windows.Forms.SendKeys]::SendWait("{TAB 9}")
}

function goToName() {
  Write-Host "Переходим к поставщику от даты"
  [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
}

function goToConception() {
  Write-Host "Переходим к концепции от поставщика"
  [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
}

function goToStock() {
  Write-Host "Переходим к складу от концепции"
  [System.Windows.Forms.SendKeys]::SendWait("{TAB 2}")
}

function getAddressfromName {
  param (
    [string]$name
  )
  $address = $name
  $address = $address -replace "БОЛЬШОЙ ФОРМАТ", ""
  $address = $address -replace "МИНИ ФОРМАТ", ""
  $address = $address -replace "ЖМ", ""
  $address = $address -replace "Тюмень", ""
  $address = $address -replace "Челябинск", ""
  $address = $address -replace ",", ""
  return $address.Trim()
}

function inputName() {
  $selectedObject = $excelData[$global:currentItemIndex]
  $name = $selectedObject.$nameCol
  $address = getAddressfromName $name
  Write-Host "вводим имя $($name), номер $($global:currentItemIndex), адрес $($address)"
  [System.Windows.Forms.SendKeys]::SendWait("{DELETE}$($address){ENTER}")
}

function inputConception() {
  $conception = if ($worksheetName -eq "ЗАЯВКА") { "Лаборатория Вкуса ЕКБ" } else {"КО{+}"}
  Write-Host "вводим концепцию $($conception)"
  [System.Windows.Forms.SendKeys]::SendWait("{DELETE}$($conception){ENTER}")
}

function inputStock() {
  $stock = if ($worksheetName -eq "ЗАЯВКА") { "лаб екб готовая" } else {"Тюмень"}
  Write-Host "вводим склад $($stock)"
  [System.Windows.Forms.SendKeys]::SendWait("{DELETE}$($stock){ENTER}")
}

function inputNumbers () {
  Write-Host "currentItemIndex $($global:currentItemIndex)"
  $selectedObject = $excelData[$global:currentItemIndex]
  $numbers = getValues($selectedObject)
  Write-Host "numbers $($numbers)"
  
  # Start a loop to process each number
  foreach ($number in $numbers) {
      # Output the current number (you can replace this with your desired action)
      Write-Host "Processing number: $number"
  
      # Send number if not zero using SendKeys
      if ($number -gt 0) {
        [System.Windows.Forms.SendKeys]::SendWait("$($number)")
      }

      # Send down arrow key press using SendKeys
      [System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
  
      # Sleep for a short duration to give time for the key press to take effect
      Start-Sleep -Milliseconds 50
  }
}


# Add an event handler for ComboBox selection change
$comboBox.Add_SelectedIndexChanged({
  $selectedIndex = $comboBox.SelectedIndex;
  selectByIndex $selectedIndex
})

$forwardButton.Add_Click({
  selectByIndex($global:currentItemIndex + 1)
})

$backwardButton.Add_Click({
  selectByIndex($global:currentItemIndex - 1)
})
$inputButton.Add_Click({
  Start-Sleep 3
  inputNumbers
})

$inputWithDateButton.Add_Click({
  Start-Sleep 3
  inputDate
  goToTableFromDate
  inputNumbers
})

$inputHeader.Add_Click({
  Start-Sleep 3
  inputDate
  goToName
  inputName
  goToConception
  inputConception
  goToStock
  inputStock
  #goToTableFromName
  #inputNumbers
})

selectByIndex 0

# Show the form
$form.ShowDialog()


