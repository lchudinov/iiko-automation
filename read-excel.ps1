# Import the Import-Excel module
Import-Module ImportExcel

# Specify the path to your Excel file
$excelFilePath = "7.01.xlsx"
$worksheetName = "ТЮМЕНЬ"
$worksheetName = "ЗАЯВКА"
# $worksheetName = "РЦ"
# $worksheetName = "РЦ1"
# $worksheetName = "ЧЕЛЯБИНСК"

# Use Import-Excel to read the Excel file
$excelData = Import-Excel -Path $excelFilePath -WorksheetName $worksheetName

# Access the first row
$firstRow = $excelData[0]

# Display the data in the first row
$firstRow

$excelData
