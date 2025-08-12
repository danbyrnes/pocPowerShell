# Import the ImportExcel module
Import-Module ImportExcel

# Define the path to the Excel file

#SampleFile-InvTracker.xlsx
$excelFilePath = "d:\!Programming\PS\Excel\InvestmentTracker.xlsx"

# Import data from the first worksheet of the Excel file
$data = Import-Excel -Path $excelFilePath

$data | get-member

# Display the data in the console
$data | Format-Table -AutoSize

