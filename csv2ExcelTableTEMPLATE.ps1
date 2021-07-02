### Template for converting data from Powershell objects to an Excel Table ###

#Define variables
$csv_path = "$HOME\Desktop\temp.csv" #Temporary file, will be deleted at end of script
$xlsx_path = "$HOME\Desktop\Spreadsheet.xlsx" #filename for the generated Excel file
$Worksheet_name = "Sheet1" #name of the worksheet
<#
#Powershell Cmdlets have the syntax: verb-noun -parameters <values> |  Where-Object{comparison statement } | select PropertyNames 
ex 1. Get-Service -Displayname "*network*"
ex 2. Get-AzureADUser -Department Finance | Where-Object{$_.Country -eq "US" }
ex 3. get-childitem -Path $HOME\Desktop | select Fullname
#>
<Enter Cmdlet here>| select `<enter the property names for the Table Headers> `
| sort `<Enter Property to sort by> `
| Export-Csv -Path $csv_path -NoTypeInformation

#create Excel COM Object, and import CSV data as a table with headers
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false #  Change to true to open Excel, useful for debugging
$wb = $excel.Workbooks.Open("$csv_path") 
$excel.Worksheets[1].Columns.AutoFit() | Out-Null
$worksheet = $excel.Worksheets.item(1)
$worksheet.name = "$Worksheet_name"
$list = $excel.ActiveSheet.ListObjects.Add(
            [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, 
            $excel.ActiveCell.CurrentRegion, 
            $null,
            [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes 
        )
$excel.DisplayAlerts = $false 
$excel.ActiveWorkbook.SaveAs(
            "$xlsx_path", 
            [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault,
            $null,
            $null,
            $null,
            $false, # Do not create backup
            $null,
            [Microsoft.Office.Interop.Excel.XlSaveConflictResolution]::xlLocalSessionChanges 
        )
#Clean-up resources
Remove-Item $csv_path -Force
$excel.Quit()