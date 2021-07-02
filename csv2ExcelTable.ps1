#Define locations and variables
$csv_path = "$HOME\Desktop\ExcelDemo\Input.csv"
$xlsx_path = "$HOME\Desktop\ExcelDemo\Output.xlsx" #Desired location of output
$Worksheet_name = "Programs" # Name of the sheet

#This example is for the installed programs on the machine:
Get-WmiObject -Class Win32_Product |Where-Object{$_.name -ne $NULL}| select Name, Vendor, InstallDate, Caption `
| sort Name | Export-Csv -Path $csv_path -NoTypeInformation

#create Excel COM Object, and import CSV data as a table with headers
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false 
$wb = $excel.Workbooks.Open("$csv_path") 
$excel.Worksheets[1].Columns.AutoFit() | Out-Null
$worksheet = $excel.Worksheets.item(1)
$worksheet.name = "$Worksheet_name"
$list = $excel.ActiveSheet.ListObjects.Add(
            [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, # Add a range
            $excel.ActiveCell.CurrentRegion, # Get the current region, by default A1 is selected so it'll select all contiguous rows
            $null,
            [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes # Yes, we have a header row
        )
$excel.DisplayAlerts = $false # Ignore / hide alerts
$excel.ActiveWorkbook.SaveAs(
            "$xlsx_path", 
            [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault, # Save as .xlsx
            $null,
            $null,
            $null,
            $false, # Do not create backup
            $null,
            [Microsoft.Office.Interop.Excel.XlSaveConflictResolution]::xlLocalSessionChanges # Do not prompt for changes
        )
#Clean-up resources
Remove-Item $csv_path -Force
$excel.Quit()