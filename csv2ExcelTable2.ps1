#Define locations and variables
$csv_path = "$HOME\Desktop\ExcelDemo\temp.csv" #Location of the source file
$xlsx_path = "$HOME\Desktop\ExcelDemo\Output.xlsx" #Desired location of output
$Worksheet_name = "Domain Users" # Name of the sheet

#Check connection to Azure AD and prompt for credentials if not established
$Tenant = Get-AzureADTenantDetail
if($Tenant){Write-Host "Connected Azure AD Tenant: $($Tenant.DisplayName)"}
Else{
$Credential = Get-Credential
Connect-AzureAD -Credential $Credential | Out-Null
}
#Get data and export to csv, this example is for Users in an Azure Active Directory Tenant:
Get-AzureADUser -All:$true | Where-Object { $_.AccountEnabled -eq $True }`
 | select DisplayName, UserPrincipalName,JobTitle,Mobile, Department,City,State | sort DisplayName `
 | Export-Csv -Path $csv_path -NoTypeInformation

#create Excel COM Object, and import CSV data as a table with headers
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false #  Set to false to hide Excel
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
#Optional: delete CSV file
Remove-Item $csv_path -Force
#Delete Excel object to free memory
$excel.Quit()