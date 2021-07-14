## Extract Table of Contents (plus figures and tables) from a Word Document via Powershell ##

# create the object for a Word Document
$Word_obj = New-Object -comobject Word.Application
#Show/hide the word application 
$Word_obj.Visible = $False
#Select file via browser
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
[void]$FileBrowser.ShowDialog()
$file=$FileBrowser.FileName

$doc = $Word_obj.Documents.Open($file) 
$FullText = $doc.Paragraphs | ForEach-Object {
    $_.Range.Text
}
Write-Host "Word count is $($doc.Words.Count)"
#Type 88 is hyperlinks
foreach($f in $doc.Fields){
     if ($f.Type -eq 88){$f.Result.Text
     }}
$TotalLines = $FullText.Count
$Dictionary = @{}
$LineCount = 0
$FullText | foreach {
    $Line = $_
    $LineCount++
    Write-Progress -Activity "Processing words..." `
    -PercentComplete ($LineCount*100/$TotalLines)     
    $Line -split "[^a-zA-Z]" | foreach {
        $Word = $_.ToUpper()
        If ($Word[0] -ge 'A' -and $Word[0] -le "Z") {
            $WordCount++
            If ($Dictionary.ContainsKey($Word)) {
                $Dictionary.$Word++
            } else {
                $Dictionary.Add($Word, 1)
            }
        }
    } 
}

$Dictionary.GetEnumerator() | ? { $_.Name.Length -gt 4 } | 
Sort Value -Descending | Select -First 10

#Release the COM variable and terminate Word
$doc.close()
$Word_obj.Quit()