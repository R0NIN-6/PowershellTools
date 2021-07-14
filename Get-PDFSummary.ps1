## Extract Table of Contents (plus figures and tables) from a PDF Document via Powershell ##
Param 
(
    [Parameter(Mandatory = $True)]
    [ValidateNotNullOrEmpty()]
    [String]
    $FileName   
)
Function convert-PDFtoText {
	Param(
		[Parameter(Mandatory=$true)][string]$file
	)	
	Add-Type -Path "C:\DLLs\itextsharp.dll"
    $file = $file.trim('"') 
	$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file
    $bm = [iTextSharp.text.pdf.SimpleBookmark]::GetBookmark($pdf)
    #sub objects are in "Kids" property
    $bookmarks = [iTextSharp.text.pdf.SimpleBookmark]::GetBookmark($pdf)
    $PageCount = $pdf.NumberOfPages
    Write-Host "Number of Pages: $PageCount"
    if($bookmarks -ne $NULL)
    {
        Write-host "Bookmarks: `n"
        foreach($bookmark in $bookmarks.Kids){Write-Host $bookmark.Title}
    }
	for ($page = 1; $page -le $pdf.NumberOfPages; $page++){
		$text=[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
		Write-Output $text 
	}	
	$pdf.Close()
}
#Call the function above
$FullText = convert-PDFtoText $FileName 
$TotalLines = $File.Count
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