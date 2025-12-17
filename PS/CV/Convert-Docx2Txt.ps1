param(
    [Parameter(Mandatory = $true)]
    [string]$SourceFolder,

    [Parameter(Mandatory = $true)]
    [string]$OutputFolder,

    [Nullable[DateTime]]$From = $null,
    [Nullable[DateTime]]$To   = $null
)

#.\Convert-Docx2Txt.ps1 "c:\Users\t\Documents\CV\doc2txt\AS" "c:\temp\doc2txt\dest\" -From "2025-11-05" -To "2025-11-30"

Write-Host "Start"

if ($null -ne $From -and $null -ne $To -and $From -gt $To) {
    Write-Host "Done.exit"
}

if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$logFileName = "conversion_log.txt"
$logFile = Join-Path $OutputFolder $logFileName
if (!(Test-Path $logFile)) {
    "Timestamp,Status,SourceFile,OutputFile,Message" | Out-File $logFile -Encoding UTF8
}

"$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),START,,,Beginning processing" |
    Out-File $logFile -Append -Encoding UTF8

$word = New-Object -ComObject Word.Application
$word.Visible = $false

$files = Get-ChildItem -Path $SourceFolder -Filter *.docx -Recurse

foreach ($file in $files) {
    if ($From -ne $null -and $file.LastWriteTime -lt $From) { continue }
    if ($To   -ne $null -and $file.LastWriteTime -gt $To)   { continue }
    
    $relativePath = $file.FullName.Substring($SourceFolder.Length).TrimStart('\')
    $relativeTextPath = [System.IO.Path]::ChangeExtension($relativePath, ".txt")
    $destPath = Join-Path $OutputFolder $relativeTextPath
    $destDir = Split-Path $destPath
    if (!(Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    Write-Host "Converting: $($file.FullName)"
    Write-Host "   → $destPath"

    try {
        $doc = $word.Documents.Open($file.FullName)
        # 2 = wdFormatText
        $doc.SaveAs([ref]$destPath, [ref]2)
        $doc.Close()
        "Success,$($file.FullName),$destPath,OK" | Out-File $logFile -Append -Encoding UTF8
        }
    catch {
        $msg = $_.Exception.Message.Replace(",", ";")
        Write-Warning "Failed to convert: $($file.FullName)"
        "Failed,$($file.FullName),,${msg}" | Out-File $logFile -Append -Encoding UTF8
    }
}

$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

"$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),END,,,Finished processing" |
    Out-File $logFile -Append -Encoding UTF8
"" | Out-File $logFile -Append

Write-Host "Done."
