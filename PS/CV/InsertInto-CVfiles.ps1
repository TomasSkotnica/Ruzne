param (
    [string]$RootPath,
    [string]$OutPath
)

# cd c:\work\Ruzne\PS\CV\; powershell -ExecutionPolicy Bypass -File '.\InsertInto-CVfiles.ps1'; & .\InsertInto-CVfiles.ps1 -RootPath c:\Users\t\Documents\CV\_backoffice\ -OutPath c:\temp\doc2txt\insert2SQL.sql

    # SQL-safe values
    function SqlValue($v) {
        if ($null -eq $v) { "NULL" }
        else { "'" + ($v -replace "'", "''") + "'" }
    }

Write-Host "RootPath = $RootPath"
$RootPath = (Resolve-Path $RootPath).Path

Remove-Item $OutPath -Confirm

Get-ChildItem -Path $RootPath -Recurse -File | ForEach-Object {
    Write-Host "FullName = $_"
    # Relative path (without root)
    $relativePath = $_.FullName.Substring($RootPath.Length).TrimStart('\')

    # Split path into parts
    $parts = $relativePath -split '\\'

    # Folder levels
    $level1 = if ($parts.Count -ge 2) { $parts[0] } else { $null }
    $level2 = if ($parts.Count -ge 3) { $parts[1] } else { $null }
    $level3 = if ($parts.Count -ge 4) { $parts[2] } else { $null }

    $fileName = SqlValue $_.Name
    $relPath  = SqlValue $relativePath
    $date     = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")

    $l1 = SqlValue $level1
    $l2 = SqlValue $level2
    $l3 = SqlValue $level3

    # Output SQL INSERT
    
    "INSERT INTO Files (FileName, RelativePath, LastModified, Level1, Level2, Level3) VALUES ($fileName, $relPath, $date, $l1, $l2, $l3)" | Out-File -Append $OutPath -Encoding UTF8

}


    # SQL-safe values
    function SqlValue1($v) {
        if ($null -eq $v) { return $null }
        else { return  $v }
    }
