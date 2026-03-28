param (
    [Parameter(Mandatory=$true)]
    [string]$FolderPath
)

# Check if folder exists
if (-Not (Test-Path $FolderPath)) {
    Write-Host "Folder not found: $FolderPath"
    exit
}

# Create PDFs folder inside the given folder
$PdfFolder = Join-Path $FolderPath "PDFs"
if (-Not (Test-Path $PdfFolder)) {
    New-Item -ItemType Directory -Path $PdfFolder | Out-Null
}

# Get all PPTX and PPT files
$PptFiles = Get-ChildItem -Path $FolderPath -Include *.ppt, *.pptx -Recurse

if ($PptFiles.Count -eq 0) {
    Write-Host "No PPT or PPTX files found in $FolderPath"
    exit
}

# Initialize summary counters
$total = $PptFiles.Count
$converted = 0
$skipped = 0
$failed = 0

# Create PowerPoint COM object
$PowerPoint = New-Object -ComObject PowerPoint.Application
$PowerPoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

Write-Host "Starting conversion of $total files..."

foreach ($file in $PptFiles) {
    $pdfPath = Join-Path $PdfFolder ($file.BaseName + ".pdf")
    
    if (Test-Path $pdfPath) {
        Write-Host "Skipping (already exists): $($file.Name)"
        $skipped++
        continue
    }

    try {
        Write-Host "Converting: $($file.Name)"
        $presentation = $PowerPoint.Presentations.Open($file.FullName, $false, $false, $false)
        $presentation.SaveAs($pdfPath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
        $presentation.Close()
        $converted++
    } catch {
        Write-Host "Failed to convert: $($file.Name)"
        $failed++
    }
}

# Quit PowerPoint
$PowerPoint.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($PowerPoint) | Out-Null

# Summary
Write-Host ""
Write-Host "Conversion completed!"
Write-Host "Total files found: $total"
Write-Host "Converted: $converted"
Write-Host "Skipped: $skipped"
Write-Host "Failed: $failed"