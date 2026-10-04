<#
.SYNOPSIS
    Applies stock updates to a Shopee Seller Centre Sales Info Mass Update template.
.DESCRIPTION
    Preserves exact template metadata and formatting (rows 1-6).
    Iterates row-by-row starting from row 7 to update Column 11 (Stock).
    Outputs both ready-to-upload XLSX and CSV versions.
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$OriginalShopeeXlsx,

    [Parameter(Mandatory=$true)]
    [string]$DiscrepanciesCsv,

    [Parameter(Mandatory=$false)]
    [string]$OutputXlsx,

    [Parameter(Mandatory=$false)]
    [string]$OutputCsv
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $OriginalShopeeXlsx)) {
    throw "Original Shopee XLSX not found at: $OriginalShopeeXlsx"
}
if (-not (Test-Path $DiscrepanciesCsv)) {
    throw "Discrepancies CSV not found at: $DiscrepanciesCsv"
}

# Determine default output filenames if not provided
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$origDir = Split-Path -Parent (Resolve-Path $OriginalShopeeXlsx)
$origName = [System.IO.Path]::GetFileNameWithoutExtension($OriginalShopeeXlsx)

if (-not $OutputXlsx) {
    $OutputXlsx = Join-Path $origDir "${origName}_UPDATED_${timestamp}.xlsx"
}
if (-not $OutputCsv) {
    $OutputCsv = Join-Path $origDir "${origName}_UPDATED_${timestamp}.csv"
}

Write-Host "=== Shopee Template Stock Updater ===" -ForegroundColor Cyan
Write-Host "Input Shopee File : $OriginalShopeeXlsx"
Write-Host "Input Updates CSV : $DiscrepanciesCsv"
Write-Host "Target Output XLSX: $OutputXlsx"
Write-Host "Target Output CSV : $OutputCsv"

# Load updates
$updatesCsvData = Import-Csv $DiscrepanciesCsv
$stockMap = @{}

foreach ($row in $updatesCsvData) {
    $prodId = $row.'Product ID'
    $vId = $row.'Variation ID'
    if (-not $prodId) { $prodId = $row.ProductId }
    if (-not $vId) { $vId = $row.VariationId }
    
    $newStock = $row.'New Stock'
    if ($null -eq $newStock -or $newStock -eq "") {
        $newStock = $row.'POS SOHQ'
    }
    if ($null -eq $newStock -or $newStock -eq "") {
        $newStock = $row.'Target Stock'
    }

    if ($prodId -and $vId -and ($null -ne $newStock)) {
        $key = "${prodId}_${vId}".Trim()
        $stockMap[$key] = [int][Math]::Floor([double]$newStock)
    }
}

Write-Host "Loaded $($stockMap.Count) distinct variation stock update targets." -ForegroundColor Green

# Open Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null

try {
    $resolvedOriginal = (Resolve-Path $OriginalShopeeXlsx).Path
    $wb = $excel.Workbooks.Open($resolvedOriginal)
    $ws = $wb.Sheets.Item(1)
    $totalRows = $ws.UsedRange.Rows.Count

    Write-Host "Shopee template total rows in sheet: $totalRows (Data begins at row 7)" -ForegroundColor Cyan

    $updatedCount = 0
    $unchangedCount = 0
    $updateLog = @()

    for ($r = 7; $r -le $totalRows; $r++) {
        $rowPid = $ws.Cells.Item($r, 1).Text.Trim()
        $rowVid = $ws.Cells.Item($r, 3).Text.Trim()
        $key = "${rowPid}_${rowVid}"

        if ($stockMap.ContainsKey($key)) {
            $oldStock = $ws.Cells.Item($r, 11).Text.Trim()
            $targetStock = $stockMap[$key]

            if ($oldStock -ne "$targetStock") {
                $ws.Cells.Item($r, 11).Value2 = $targetStock
                $updatedCount++
                $updateLog += "Row $r | PID: $rowPid | VID: $rowVid | Stock changed: $oldStock -> $targetStock"
            } else {
                $unchangedCount++
            }
        }
    }

    Write-Host "Successfully updated $updatedCount rows (already identical: $unchangedCount)." -ForegroundColor Green
    
    # Save as XLSX
    $wb.SaveCopyAs($OutputXlsx)
    Write-Host "Saved updated XLSX: $OutputXlsx" -ForegroundColor Green

    # Save as CSV
    # Excel COM XlFileFormat xlCSV = 6
    $wb.SaveAs($OutputCsv, 6)
    Write-Host "Saved updated CSV : $OutputCsv" -ForegroundColor Green

    return [PSCustomObject]@{
        Status = "Success"
        TotalRowsAudited = $totalRows - 6
        RowsUpdated = $updatedCount
        RowsUnchanged = $unchangedCount
        OutputXlsx = $OutputXlsx
        OutputCsv = $OutputCsv
    }
}
finally {
    if ($wb) {
        $wb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
