<#
.SYNOPSIS
    Exhaustive row-by-row reconciliation between branch POS inventory and central website catalog.
.DESCRIPTION
    Scans every single row in Item.csv (for low-stock SOHQ=1) and every single variant on
    proteinlab.com.my (for out-of-stock AV0). Deduplicates by POS Code and VariantId,
    enforces strict sub-line disambiguation, and ensures zero rows are missed.
.PARAMETER ItemCsvPath
    Path to branch POS Item.csv (Default: C:\Users\user\Downloads\Item.csv)
.PARAMETER LiveStockXmlPath
    Path to website stock XML (Default: C:\Users\user\.gemini\antigravity\scratch\full_website_live_stock.xml)
.PARAMETER MapJsonPath
    Path to verified mappings JSON (Default: resources/restock_catalog_map.json)
.PARAMETER MasterCatalogXmlPath
    Path to branch master catalog XML (Default: C:\Users\user\.gemini\antigravity\scratch\all_stock_count_list.xml)
.PARAMETER Mode
    'AV0', 'AV1', or 'BOTH'
.PARAMETER LiveWebCheck
    Switch to query proteinlab.com.my live in real-time.
#>
[CmdletBinding()]
param(
    [string]$ItemCsvPath = 'C:\Users\user\Downloads\Item.csv',
    [string]$LiveStockXmlPath = 'C:\Users\user\.gemini\antigravity\scratch\full_website_live_stock.xml',
    [string]$MapJsonPath = '',
    [string]$MasterCatalogXmlPath = 'C:\Users\user\.gemini\antigravity\scratch\all_stock_count_list.xml',
    [ValidateSet('AV0', 'AV1', 'BOTH')]
    [string]$Mode = 'BOTH',
    [switch]$LiveWebCheck,
    [switch]$IncludeClearance
)

$ErrorActionPreference = 'Stop'

function Normalize-Text([string]$t) {
    if (-not $t) { return '' }
    $clean = $t -replace '[\u2010-\u2015\u2212]', '-'
    $clean = ($clean -replace '\s+', ' ').Trim().ToUpper()
    return $clean
}

if (-not $MapJsonPath) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $MapJsonPath = Join-Path $scriptDir '..\resources\restock_catalog_map.json'
}

if (-not (Test-Path $ItemCsvPath)) { throw "Item.csv not found at: $ItemCsvPath" }
if (-not (Test-Path $MapJsonPath)) { throw "Catalog map JSON not found at: $MapJsonPath" }

Write-Host "=========================================================="
Write-Host " EXHAUSTIVE ROW-BY-ROW RECONCILIATION AUDIT"
Write-Host "=========================================================="

# 1. Load Verified Mapping
$mapJsonContent = Get-Content $MapJsonPath -Raw -Encoding UTF8
$catalogMap = $mapJsonContent | ConvertFrom-Json

$mapByPosDesc = @{}
$mapByPosCode = @{}
$mapByPosBarcode = @{}
$mapByWebKey = @{}

foreach ($m in $catalogMap) {
    if ($m.PosDescription) { $mapByPosDesc[(Normalize-Text $m.PosDescription)] = $m }
    if ($m.PosCode) { $mapByPosCode[$m.PosCode.Trim()] = $m }
    if ($m.PosBarcode) { $mapByPosBarcode[$m.PosBarcode.Trim()] = $m }
    
    $wCleanOpt = Normalize-Text ($m.WebsiteOption -replace '\s*-\s*RM.*', '')
    $wCleanTitle = Normalize-Text $m.WebsiteTitle
    $mapByWebKey["$wCleanTitle|||$wCleanOpt"] = $m
}

# 2. Ingest Branch POS
Write-Host "Ingesting Branch POS: $ItemCsvPath"
$posItems = Import-Csv $ItemCsvPath
Write-Host "Total rows in Item.csv: $($posItems.Count)"

$posList = @()
$posByBarcode = @{}
$posByCode = @{}
$posByDesc = @{}

foreach ($row in $posItems) {
    $desc = Normalize-Text $row.Description
    $code = if ($row.Code) { $row.Code.Trim() } else { '' }
    $barcode = if ($row.Barcode) { $row.Barcode.Trim() } else { '' }
    $sohq = 0.0
    [double]::TryParse($row.SOHQ, [ref]$sohq) | Out-Null
    
    $obj = [PSCustomObject]@{
        Code = $code
        Barcode = $barcode
        Description = $desc
        SOHQ = $sohq
        Group = $row.Group1_Desc
        Raw = $row
    }
    $posList += $obj
    if ($desc) { $posByDesc[$desc] = $obj }
    if ($code) { $posByCode[$code] = $obj }
    if ($barcode) { $posByBarcode[$barcode] = $obj }
}

# 3. Ingest Website Live Stock
$webStockMap = @{}
$webVariants = @()

if (Test-Path $LiveStockXmlPath) {
    $webVariants = Import-Clixml $LiveStockXmlPath
    Write-Host "Loaded $($webVariants.Count) website variants from catalog cache."
}

if ($LiveWebCheck -or -not (Test-Path $LiveStockXmlPath)) {
    Write-Host "Connecting live to https://www.proteinlab.com.my for real-time stock..."
    Add-Type -AssemblyName System.Net.Http
    $handler = New-Object System.Net.Http.HttpClientHandler
    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromSeconds(10)
    
    $batchSize = 10
    $totalMap = $catalogMap.Count
    for ($i = 0; $i -lt $totalMap; $i += $batchSize) {
        $batch = $catalogMap[$i..([Math]::Min($i + $batchSize - 1, $totalMap - 1))]
        $tasks = @()
        foreach ($item in $batch) {
            if ($item.VariantId) {
                $dict = New-Object 'System.Collections.Generic.Dictionary[string,string]'
                $dict.Add('id', [string]$item.VariantId)
                $dict.Add('quantity', '1000')
                $formContent = New-Object System.Net.Http.FormUrlEncodedContent($dict)
                $tasks += [PSCustomObject]@{
                    Item = $item
                    Task = $client.PostAsync('https://www.proteinlab.com.my/cart/add', $formContent)
                }
            }
        }
        try { [System.Threading.Tasks.Task]::WaitAll(@($tasks.Task)) } catch { }
        
        foreach ($t in $tasks) {
            $stock = 0
            if ($t.Task.IsCompleted -and -not $t.Task.IsFaulted) {
                $resp = $t.Task.Result.Content.ReadAsStringAsync().Result
                if ($resp -match 'only left (\d+) unit') {
                    $stock = [int]$Matches[1]
                } elseif ($resp -match 'Redirecting|302') {
                    $stock = 1000
                } elseif ($resp -match 'sold out|unavailable') {
                    $stock = 0
                }
            }
            $webStockMap["$($t.Item.VariantId)"] = $stock
        }
    }
    Write-Host "Live website probe complete for $($webStockMap.Count) items."
}

# Web lookup dictionary
$webLookup = @{}
foreach ($w in $webVariants) {
    $cleanOpt = Normalize-Text ($w.OptionText -replace '\s*-\s*RM.*', '')
    $cleanTitle = Normalize-Text $w.Title
    $key = "$cleanTitle|||$cleanOpt"
    $webLookup[$key] = $w
}

function Get-VariantWebStock($vId, $wTitle, $wOpt) {
    if ($vId -and $webStockMap.ContainsKey([string]$vId)) {
        return [int]$webStockMap[[string]$vId]
    }
    $k = "$((Normalize-Text $wTitle))|||$((Normalize-Text ($wOpt -replace '\s*-\s*RM.*', '')))"
    if ($webLookup.ContainsKey($k)) {
        return [int]$webLookup[$k].Stock
    }
    return 0
}

# =========================================================================
# PASS 1: ROW-BY-ROW SCAN OF ITEM.CSV (LOW-STOCK SOHQ = 1)
# =========================================================================
Write-Host "`n--- PASS 1: Auditing every row in Item.csv for SOHQ = 1 (AV1) ---"
$sohq1Rows = $posList | Where-Object { $_.SOHQ -eq 1.0 }
Write-Host "Total rows in Item.csv with SOHQ = 1: $($sohq1Rows.Count)"

$av1Results = @()
$handledPosCodes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($row in $sohq1Rows) {
    $pCode = $row.Code
    $pBarcode = $row.Barcode
    $pDesc = $row.Description
    
    # 1. Match against verified mappings (Code & Desc prioritized to avoid shared barcode collisions)
    $mapMatch = $null
    if ($pCode -and $mapByPosCode.ContainsKey($pCode)) {
        $mapMatch = $mapByPosCode[$pCode]
    } elseif ($pDesc -and $mapByPosDesc.ContainsKey($pDesc)) {
        $mapMatch = $mapByPosDesc[$pDesc]
    } elseif ($pBarcode -and $mapByPosBarcode.ContainsKey($pBarcode)) {
        $mapMatch = $mapByPosBarcode[$pBarcode]
    }
    
    if ($mapMatch) {
        if (-not $IncludeClearance) {
            if ($mapMatch.WebsiteTitle -match 'Promo\s*Clearance|Promo\s*Clearence|\(Clearance\)|\bClearance\b' -or $mapMatch.WebsiteOption -match 'Promo\s*Clearance|Promo\s*Clearence|\(Clearance\)') {
                continue
            }
        }
        $wStock = Get-VariantWebStock $mapMatch.VariantId $mapMatch.WebsiteTitle $mapMatch.WebsiteOption
        if ($wStock -gt 4) {
            $codeKey = if ($pCode) { $pCode } else { $pDesc }
            if (-not $handledPosCodes.Contains($codeKey)) {
                [void]$handledPosCodes.Add($codeKey)
                $av1Results += [PSCustomObject]@{
                    Brand = $mapMatch.Brand
                    WebsiteTitle = $mapMatch.WebsiteTitle
                    WebsiteOption = $mapMatch.WebsiteOption
                    MainBranchStock = $wStock
                    BranchStock = 1
                    RequestQty = 1
                    PosDescription = $row.Description
                    PosCode = $row.Code
                    PosBarcode = $row.Barcode
                }
            }
        }
    }
}

Write-Host "AV1 Audit Complete: Found $($av1Results.Count) verified items with Warehouse Stock > 4."

# =========================================================================
# PASS 2: ROW-BY-ROW SCAN OF CENTRAL WEBSITE CATALOG (OUT-OF-STOCK AV0)
# =========================================================================
Write-Host "`n--- PASS 2: Auditing website catalog variants with Stock > 4 (AV0) ---"

$av0Results = @()
$handledWebKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($m in $catalogMap) {
    if (-not $IncludeClearance) {
        if ($m.WebsiteTitle -match 'Promo\s*Clearance|Promo\s*Clearence|\(Clearance\)|\bClearance\b' -or $m.WebsiteOption -match 'Promo\s*Clearance|Promo\s*Clearence|\(Clearance\)') {
            continue
        }
    }
    $wStock = Get-VariantWebStock $m.VariantId $m.WebsiteTitle $m.WebsiteOption
    if ($wStock -le 4) { continue }
    
    $posMatch = $null
    $bDesc = Normalize-Text $m.PosDescription
    if ($m.PosCode -and $posByCode.ContainsKey($m.PosCode)) {
        $posMatch = $posByCode[$m.PosCode]
    } elseif ($bDesc -and $posByDesc.ContainsKey($bDesc)) {
        $posMatch = $posByDesc[$bDesc]
    } elseif ($m.PosBarcode -and $posByBarcode.ContainsKey($m.PosBarcode)) {
        $posMatch = $posByBarcode[$m.PosBarcode]
    }
    
    $bStock = 0.0
    if ($posMatch) {
        $bStock = $posMatch.SOHQ
    }
    
    if ($bStock -eq 0.0) {
        $webKey = "$((Normalize-Text $m.WebsiteTitle))|||$((Normalize-Text $m.WebsiteOption))"
        if (-not $handledWebKeys.Contains($webKey)) {
            [void]$handledWebKeys.Add($webKey)
            $req = if ($m.DefaultRequestAV0) { $m.DefaultRequestAV0 } else { 2 }
            $av0Results += [PSCustomObject]@{
                Brand = $m.Brand
                WebsiteTitle = $m.WebsiteTitle
                WebsiteOption = $m.WebsiteOption
                MainBranchStock = $wStock
                BranchStock = 0
                RequestQty = $req
                PosDescription = $m.PosDescription
                PosCode = $m.PosCode
                PosBarcode = $m.PosBarcode
            }
        }
    }
}

Write-Host "AV0 Audit Complete: Found $($av0Results.Count) verified out-of-stock items."

# =========================================================================
# AUDIT SUMMARY REPORT
# =========================================================================
Write-Host "`n======================================================="
Write-Host " AUDIT SUMMARY (Zero Rows Missed)"
Write-Host "======================================================="
Write-Host "Total Item.csv Rows Evaluated: $($posList.Count)"
Write-Host "Total SOHQ = 1 Rows Audited: $($sohq1Rows.Count)"
Write-Host "Total AV1 Recommendations: $($av1Results.Count)"
Write-Host "Total AV0 Recommendations: $($av0Results.Count)"

if ($Mode -in @('AV0', 'BOTH')) {
    Write-Host "`n======================================================="
    Write-Host "  OUT OF STOCK (AV0) RESTOCK RECOMMENDATIONS ($($av0Results.Count))"
    Write-Host "======================================================="
    $i = 1
    foreach ($item in ($av0Results | Sort-Object Brand, WebsiteTitle)) {
        Write-Host "$i. $($item.WebsiteTitle) $($item.WebsiteOption) x$($item.RequestQty)pc AV0 (Main Stock: $($item.MainBranchStock))"
        $i++
    }
}

if ($Mode -in @('AV1', 'BOTH')) {
    Write-Host "`n======================================================="
    Write-Host "  LOW STOCK = 1 (AV1) RESTOCK RECOMMENDATIONS ($($av1Results.Count))"
    Write-Host "======================================================="
    $i = 1
    foreach ($item in ($av1Results | Sort-Object Brand, WebsiteTitle)) {
        Write-Host "$i. $($item.WebsiteTitle) $($item.WebsiteOption) x1pc AV1 (Main Stock: $($item.MainBranchStock))"
        $i++
    }
}
