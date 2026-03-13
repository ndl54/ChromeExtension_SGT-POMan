param(
    [string]$PartnerOutputPath = "doitacgiaohang.csv",
    [string]$SupplierOutputPath = "nhacungcap.csv",
    [string]$ProductOutputPath = "sanpham.csv"
)

$sheetId = "1jjzb4CUl_9iJ9Hlgov7tqqifrRJPojTGkCItJ22PSTk"
$partnerSheet = "Doi_Tac_Giao_Hang"
$supplierSheet = "Nha_Cung_Cap"
$productSheet = "San_Pham"
$partnerUrl = "https://docs.google.com/spreadsheets/d/$sheetId/export?format=csv&sheet=$partnerSheet"
$supplierUrl = "https://docs.google.com/spreadsheets/d/$sheetId/export?format=csv&sheet=$supplierSheet"
$productUrl = "https://docs.google.com/spreadsheets/d/$sheetId/export?format=csv&sheet=$productSheet"

function Get-CsvRows([string]$Uri) {
    return (Invoke-WebRequest -Uri $Uri).Content | ConvertFrom-Csv
}

function Export-CsvRows([object[]]$Rows, [string]$OutputPath) {
    $Rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
}

function Normalize-Cell([string]$Value) {
    return ([string]$Value -replace '\s+', ' ').Trim()
}

function Sanitize-ProductCode([string]$Value) {
    $sanitized = Normalize-Cell $Value
    if (-not $sanitized) {
        return ''
    }

    $suffixPatterns = @(
        '\s+\[IMEI\]$',
        '\s+\[(?:CŨ|CU)\]$',
        '\s+MN$',
        '\s+ML$'
    )

    do {
        $previous = $sanitized
        foreach ($pattern in $suffixPatterns) {
            $sanitized = [regex]::Replace($sanitized, $pattern, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $sanitized = Normalize-Cell $sanitized
        }
    } while ($sanitized -ne $previous)

    return $sanitized
}

$partnerRows = Get-CsvRows -Uri $partnerUrl | Sort-Object `
    @{ Expression = { [string]($_.PSObject.Properties[1].Value).Trim().Length } }, `
    @{ Expression = { [string]($_.PSObject.Properties[1].Value).Trim().ToLowerInvariant() } }, `
    @{ Expression = { [string]($_.PSObject.Properties[0].Value).Trim().ToLowerInvariant() } }
Export-CsvRows -Rows $partnerRows -OutputPath $PartnerOutputPath
Write-Host "Saved sorted partner data to $PartnerOutputPath"

$supplierRows = Get-CsvRows -Uri $supplierUrl | Sort-Object `
    @{ Expression = { [string]($_.PSObject.Properties[1].Value).Trim().Length } }, `
    @{ Expression = { [string]($_.PSObject.Properties[1].Value).Trim().ToLowerInvariant() } }, `
    @{ Expression = { [string]($_.PSObject.Properties[0].Value).Trim().ToLowerInvariant() } }
Export-CsvRows -Rows $supplierRows -OutputPath $SupplierOutputPath
Write-Host "Saved sorted supplier data to $SupplierOutputPath"

$productRowsRaw = Get-CsvRows -Uri $productUrl
$productHeader = if ($productRowsRaw.Count -gt 0) { $productRowsRaw[0].PSObject.Properties[0].Name } else { 'Mã hàng' }
$seenProducts = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$productRows = foreach ($row in $productRowsRaw) {
    $sanitizedCode = Sanitize-ProductCode ([string]$row.PSObject.Properties[0].Value)
    if (-not $sanitizedCode) {
        continue
    }

    if (-not $seenProducts.Add($sanitizedCode)) {
        continue
    }

    [PSCustomObject]@{
        $productHeader = $sanitizedCode
    }
} | Sort-Object `
    @{ Expression = { [string]($_.PSObject.Properties[0].Value).Trim().Length } }, `
    @{ Expression = { [string]($_.PSObject.Properties[0].Value).Trim().ToLowerInvariant() } }
Export-CsvRows -Rows $productRows -OutputPath $ProductOutputPath
Write-Host "Saved sanitized product data to $ProductOutputPath"
