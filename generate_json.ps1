[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$cleaned = Import-Csv -Path .\cleaned_sales_data.csv -Encoding utf8
Write-Host "Total cleaned rows: $($cleaned.Count)"

# Channel stats
$channels = $cleaned | Group-Object 'ช่องทางการขาย' | Select-Object Name, Count
# Region stats
$regions = $cleaned | Group-Object 'ภาค' | Select-Object Name, Count
# Payment stats
$payments = $cleaned | Group-Object 'วิธีชำระเงิน' | Select-Object Name, Count
# Status stats
$statuses = $cleaned | Group-Object 'สถานะ' | Select-Object Name, Count

# Top products variance
$prodGroup = $cleaned | Group-Object 'ชื่อสินค้า' | ForEach-Object {
    $prices = $_.Group | ForEach-Object { [double]$_.'ราคาต่อหน่วย' } | Select-Object -Unique
    [PSCustomObject]@{
        Name = $_.Name
        Count = $_.Count
        UniquePrices = $prices.Count
        Min = ($prices | Measure-Object -Minimum).Minimum
        Max = ($prices | Measure-Object -Maximum).Maximum
        Base = $_.Group[0].'ราคาฐานมาตรฐาน'
    }
} | Sort-Object UniquePrices -Descending | Select-Object -First 10

# Export summary JSON
$summaryData = [ordered]@{
    totalOriginal = 10000
    totalCleaned = $cleaned.Count
    droppedGender = 197
    fixedAyutthaya = 264
    fixedPosCod = 396
    fixedCancelled = 456
    flaggedMinorCc = 234
    adjustedAge = 4379
    channels = $channels
    regions = $regions
    payments = $payments
    statuses = $statuses
    topVarianceProducts = $prodGroup
    sampleRows = ($cleaned | Select-Object -First 150)
}

$summaryData | ConvertTo-Json -Depth 5 | Out-File -FilePath .\dashboard_data.json -Encoding utf8
Write-Host "Generated dashboard_data.json successfully!"
