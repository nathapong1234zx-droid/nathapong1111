[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$cleaned = Import-Csv -Path .\cleaned_sales_data.csv -Encoding utf8
Write-Host "Cleaned rows count: $($cleaned.Count)"
Write-Host "Columns: $(($cleaned[0].PSObject.Properties.Name) -join ', ')"
Write-Host "Check any 'ไม่ระบุ' sex: $(($cleaned | Where-Object { $_.'เพศ' -eq 'ไม่ระบุ' }).Count)"
Write-Host "Check any 'อยุธยา' province: $(($cleaned | Where-Object { $_.'จังหวัด' -eq 'อยุธยา' }).Count)"
Write-Host "Check any POS with COD: $(($cleaned | Where-Object { $_.'ช่องทางการขาย' -eq 'หน้าร้าน' -and $_.'วิธีชำระเงิน' -eq 'เก็บเงินปลายทาง' }).Count)"
Write-Host "Check cancelled with ship fee > 0: $(($cleaned | Where-Object { $_.'สถานะ' -eq 'ยกเลิก' -and [double]$_.'ค่าจัดส่ง' -gt 0 }).Count)"
Write-Host "Check cancelled with ship days > 0: $(($cleaned | Where-Object { $_.'สถานะ' -eq 'ยกเลิก' -and -not [string]::IsNullOrWhiteSpace($_.'จำนวนวันจัดส่ง') }).Count)"

Write-Host "`nSample Row 1:"
$cleaned[0] | Format-List

