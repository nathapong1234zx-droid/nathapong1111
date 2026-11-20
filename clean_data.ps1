[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$inputPath = ".\sales_data.csv"
$outputPath = ".\cleaned_sales_data.csv"

Write-Host "Reading $inputPath ..."
$rows = Import-Csv -Path $inputPath -Encoding UTF8
$totalInitial = $rows.Count
Write-Host "Total rows loaded: $totalInitial"

# 1. Customer baseline age tracking (Issue: age doesn't increase over years)
# For each customer, find their minimum order year and baseline age at that year.
$custBase = @{}
foreach ($r in $rows) {
    $cId = $r.'รหัสลูกค้า'
    $yr = [int]($r.'วันที่สั่งซื้อ'.Substring(0, 4))
    $age = [int]$r.'อายุ'
    if (-not $custBase.ContainsKey($cId)) {
        $custBase[$cId] = @{ MinYear = $yr; BaseAge = $age }
    } else {
        if ($yr -lt $custBase[$cId].MinYear) {
            $custBase[$cId].MinYear = $yr
            $custBase[$cId].BaseAge = $age
        }
    }
}

# 2. Base price calculation per product (Issue: multiple prices per product)
# Calculate Median price for each product
$productPrices = @{}
foreach ($r in $rows) {
    $pName = $r.'ชื่อสินค้า'
    $price = [double]$r.'ราคาต่อหน่วย'
    if (-not $productPrices.ContainsKey($pName)) {
        $productPrices[$pName] = [System.Collections.Generic.List[double]]::new()
    }
    $productPrices[$pName].Add($price)
}

$productMedianPrice = @{}
foreach ($pName in $productPrices.Keys) {
    $list = $productPrices[$pName]
    $sorted = $list | Sort-Object
    $mid = [math]::Floor($sorted.Count / 2)
    if ($sorted.Count % 2 -eq 0) {
        $median = ($sorted[$mid - 1] + $sorted[$mid]) / 2.0
    } else {
        $median = $sorted[$mid]
    }
    $productMedianPrice[$pName] = [math]::Round($median, 2)
}

# Counters for reporting
$countGenderDropped = 0
$countAyutthayaFixed = 0
$countPosCodFixed = 0
$countCancelledOrdersReset = 0
$countMinorCcFlagged = 0
$countAgeAdjusted = 0

$cleanedRows = [System.Collections.Generic.List[PSObject]]::new()

foreach ($r in $rows) {
    # Issue 1: Gender filter (Keep only ชาย and หญิง)
    if ($r.'เพศ' -ne 'ชาย' -and $r.'เพศ' -ne 'หญิง') {
        $countGenderDropped++
        continue
    }

    # Issue 3: Standardize province name
    if ($r.'จังหวัด' -eq 'อยุธยา') {
        $r.'จังหวัด' = 'พระนครศรีอยุธยา'
        $countAyutthayaFixed++
    }

    # Issue 8: Dynamic age adjustment based on purchase year
    $orderYear = [int]($r.'วันที่สั่งซื้อ'.Substring(0, 4))
    $cId = $r.'รหัสลูกค้า'
    $expectedAge = $custBase[$cId].BaseAge + ($orderYear - $custBase[$cId].MinYear)
    if ([int]$r.'อายุ' -ne $expectedAge) {
        $r.'อายุ' = [string]$expectedAge
        $countAgeAdjusted++
    }

    # Issue 5: POS with COD conflict
    # If POS and COD, adjust to 'เงินสด'
    if ($r.'ช่องทางการขาย' -eq 'หน้าร้าน' -and $r.'วิธีชำระเงิน' -eq 'เก็บเงินปลายทาง') {
        $r.'วิธีชำระเงิน' = 'เงินสด'
        $countPosCodFixed++
    }

    # Issue 6: Cancelled order logic
    if ($r.'สถานะ' -eq 'ยกเลิก') {
        $hadShipFee = ([double]$r.'ค่าจัดส่ง' -gt 0)
        $hadShipDays = (-not [string]::IsNullOrWhiteSpace($r.'จำนวนวันจัดส่ง'))
        if ($hadShipFee -or $hadShipDays) {
            $countCancelledOrdersReset++
        }
        $r.'ค่าจัดส่ง' = "0"
        $r.'จำนวนวันจัดส่ง' = ""
        # Recalculate total net = pre-shipping (as shipping fee is 0)
        $preShipping = [double]$r.'ยอดก่อนค่าส่ง'
        $r.'ยอดรวมสุทธิ' = [string]$preShipping
    }

    # Issue 7: Minor (15-17) with Credit card
    $age = [int]$r.'อายุ'
    $minorFlag = "ปกติ"
    if ($age -ge 15 -and $age -le 17 -and $r.'วิธีชำระเงิน' -eq 'บัตรเครดิต') {
        $minorFlag = "บัตรเสริม/ผู้ปกครองยินยอม"
        $countMinorCcFlagged++
    }

    # Issue 9: Product Price consistency
    $pName = $r.'ชื่อสินค้า'
    $basePrice = $productMedianPrice[$pName]
    $unitPrice = [double]$r.'ราคาต่อหน่วย'
    $priceDiff = [math]::Round($unitPrice - $basePrice, 2)

    # Issue 4: Missing review scores remain empty string / blank

    # Create cleaned object with all original fields + audit flags
    $rowObj = [ordered]@{
        'รหัสคำสั่งซื้อ'      = $r.'รหัสคำสั่งซื้อ'
        'วันที่สั่งซื้อ'       = $r.'วันที่สั่งซื้อ'
        'เวลาสั่งซื้อ'        = $r.'เวลาสั่งซื้อ'
        'รหัสลูกค้า'        = $r.'รหัสลูกค้า'
        'เพศ'             = $r.'เพศ'
        'อายุ'             = $r.'อายุ'
        'จังหวัด'          = $r.'จังหวัด'
        'ภาค'             = $r.'ภาค'
        'ช่องทางการขาย'     = $r.'ช่องทางการขาย'
        'หมวดสินค้า'        = $r.'หมวดสินค้า'
        'ชื่อสินค้า'         = $r.'ชื่อสินค้า'
        'ราคาต่อหน่วย'      = $r.'ราคาต่อหน่วย'
        'ราคาฐานมาตรฐาน'    = [string]$basePrice
        'ส่วนต่างราคา'       = [string]$priceDiff
        'จำนวน'           = $r.'จำนวน'
        'ส่วนลด(%)'        = $r.'ส่วนลด(%)'
        'ยอดก่อนค่าส่ง'     = $r.'ยอดก่อนค่าส่ง'
        'ค่าจัดส่ง'         = $r.'ค่าจัดส่ง'
        'ยอดรวมสุทธิ'       = $r.'ยอดรวมสุทธิ'
        'วิธีชำระเงิน'       = $r.'วิธีชำระเงิน'
        'หมายเหตุการชำระเงิน' = $minorFlag
        'สถานะ'           = $r.'สถานะ'
        'จำนวนวันจัดส่ง'    = $r.'จำนวนวันจัดส่ง'
        'คะแนนรีวิว'       = $r.'คะแนนรีวิว'
    }

    $cleanedRows.Add([PSCustomObject]$rowObj)
}

Write-Host "Exporting to $outputPath ..."
$cleanedRows | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "--- SUMMARY OF DATA CLEANSING ---"
Write-Host "Total original rows: $totalInitial"
Write-Host "Rows after dropping 'ไม่ระบุ' gender: $($cleanedRows.Count) (Dropped: $countGenderDropped)"
Write-Host "Ayutthaya renamed to พระนครศรีอยุธยา: $countAyutthayaFixed"
Write-Host "POS with COD adjusted to เงินสด: $countPosCodFixed"
Write-Host "Cancelled orders with fee/days reset: $countCancelledOrdersReset"
Write-Host "Minor credit card transactions flagged: $countMinorCcFlagged"
Write-Host "Customer ages adjusted for year progression: $countAgeAdjusted"
Write-Host "Finished successfully!"
