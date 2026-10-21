#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$WorkDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $WorkDir) { $WorkDir = (Get-Location).Path }

# Polskie znaki definiowane bezpiecznie dla PowerShell 5.1 (niezaleznie od kodowania pliku .ps1)
$S_OBWOD_PREFIX = "OBW$([char]0x00D3)D"
$S_OKREG_OBWOD  = "OKR$([char]0x0118)G/OBW$([char]0x00D3)D"
$S_SZKOLA_HDR   = "SZKO$([char]0x0141)A (SIEDZIBA)"
$S_SREDNIA      = "$([char]0x015A)REDNIA"
$S_SREDNIA_SZK  = "$([char]0x015A)REDNIA NA SZKO$([char]0x0141)$([char]0x0118)"
$S_UDZIAL_PCT   = "UDZIA$([char]0x0141) %"
$S_UDZIAL_ROCZ  = "Udzia$([char]0x0142) w roczniku %"
$S_WSZYSTKIE_SZ = "WSZYSTKIE OBWODY SZKOLNE (17 SZK$([char]0x00D3)$([char]0x0141))"
$S_OGOLEM       = "Og$([char]0x00F3)$([char]0x0142)em"
$S_OBWOD_COL    = "Obw$([char]0x00F3)d"
$S_SZKOLA_COL   = "Szko$([char]0x0142)a (siedziba)"
$S_KAT_PODM     = "Obwody podmiejskie (Che$([char]0x0142)miec / Pi$([char]0x0105)tkowa)"
$S_KAT_MIEJ     = "Szko$([char]0x0142)y miejskie (SP nr 1-21)"
$S_SP_PREFIX    = "^SZKO$([char]0x0141)A PODSTAWOWA NR\s+"
$S_SP_CHELMIEC  = "^SZKO$([char]0x0141)A PODSTAWOWA W\s+CHE$([char]0x0141)MCU$"
$S_SP_PIATKOWA  = "^SZKO$([char]0x0141)A PODSTAWOWA W\s+PI$([char]0x0104)TKOWEJ$"
$S_SHORT_CHELM  = "SP w Che$([char]0x0142)mcu"
$S_SHORT_PIATK  = "SP w Pi$([char]0x0105)tkowej"

function Read-SourceXlsx {
    param([string]$FilePath)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
    try {
        $ssEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' }
        $strings = @()
        if ($ssEntry) {
            $sr = New-Object System.IO.StreamReader($ssEntry.Open(), [System.Text.Encoding]::UTF8)
            [xml]$ssXml = $sr.ReadToEnd()
            $sr.Close()
            foreach ($si in $ssXml.sst.si) {
                if ($si.t -is [string]) { $strings += $si.t }
                elseif ($si.t.'#text') { $strings += $si.t.'#text' }
                else { $strings += $si.InnerText }
            }
        }
        $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/worksheets/sheet1.xml' }
        $sr = New-Object System.IO.StreamReader($sheetEntry.Open(), [System.Text.Encoding]::UTF8)
        [xml]$sheetXml = $sr.ReadToEnd()
        $sr.Close()

        $rows = @()
        foreach ($row in $sheetXml.worksheet.sheetData.row) {
            $vals = @()
            foreach ($c in $row.c) {
                if ($c.t -eq 's') { $vals += $strings[[int]$c.v] }
                elseif ($c.is) { $vals += $c.is.InnerText }
                else { $vals += $c.v }
            }
            $rows += ,$vals
        }
        return $rows
    }
    finally {
        $zip.Dispose()
    }
}

function ConvertTo-ExcelCol {
    param([int]$ColIndex1Based)
    [int]$dividend = $ColIndex1Based
    $colName = ''
    while ($dividend -gt 0) {
        [int]$modulo = ($dividend - 1) % 26
        $colName = [char]([int](65 + $modulo)) + $colName
        $dividend = [int][math]::Floor(($dividend - $modulo) / 26)
    }
    return $colName
}

function Escape-Xml {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-CellXml {
    param(
        [string]$ColLetter,
        [int]$RowIndex,
        $Value,
        [int]$StyleIndex = 0,
        [string]$Type = 'n',
        [string]$Formula = ''
    )
    $ref = "$ColLetter$RowIndex"
    if ($Type -eq 's') {
        $escaped = Escape-Xml ([string]$Value)
        return "<c r=`"$ref`" s=`"$StyleIndex`" t=`"inlineStr`"><is><t>$escaped</t></is></c>"
    }
    elseif ($Type -eq 'f') {
        $valStr = if ($null -ne $Value -and $Value -ne '') { "<v>$([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0}', $Value))</v>" } else { '' }
        return "<c r=`"$ref`" s=`"$StyleIndex`"><f>$Formula</f>$valStr</c>"
    }
    else {
        $numStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0}', $Value)
        return "<c r=`"$ref`" s=`"$StyleIndex`"><v>$numStr</v></c>"
    }
}

function Add-ZipTextEntry {
    param($ZipArchive, [string]$EntryPath, [string]$Content)
    $entry = $ZipArchive.CreateEntry($EntryPath)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $sw = New-Object System.IO.StreamWriter($entry.Open(), $utf8NoBom)
    $sw.Write($Content)
    $sw.Close()
}

Write-Host "Wczytywanie plikow rocznikowych .xlsx z katalogu: $WorkDir" -ForegroundColor Cyan

$yearFiles = Get-ChildItem -Path $WorkDir -Filter '*.xlsx' | Where-Object { $_.BaseName -match '^\d{4}$' } | Sort-Object { [int]$_.BaseName }

if ($yearFiles.Count -eq 0) {
    throw "Nie znaleziono plikow rocznikowych RRRR.xlsx (np. 2001.xlsx) w folderze $WorkDir"
}

$allYears = @()
$schoolsMap = [ordered]@{}
$allYearTotals = @{}

foreach ($file in $yearFiles) {
    $year = [int]$file.BaseName
    $allYears += $year
    $rows = Read-SourceXlsx -FilePath $file.FullName
    $yearSum = 0

    for ($i = 1; $i -lt $rows.Count; $i++) {
        $r = $rows[$i]
        if ($r.Count -lt 2) { continue }
        $obwodRaw = ([string]$r[0]).Trim()
        $countVal = 0
        [void][int]::TryParse(([string]$r[1]).Trim(), [ref]$countVal)
        $siedzibaRaw = if ($r.Count -ge 3 -and $null -ne $r[2]) { ([string]$r[2]).Trim() } else { '' }

        # Pomijamy wiersze techniczne OBWOD 100, OBWOD 101 oraz SUMA
        if ($obwodRaw -eq 'SUMA' -or $obwodRaw -eq "$S_OBWOD_PREFIX 100" -or $obwodRaw -eq "$S_OBWOD_PREFIX 101") {
            continue
        }
        if ($obwodRaw -notmatch "^$S_OBWOD_PREFIX\s+(\d+)$") {
            continue
        }
        $obwodNum = [int]$Matches[1]

        # Usuwamy prefiks "w " z nazwy szkoly (zgodnie z wymaganiem uzytkownika)
        $schoolClean = $siedzibaRaw -replace '^[wW]\s+', ''
        $schoolClean = $schoolClean.Trim()

        # Kategoria obwodu
        $category = if ($obwodNum -in @(22, 23)) { $S_KAT_PODM } else { $S_KAT_MIEJ }
        $shortCategory = if ($obwodNum -in @(22, 23)) { 'Podmiejskie' } else { 'Miejskie' }

        # Skrocona nazwa do czytelnych osi wykresow
        $shortName = $schoolClean -replace $S_SP_PREFIX, 'SP nr ' -replace $S_SP_CHELMIEC, $S_SHORT_CHELM -replace $S_SP_PIATKOWA, $S_SHORT_PIATK

        if (-not $schoolsMap.Contains($obwodRaw)) {
            $schoolsMap[$obwodRaw] = [PSCustomObject]@{
                Obwod         = $obwodRaw
                ObwodNum      = $obwodNum
                Siedziba      = $schoolClean
                ShortName     = $shortName
                Kategoria     = $category
                ShortCategory = $shortCategory
                ValuesByYear  = @{}
            }
        }
        $schoolsMap[$obwodRaw].ValuesByYear[$year] = $countVal
        $yearSum += $countVal
    }
    $allYearTotals[$year] = $yearSum
}

$schoolsList = @($schoolsMap.Values | Sort-Object ObwodNum)

function Build-DashboardAndExcelVariant {
    param(
        [int]$MinYear,
        [int]$MaxYear,
        [string]$ExcelFileName,
        [string[]]$HtmlFileNames
    )

    $years = @($allYears | Where-Object { $_ -ge $MinYear -and $_ -le $MaxYear })
    if ($years.Count -eq 0) {
        Write-Warning "Brak rocznikow w zakresie $MinYear-$MaxYear"
        return
    }

    $yearTotals = @{}
    foreach ($y in $years) {
        $yearTotals[$y] = $allYearTotals[$y]
    }
    $grandTotal = ($yearTotals.Values | Measure-Object -Sum).Sum

    $flatRecords = New-Object System.Collections.Generic.List[object]
    foreach ($y in $years) {
        $yTotal = $yearTotals[$y]
        foreach ($s in $schoolsList) {
            $val = if ($s.ValuesByYear.ContainsKey($y)) { $s.ValuesByYear[$y] } else { 0 }
            $share = if ($yTotal -gt 0) { [math]::Round(($val / $yTotal) * 100, 2) } else { 0 }
            $flatRecords.Add([PSCustomObject]@{
                Rocznik   = $y
                ObwodNum  = $s.ObwodNum
                Obwod     = $s.Obwod
                Szkola    = $s.Siedziba
                ShortName = $s.ShortName
                Kategoria = $s.Kategoria
                Liczba    = $val
                UdzialPct = $share
            })
        }
    }

    Write-Host ("Variant {0}-{1}: {2} rocznikow, {3} obwodow szkolnych, suma zameldowanych: {4}" -f $years[0], $years[-1], $years.Count, $schoolsList.Count, $grandTotal) -ForegroundColor Green

    # ==============================================================================
    # BUDOWA ZBIORCZEGO PLIKU EXCEL (.XLSX)
    # ==============================================================================
    $excelOutPath = Join-Path $WorkDir $ExcelFileName
    if (Test-Path $excelOutPath) { Remove-Item $excelOutPath -Force }

    # ARKUSZ 1: Macierz_Szkoly_Roczniki
    $s1Rows = New-Object System.Text.StringBuilder
    $lastYearColLetter = ConvertTo-ExcelCol (3 + $years.Count)
    $sumColLetter = ConvertTo-ExcelCol (3 + $years.Count + 1)
    $avgColLetter = ConvertTo-ExcelCol (3 + $years.Count + 2)
    $minColLetter = ConvertTo-ExcelCol (3 + $years.Count + 3)
    $maxColLetter = ConvertTo-ExcelCol (3 + $years.Count + 4)
    $pctColLetter = ConvertTo-ExcelCol (3 + $years.Count + 5)

    [void]$s1Rows.Append('<row r="1" ht="28" customHeight="1">')
    [void]$s1Rows.Append((New-CellXml -ColLetter 'A' -RowIndex 1 -Value $S_OKREG_OBWOD -StyleIndex 1 -Type 's'))
    [void]$s1Rows.Append((New-CellXml -ColLetter 'B' -RowIndex 1 -Value $S_SZKOLA_HDR -StyleIndex 1 -Type 's'))
    [void]$s1Rows.Append((New-CellXml -ColLetter 'C' -RowIndex 1 -Value 'KATEGORIA' -StyleIndex 1 -Type 's'))
    for ($yi = 0; $yi -lt $years.Count; $yi++) {
        $cLet = ConvertTo-ExcelCol (4 + $yi)
        [void]$s1Rows.Append((New-CellXml -ColLetter $cLet -RowIndex 1 -Value $years[$yi] -StyleIndex 1 -Type 'n'))
    }
    [void]$s1Rows.Append((New-CellXml -ColLetter $sumColLetter -RowIndex 1 -Value 'SUMA' -StyleIndex 2 -Type 's'))
    [void]$s1Rows.Append((New-CellXml -ColLetter $avgColLetter -RowIndex 1 -Value $S_SREDNIA -StyleIndex 2 -Type 's'))
    [void]$s1Rows.Append((New-CellXml -ColLetter $minColLetter -RowIndex 1 -Value 'MIN' -StyleIndex 2 -Type 's'))
    [void]$s1Rows.Append((New-CellXml -ColLetter $maxColLetter -RowIndex 1 -Value 'MAX' -StyleIndex 2 -Type 's'))
    [void]$s1Rows.Append((New-CellXml -ColLetter $pctColLetter -RowIndex 1 -Value $S_UDZIAL_PCT -StyleIndex 2 -Type 's'))
    [void]$s1Rows.Append('</row>')

    $rowIdx = 2
    $lastSchoolRow = 1 + $schoolsList.Count
    foreach ($s in $schoolsList) {
        [void]$s1Rows.Append("<row r=`"$rowIdx`" ht=`"20`" customHeight=`"1`">")
        [void]$s1Rows.Append((New-CellXml -ColLetter 'A' -RowIndex $rowIdx -Value $s.Obwod -StyleIndex 3 -Type 's'))
        [void]$s1Rows.Append((New-CellXml -ColLetter 'B' -RowIndex $rowIdx -Value $s.Siedziba -StyleIndex 4 -Type 's'))
        [void]$s1Rows.Append((New-CellXml -ColLetter 'C' -RowIndex $rowIdx -Value $s.ShortCategory -StyleIndex 3 -Type 's'))

        $vals = @()
        for ($yi = 0; $yi -lt $years.Count; $yi++) {
            $y = $years[$yi]
            $val = if ($s.ValuesByYear.ContainsKey($y)) { $s.ValuesByYear[$y] } else { 0 }
            $vals += $val
            $cLet = ConvertTo-ExcelCol (4 + $yi)
            [void]$s1Rows.Append((New-CellXml -ColLetter $cLet -RowIndex $rowIdx -Value $val -StyleIndex 5 -Type 'n'))
        }
        $rSum = ($vals | Measure-Object -Sum).Sum
        $rAvg = [math]::Round(($vals | Measure-Object -Average).Average, 1)
        $rMin = ($vals | Measure-Object -Minimum).Minimum
        $rMax = ($vals | Measure-Object -Maximum).Maximum
        $rPct = if ($grandTotal -gt 0) { [math]::Round($rSum / $grandTotal, 4) } else { 0 }

        $rangeRef = "D$rowIdx`:$lastYearColLetter$rowIdx"
        $sumRowIdx = $lastSchoolRow + 1
        [void]$s1Rows.Append((New-CellXml -ColLetter $sumColLetter -RowIndex $rowIdx -Value $rSum -StyleIndex 6 -Type 'f' -Formula "SUM($rangeRef)"))
        [void]$s1Rows.Append((New-CellXml -ColLetter $avgColLetter -RowIndex $rowIdx -Value $rAvg -StyleIndex 7 -Type 'f' -Formula "AVERAGE($rangeRef)"))
        [void]$s1Rows.Append((New-CellXml -ColLetter $minColLetter -RowIndex $rowIdx -Value $rMin -StyleIndex 5 -Type 'f' -Formula "MIN($rangeRef)"))
        [void]$s1Rows.Append((New-CellXml -ColLetter $maxColLetter -RowIndex $rowIdx -Value $rMax -StyleIndex 5 -Type 'f' -Formula "MAX($rangeRef)"))
        [void]$s1Rows.Append((New-CellXml -ColLetter $pctColLetter -RowIndex $rowIdx -Value $rPct -StyleIndex 8 -Type 'f' -Formula "$sumColLetter$rowIdx/`$$sumColLetter`$$sumRowIdx"))
        [void]$s1Rows.Append('</row>')
        $rowIdx++
    }

    $sumRowIdx = $lastSchoolRow + 1
    [void]$s1Rows.Append("<row r=`"$sumRowIdx`" ht=`"24`" customHeight=`"1`">")
    [void]$s1Rows.Append((New-CellXml -ColLetter 'A' -RowIndex $sumRowIdx -Value 'SUMA' -StyleIndex 9 -Type 's'))
    [void]$s1Rows.Append((New-CellXml -ColLetter 'B' -RowIndex $sumRowIdx -Value $S_WSZYSTKIE_SZ -StyleIndex 9 -Type 's'))
    [void]$s1Rows.Append((New-CellXml -ColLetter 'C' -RowIndex $sumRowIdx -Value $S_OGOLEM -StyleIndex 9 -Type 's'))
    for ($yi = 0; $yi -lt $years.Count; $yi++) {
        $y = $years[$yi]
        $cLet = ConvertTo-ExcelCol (4 + $yi)
        $yTotal = $yearTotals[$y]
        [void]$s1Rows.Append((New-CellXml -ColLetter $cLet -RowIndex $sumRowIdx -Value $yTotal -StyleIndex 10 -Type 'f' -Formula "SUM($cLet`2:$cLet$lastSchoolRow)"))
    }
    $avgYearTotal = [math]::Round(($yearTotals.Values | Measure-Object -Average).Average, 1)
    $minYearTotal = ($yearTotals.Values | Measure-Object -Minimum).Minimum
    $maxYearTotal = ($yearTotals.Values | Measure-Object -Maximum).Maximum
    $sumRangeYears = "D$sumRowIdx`:$lastYearColLetter$sumRowIdx"

    [void]$s1Rows.Append((New-CellXml -ColLetter $sumColLetter -RowIndex $sumRowIdx -Value $grandTotal -StyleIndex 10 -Type 'f' -Formula "SUM($sumColLetter`2:$sumColLetter$lastSchoolRow)"))
    [void]$s1Rows.Append((New-CellXml -ColLetter $avgColLetter -RowIndex $sumRowIdx -Value $avgYearTotal -StyleIndex 11 -Type 'f' -Formula "AVERAGE($sumRangeYears)"))
    [void]$s1Rows.Append((New-CellXml -ColLetter $minColLetter -RowIndex $sumRowIdx -Value $minYearTotal -StyleIndex 10 -Type 'f' -Formula "MIN($sumRangeYears)"))
    [void]$s1Rows.Append((New-CellXml -ColLetter $maxColLetter -RowIndex $sumRowIdx -Value $maxYearTotal -StyleIndex 10 -Type 'f' -Formula "MAX($sumRangeYears)"))
    [void]$s1Rows.Append((New-CellXml -ColLetter $pctColLetter -RowIndex $sumRowIdx -Value 1 -StyleIndex 12 -Type 'f' -Formula "SUM($pctColLetter`2:$pctColLetter$lastSchoolRow)"))
    [void]$s1Rows.Append('</row>')

    $s1ColsXml = @"
<cols>
  <col min="1" max="1" width="14" customWidth="1"/>
  <col min="2" max="2" width="36" customWidth="1"/>
  <col min="3" max="3" width="15" customWidth="1"/>
  <col min="4" max="$(3 + $years.Count)" width="8.5" customWidth="1"/>
  <col min="$(3 + $years.Count + 1)" max="$(3 + $years.Count + 5)" width="13" customWidth="1"/>
</cols>
"@

    $sheet1Xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane xSplit="3" ySplit="1" topLeftCell="D2" activePane="bottomRight" state="frozen"/>
      <selection pane="bottomRight" activeCell="D2" sqref="D2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  $s1ColsXml
  <sheetData>
    $($s1Rows.ToString())
  </sheetData>
  <autoFilter ref="A1:$pctColLetter$lastSchoolRow"/>
  <conditionalFormatting sqref="D2:$lastYearColLetter$lastSchoolRow">
    <cfRule type="colorScale" priority="1">
      <colorScale>
        <cfvo type="min"/>
        <cfvo type="percentile" val="50"/>
        <cfvo type="max"/>
        <color rgb="FFF8FAFC"/>
        <color rgb="FF93C5FD"/>
        <color rgb="FF1D4ED8"/>
      </colorScale>
    </cfRule>
  </conditionalFormatting>
</worksheet>
"@

    # ARKUSZ 2: Roczniki_Szkoly
    $s2Rows = New-Object System.Text.StringBuilder
    $lastSchoolColS2 = ConvertTo-ExcelCol (1 + $schoolsList.Count)
    $s2SumCol = ConvertTo-ExcelCol (2 + $schoolsList.Count)
    $s2AvgCol = ConvertTo-ExcelCol (3 + $schoolsList.Count)
    $s2MinCol = ConvertTo-ExcelCol (4 + $schoolsList.Count)
    $s2MaxCol = ConvertTo-ExcelCol (5 + $schoolsList.Count)

    [void]$s2Rows.Append('<row r="1" ht="30" customHeight="1">')
    [void]$s2Rows.Append((New-CellXml -ColLetter 'A' -RowIndex 1 -Value 'ROCZNIK' -StyleIndex 1 -Type 's'))
    for ($si = 0; $si -lt $schoolsList.Count; $si++) {
        $cLet = ConvertTo-ExcelCol (2 + $si)
        $headerLabel = "$($schoolsList[$si].ShortName) ($($schoolsList[$si].Obwod))"
        [void]$s2Rows.Append((New-CellXml -ColLetter $cLet -RowIndex 1 -Value $headerLabel -StyleIndex 1 -Type 's'))
    }
    [void]$s2Rows.Append((New-CellXml -ColLetter $s2SumCol -RowIndex 1 -Value 'SUMA ROCZNIKA' -StyleIndex 2 -Type 's'))
    [void]$s2Rows.Append((New-CellXml -ColLetter $s2AvgCol -RowIndex 1 -Value $S_SREDNIA_SZK -StyleIndex 2 -Type 's'))
    [void]$s2Rows.Append((New-CellXml -ColLetter $s2MinCol -RowIndex 1 -Value 'MIN' -StyleIndex 2 -Type 's'))
    [void]$s2Rows.Append((New-CellXml -ColLetter $s2MaxCol -RowIndex 1 -Value 'MAX' -StyleIndex 2 -Type 's'))
    [void]$s2Rows.Append('</row>')

    $rIdx2 = 2
    $lastYearRowS2 = 1 + $years.Count
    foreach ($y in $years) {
        [void]$s2Rows.Append("<row r=`"$rIdx2`" ht=`"20`" customHeight=`"1`">")
        [void]$s2Rows.Append((New-CellXml -ColLetter 'A' -RowIndex $rIdx2 -Value $y -StyleIndex 3 -Type 'n'))
        $valsY = @()
        for ($si = 0; $si -lt $schoolsList.Count; $si++) {
            $s = $schoolsList[$si]
            $val = if ($s.ValuesByYear.ContainsKey($y)) { $s.ValuesByYear[$y] } else { 0 }
            $valsY += $val
            $cLet = ConvertTo-ExcelCol (2 + $si)
            [void]$s2Rows.Append((New-CellXml -ColLetter $cLet -RowIndex $rIdx2 -Value $val -StyleIndex 5 -Type 'n'))
        }
        $ySum = ($valsY | Measure-Object -Sum).Sum
        $yAvg = [math]::Round(($valsY | Measure-Object -Average).Average, 1)
        $yMin = ($valsY | Measure-Object -Minimum).Minimum
        $yMax = ($valsY | Measure-Object -Maximum).Maximum
        $rowRangeS2 = "B$rIdx2`:$lastSchoolColS2$rIdx2"

        [void]$s2Rows.Append((New-CellXml -ColLetter $s2SumCol -RowIndex $rIdx2 -Value $ySum -StyleIndex 6 -Type 'f' -Formula "SUM($rowRangeS2)"))
        [void]$s2Rows.Append((New-CellXml -ColLetter $s2AvgCol -RowIndex $rIdx2 -Value $yAvg -StyleIndex 7 -Type 'f' -Formula "AVERAGE($rowRangeS2)"))
        [void]$s2Rows.Append((New-CellXml -ColLetter $s2MinCol -RowIndex $rIdx2 -Value $yMin -StyleIndex 5 -Type 'f' -Formula "MIN($rowRangeS2)"))
        [void]$s2Rows.Append((New-CellXml -ColLetter $s2MaxCol -RowIndex $rIdx2 -Value $yMax -StyleIndex 5 -Type 'f' -Formula "MAX($rowRangeS2)"))
        [void]$s2Rows.Append('</row>')
        $rIdx2++
    }

    $sumRowS2 = $lastYearRowS2 + 1
    [void]$s2Rows.Append("<row r=`"$sumRowS2`" ht=`"24`" customHeight=`"1`">")
    [void]$s2Rows.Append((New-CellXml -ColLetter 'A' -RowIndex $sumRowS2 -Value 'SUMA' -StyleIndex 9 -Type 's'))
    for ($si = 0; $si -lt $schoolsList.Count; $si++) {
        $s = $schoolsList[$si]
        $cLet = ConvertTo-ExcelCol (2 + $si)
        $sTotal = 0
        foreach ($y in $years) {
            if ($s.ValuesByYear.ContainsKey($y)) { $sTotal += $s.ValuesByYear[$y] }
        }
        [void]$s2Rows.Append((New-CellXml -ColLetter $cLet -RowIndex $sumRowS2 -Value $sTotal -StyleIndex 10 -Type 'f' -Formula "SUM($cLet`2:$cLet$lastYearRowS2)"))
    }
    [void]$s2Rows.Append((New-CellXml -ColLetter $s2SumCol -RowIndex $sumRowS2 -Value $grandTotal -StyleIndex 10 -Type 'f' -Formula "SUM($s2SumCol`2:$s2SumCol$lastYearRowS2)"))
    [void]$s2Rows.Append((New-CellXml -ColLetter $s2AvgCol -RowIndex $sumRowS2 -Value ([math]::Round($grandTotal / $years.Count, 1)) -StyleIndex 11 -Type 'f' -Formula "AVERAGE($s2SumCol`2:$s2SumCol$lastYearRowS2)"))
    [void]$s2Rows.Append((New-CellXml -ColLetter $s2MinCol -RowIndex $sumRowS2 -Value $minYearTotal -StyleIndex 10 -Type 'f' -Formula "MIN($s2SumCol`2:$s2SumCol$lastYearRowS2)"))
    [void]$s2Rows.Append((New-CellXml -ColLetter $s2MaxCol -RowIndex $sumRowS2 -Value $maxYearTotal -StyleIndex 10 -Type 'f' -Formula "MAX($s2SumCol`2:$s2SumCol$lastYearRowS2)"))
    [void]$s2Rows.Append('</row>')

    $sheet2Xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
      <selection pane="bottomRight" activeCell="B2" sqref="B2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="12" customWidth="1"/>
    <col min="2" max="$(1 + $schoolsList.Count)" width="18" customWidth="1"/>
    <col min="$(2 + $schoolsList.Count)" max="$(5 + $schoolsList.Count)" width="16" customWidth="1"/>
  </cols>
  <sheetData>
    $($s2Rows.ToString())
  </sheetData>
  <autoFilter ref="A1:$s2MaxCol$lastYearRowS2"/>
</worksheet>
"@

    # ARKUSZ 3: Baza_Danych_Plaska
    $s3Rows = New-Object System.Text.StringBuilder
    [void]$s3Rows.Append('<row r="1" ht="26" customHeight="1">')
    [void]$s3Rows.Append((New-CellXml -ColLetter 'A' -RowIndex 1 -Value 'Rocznik' -StyleIndex 1 -Type 's'))
    [void]$s3Rows.Append((New-CellXml -ColLetter 'B' -RowIndex 1 -Value 'Nr obwodu' -StyleIndex 1 -Type 's'))
    [void]$s3Rows.Append((New-CellXml -ColLetter 'C' -RowIndex 1 -Value $S_OBWOD_COL -StyleIndex 1 -Type 's'))
    [void]$s3Rows.Append((New-CellXml -ColLetter 'D' -RowIndex 1 -Value $S_SZKOLA_COL -StyleIndex 1 -Type 's'))
    [void]$s3Rows.Append((New-CellXml -ColLetter 'E' -RowIndex 1 -Value 'Kategoria' -StyleIndex 1 -Type 's'))
    [void]$s3Rows.Append((New-CellXml -ColLetter 'F' -RowIndex 1 -Value 'Liczba zameldowanych' -StyleIndex 2 -Type 's'))
    [void]$s3Rows.Append((New-CellXml -ColLetter 'G' -RowIndex 1 -Value $S_UDZIAL_ROCZ -StyleIndex 2 -Type 's'))
    [void]$s3Rows.Append('</row>')

    $rIdx3 = 2
    foreach ($rec in $flatRecords) {
        [void]$s3Rows.Append("<row r=`"$rIdx3`" ht=`"19`">")
        [void]$s3Rows.Append((New-CellXml -ColLetter 'A' -RowIndex $rIdx3 -Value $rec.Rocznik -StyleIndex 3 -Type 'n'))
        [void]$s3Rows.Append((New-CellXml -ColLetter 'B' -RowIndex $rIdx3 -Value $rec.ObwodNum -StyleIndex 3 -Type 'n'))
        [void]$s3Rows.Append((New-CellXml -ColLetter 'C' -RowIndex $rIdx3 -Value $rec.Obwod -StyleIndex 3 -Type 's'))
        [void]$s3Rows.Append((New-CellXml -ColLetter 'D' -RowIndex $rIdx3 -Value $rec.Szkola -StyleIndex 4 -Type 's'))
        [void]$s3Rows.Append((New-CellXml -ColLetter 'E' -RowIndex $rIdx3 -Value $rec.Kategoria -StyleIndex 4 -Type 's'))
        [void]$s3Rows.Append((New-CellXml -ColLetter 'F' -RowIndex $rIdx3 -Value $rec.Liczba -StyleIndex 5 -Type 'n'))
        $pctFrac = [math]::Round($rec.UdzialPct / 100.0, 4)
        [void]$s3Rows.Append((New-CellXml -ColLetter 'G' -RowIndex $rIdx3 -Value $pctFrac -StyleIndex 8 -Type 'n'))
        [void]$s3Rows.Append('</row>')
        $rIdx3++
    }
    $lastRowS3 = $rIdx3 - 1

    $sheet3Xml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="A2" sqref="A2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="2" width="12" customWidth="1"/>
    <col min="3" max="3" width="14" customWidth="1"/>
    <col min="4" max="4" width="36" customWidth="1"/>
    <col min="5" max="5" width="38" customWidth="1"/>
    <col min="6" max="7" width="22" customWidth="1"/>
  </cols>
  <sheetData>
    $($s3Rows.ToString())
  </sheetData>
  <autoFilter ref="A1:G$lastRowS3"/>
</worksheet>
"@

    $stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="#,##0"/>
    <numFmt numFmtId="165" formatCode="#,##0.0"/>
    <numFmt numFmtId="166" formatCode="0.00%"/>
  </numFmts>
  <fonts count="4">
    <font><sz val="10"/><color rgb="FF0F172A"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FF0F172A"/><name val="Segoe UI"/></font>
    <font><b/><sz val="10"/><color rgb="FF1E3A8A"/><name val="Segoe UI"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1D4ED8"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFE2E8F0"/></left>
      <right style="thin"><color rgb="FFE2E8F0"/></right>
      <top style="thin"><color rgb="FFE2E8F0"/></top>
      <bottom style="thin"><color rgb="FFE2E8F0"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF94A3B8"/></left>
      <right style="thin"><color rgb="FF94A3B8"/></right>
      <top style="medium"><color rgb="FF1E293B"/></top>
      <bottom style="double"><color rgb="FF1E293B"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="13">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="164" fontId="3" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="165" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="166" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="5" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="left" vertical="center"/>
    </xf>
    <xf numFmtId="164" fontId="2" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="165" fontId="2" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
    <xf numFmtId="166" fontId="2" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="right" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>
'@

    $workbookXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Macierz_Szkoly_Roczniki" sheetId="1" r:id="rId1"/>
    <sheet name="Roczniki_Szkoly" sheetId="2" r:id="rId2"/>
    <sheet name="Baza_Danych_Plaska" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>
'@

    $workbookRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@

    $rootRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

    $contentTypesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@

    $fs = New-Object System.IO.FileStream($excelOutPath, [System.IO.FileMode]::Create)
    $zipOut = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipTextEntry -ZipArchive $zipOut -EntryPath '[Content_Types].xml' -Content $contentTypesXml
        Add-ZipTextEntry -ZipArchive $zipOut -EntryPath '_rels/.rels' -Content $rootRelsXml
        Add-ZipTextEntry -ZipArchive $zipOut -EntryPath 'xl/workbook.xml' -Content $workbookXml
        Add-ZipTextEntry -ZipArchive $zipOut -EntryPath 'xl/_rels/workbook.xml.rels' -Content $workbookRelsXml
        Add-ZipTextEntry -ZipArchive $zipOut -EntryPath 'xl/styles.xml' -Content $stylesXml
        Add-ZipTextEntry -ZipArchive $zipOut -EntryPath 'xl/worksheets/sheet1.xml' -Content $sheet1Xml
        Add-ZipTextEntry -ZipArchive $zipOut -EntryPath 'xl/worksheets/sheet2.xml' -Content $sheet2Xml
        Add-ZipTextEntry -ZipArchive $zipOut -EntryPath 'xl/worksheets/sheet3.xml' -Content $sheet3Xml
    }
    finally {
        $zipOut.Dispose()
        $fs.Dispose()
    }

    Write-Host "Wygenerowano zbiorczy plik Excel: $excelOutPath" -ForegroundColor Green

    # ==============================================================================
    # PRZYGOTOWANIE PAYLOADU JSON I AKTUALIZACJA DASHBOARDU HTML
    # ==============================================================================
    $schoolsPayload = @()
    foreach ($s in $schoolsList) {
        $valuesMap = [ordered]@{}
        foreach ($y in $years) {
            $valuesMap["$y"] = if ($s.ValuesByYear.ContainsKey($y)) { $s.ValuesByYear[$y] } else { 0 }
        }
        $schoolsPayload += [PSCustomObject]@{
            id            = $s.Obwod
            num           = $s.ObwodNum
            name          = $s.Siedziba
            shortName     = $s.ShortName
            category      = $s.Kategoria
            shortCategory = $s.ShortCategory
            values        = $valuesMap
        }
    }

    $payloadObj = [PSCustomObject]@{
        generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        years       = $years
        schools     = $schoolsPayload
    }

    $jsonString = $payloadObj | ConvertTo-Json -Depth 6 -Compress

    foreach ($htmlFile in $HtmlFileNames) {
        $htmlPath = Join-Path $WorkDir $htmlFile
        if (Test-Path $htmlPath) {
            $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
            $htmlContent = [System.IO.File]::ReadAllText($htmlPath, $utf8NoBom)
            $pattern = '(?s)/\*__DATA_START__\*/.*?/\*__DATA_END__\*/'
            $replacement = "/*__DATA_START__*/const INITIAL_DATA = $jsonString;/*__DATA_END__*/"
            if ($htmlContent -match $pattern) {
                $updatedHtml = [regex]::Replace($htmlContent, $pattern, $replacement)
                [System.IO.File]::WriteAllText($htmlPath, $updatedHtml, $utf8NoBom)
                Write-Host "Zaktualizowano dane w pliku HTML: $htmlPath" -ForegroundColor Green
            }
        }
    }
}

# 1. Pełny zakres roczników (2001–2026)
Build-DashboardAndExcelVariant `
    -MinYear 2001 `
    -MaxYear 9999 `
    -ExcelFileName "zbiorczy_raport_2001_2026.xlsx" `
    -HtmlFileNames @("dashboard_zameldowani_2001_2026.html")

# 2. Zakres roczników 2021–2026 (również jako główny index.html dla GitHub Pages)
Build-DashboardAndExcelVariant `
    -MinYear 2021 `
    -MaxYear 2026 `
    -ExcelFileName "zbiorczy_raport_2021_2026.xlsx" `
    -HtmlFileNames @("dashboard_zameldowani_2021_2026.html", "index.html")
