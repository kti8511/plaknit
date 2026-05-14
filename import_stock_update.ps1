param(
  [Parameter(Position = 0)]
  [string]$StockFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataPath = Join-Path $repoRoot 'data.json'
$unmatchedPath = Join-Path $repoRoot 'stock_unmatched.json'
$desktop = [Environment]::GetFolderPath('Desktop')

function New-KoreanString([int[]]$codePoints) {
  return -join ($codePoints | ForEach-Object { [char]$_ })
}

$HDR_BARCODE = New-KoreanString @(0xBC14,0xCF54,0xB4DC) # header: barcode
$HDR_NAME = New-KoreanString @(0xCD9C,0xACE0,0xC0C1,0xD488,0xBA85) # header: outbound product name
$HDR_STOCK_OUTBOUND = New-KoreanString @(0xCD9C,0xACE0,0xAC00,0xB2A5) # header: outbound available stock
$HDR_STOCK_CURRENT = New-KoreanString @(0xD604,0xC7AC,0xACE0) # header: current stock
$HDR_STOCK_TOTAL = New-KoreanString @(0xCD1D,0xC7AC,0xACE0) # header: total stock

function Normalize([object]$value) {
  $text = ''
  if ($null -ne $value) { $text = [string]$value }
  return ([regex]::Replace($text, '\s+', '')).ToUpperInvariant()
}

function Normalize-Name([object]$value) {
  $text = Normalize $value
  $text = [regex]::Replace($text, '^\[[^\]]+\]', '')
  return [regex]::Replace($text, '[-_()\[\]/,]', '')
}

function Get-LatestStockFile {
  $patterns = @(
    '재고조회(기본)_*.xlsx',
    '재고조회_*.xlsx',
    '*재고조회*.xlsx'
  )
  $files = foreach ($pattern in $patterns) {
    Get-ChildItem -LiteralPath $desktop -Filter $pattern -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -notlike '~$*' }
  }
  if (-not $files) {
    throw "재고조회(기본) 엑셀 파일을 바탕화면에서 찾지 못했습니다."
  }
  return ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}

function Get-HeaderIndexMap([object[,]]$values, [int]$colCount) {
  $idx = @{}
  for ($c = 1; $c -le $colCount; $c++) {
    $h = ''
    if ($null -ne $values[1, $c]) { $h = [string]$values[1, $c] }
    $h = $h.Trim()
    if ($h -and -not $idx.ContainsKey($h)) { $idx[$h] = $c }
  }
  return $idx
}

function Load-StockRows([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) {
    throw "엑셀 파일 경로가 존재하지 않습니다: $path"
  }

  Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

  function Get-ColIndex([string]$cellRef) {
    $letters = ($cellRef -replace '[0-9]', '')
    $col = 0
    foreach ($ch in $letters.ToCharArray()) {
      $col = ($col * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $col
  }

  function Get-XmlText([object]$node) {
    if ($null -eq $node) { return '' }
    if ($node -is [System.Xml.XmlNode]) { return [string]$node.InnerText }
    return [string]$node
  }

  function Get-CellText($cell, [string[]]$sharedStrings) {
    if (-not $cell) { return '' }
    $t = $cell.t
    if ($t -eq 'inlineStr') {
      if ($cell.is -and $cell.is.t) { return (Get-XmlText $cell.is.t) }
      return ''
    }
    if (-not $cell.v) { return '' }
    $raw = Get-XmlText $cell.v
    if ($t -eq 's') {
      $idx = 0
      if ([int]::TryParse($raw, [ref]$idx) -and $idx -ge 0 -and $idx -lt $sharedStrings.Count) {
        return $sharedStrings[$idx]
      }
      return ''
    }
    return $raw
  }

  $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
  try {
    $sharedStrings = @()
    $sharedEntry = $zip.GetEntry('xl/sharedStrings.xml')
    if ($sharedEntry) {
      $sr = New-Object System.IO.StreamReader($sharedEntry.Open())
      try {
        $xmlText = $sr.ReadToEnd()
      } finally { $sr.Dispose() }
      [xml]$sharedXml = $xmlText
      $sis = @($sharedXml.sst.si)
      $sharedStrings = foreach ($si in $sis) {
        if ($si.t) { Get-XmlText $si.t }
        elseif ($si.r) { -join ($si.r | ForEach-Object { Get-XmlText $_.t }) }
        else { '' }
      }
      $sharedStrings = @($sharedStrings)
    }

    $sheetEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
    if (-not $sheetEntry) { throw 'sheet1.xml을 찾지 못했습니다.' }
    $sr2 = New-Object System.IO.StreamReader($sheetEntry.Open())
    try { $sheetText = $sr2.ReadToEnd() } finally { $sr2.Dispose() }
    [xml]$sheetXml = $sheetText
    $rowsXml = @($sheetXml.worksheet.sheetData.row)
    if (-not $rowsXml) { return @() }

    $headerRow = ($rowsXml | Where-Object { $_.r -eq '1' } | Select-Object -First 1)
    if (-not $headerRow) { throw '헤더(1행)를 찾지 못했습니다.' }

    $headerMap = @{}
    foreach ($cell in @($headerRow.c)) {
      $col = Get-ColIndex $cell.r
      $txt = (Get-CellText $cell $sharedStrings).Trim()
      if ($txt -and -not $headerMap.ContainsKey($txt)) { $headerMap[$txt] = $col }
    }

    $barcodeCol = $headerMap[$HDR_BARCODE]
    $nameCol = $headerMap[$HDR_NAME]
    $stockCol = $null
    foreach ($candidate in @($HDR_STOCK_OUTBOUND, $HDR_STOCK_CURRENT, $HDR_STOCK_TOTAL)) {
      if ($headerMap.ContainsKey($candidate)) { $stockCol = $headerMap[$candidate]; break }
    }

    $missing = @()
    if (-not $nameCol) { $missing += $HDR_NAME }
    if (-not $barcodeCol) { $missing += $HDR_BARCODE }
    if (-not $stockCol) { $missing += ($HDR_STOCK_OUTBOUND + '/' + $HDR_STOCK_CURRENT + '/' + $HDR_STOCK_TOTAL) }
    if ($missing.Count -gt 0) {
      throw ("필수 컬럼 누락: " + ($missing -join ', '))
    }

    $outRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rowsXml) {
      if ([int]$row.r -lt 2) { continue }

      $cellByCol = @{}
      foreach ($cell in @($row.c)) {
        $cellByCol[(Get-ColIndex $cell.r)] = $cell
      }

      $barcode = (Get-CellText $cellByCol[$barcodeCol] $sharedStrings).Trim()
      $outboundName = (Get-CellText $cellByCol[$nameCol] $sharedStrings).Trim()
      $stockRaw = (Get-CellText $cellByCol[$stockCol] $sharedStrings).Trim()

      if (-not $barcode -and -not $outboundName) { continue }

      $stockQty = 0
      try {
        if ($stockRaw) { $stockQty = [int][math]::Floor([double]($stockRaw -replace ',', '')) }
      } catch { $stockQty = 0 }

      $outRows.Add([pscustomobject]@{
        barcode       = $barcode
        outbound_name = $outboundName
        stock_qty     = $stockQty
      })
    }

    return ,$outRows.ToArray()
  } finally {
    $zip.Dispose()
  }
}

if (-not $StockFile) {
  $StockFile = Get-LatestStockFile
}

$data = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json
$stockRows = Load-StockRows $StockFile

$byBarcode = @{}
$byName = @{}
$stockNamePairs = New-Object System.Collections.Generic.List[object]

foreach ($r in $stockRows) {
  if ($r.barcode) { $byBarcode[(Normalize $r.barcode)] = $r }
  if ($r.outbound_name) { $byName[(Normalize $r.outbound_name)] = $r }
  if ($r.outbound_name) {
    $stockName = Normalize-Name $r.outbound_name
    if ($stockName.Length -ge 6) {
      $stockNamePairs.Add([pscustomobject]@{ stock_name = $stockName; row = $r })
    }
  }
}

$matched = 0
$unmatchedData = New-Object System.Collections.Generic.List[object]
$usedStockKeys = New-Object 'System.Collections.Generic.HashSet[string]'

foreach ($item in $data) {
  $candidates = @(
    (Normalize $item.match_sku),
    (Normalize $item.standard_name),
    (Normalize $item.name)
  ) | Where-Object { $_ }

  $stockRow = $null
  $matchKey = $null

  foreach ($key in $candidates) {
    if ($byBarcode.ContainsKey($key)) { $stockRow = $byBarcode[$key]; $matchKey = $key; break }
  }
  if (-not $stockRow) {
    foreach ($key in $candidates) {
      if ($byName.ContainsKey($key)) { $stockRow = $byName[$key]; $matchKey = $key; break }
    }
  }
  if (-not $stockRow) {
    $nameCandidates = @(
      (Normalize-Name $item.standard_name),
      (Normalize-Name $item.name)
    ) | Where-Object { $_.Length -ge 6 }

    $matchedByName = @()
    foreach ($cand in $nameCandidates) {
      $matchedByName = @($stockNamePairs | Where-Object { $cand.Contains($_.stock_name) -or $_.stock_name.Contains($cand) })
      if ($matchedByName.Count -gt 0) { break }
    }
    if ($matchedByName.Count -gt 0) {
      $stockRow = [pscustomobject]@{
        barcode       = ''
        outbound_name = $matchedByName[0].row.outbound_name
        stock_qty     = $matchedByName[0].row.stock_qty
      }
      $matchKey = Normalize-Name $stockRow.outbound_name
    }
  }

  if ($stockRow) {
    $item.stock_qty = [int]$stockRow.stock_qty
    $item.stock_barcode = [string]$stockRow.barcode
    $item.stock_name = [string]$stockRow.outbound_name
    $matched++
    if ($matchKey) { [void]$usedStockKeys.Add($matchKey) }
  } else {
    $item.stock_qty = 0
    $item.stock_barcode = ''
    $item.stock_name = ''
    $unmatchedData.Add([pscustomobject]@{
      retailer       = $item.retailer
      match_sku      = $item.match_sku
      standard_name  = $item.standard_name
      name           = $item.name
      color          = $item.color
      size           = $item.size
    })
  }
}

$unmatchedStock = foreach ($r in $stockRows) {
  $b = Normalize $r.barcode
  $n = Normalize $r.outbound_name
  if (-not $usedStockKeys.Contains($b) -and -not $usedStockKeys.Contains($n)) { $r }
}

$data | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $dataPath -Encoding UTF8

$unmatchedObj = [pscustomobject]@{
  stock_file            = $StockFile
  data_rows             = $data.Count
  stock_rows            = $stockRows.Count
  matched_data_rows     = $matched
  unmatched_data_rows   = $unmatchedData
  unmatched_stock_rows  = $unmatchedStock
}
($unmatchedObj | ConvertTo-Json -Depth 100) | Set-Content -LiteralPath $unmatchedPath -Encoding UTF8

$summary = [pscustomobject]@{
  stock_file            = $StockFile
  data_rows             = $data.Count
  stock_rows            = $stockRows.Count
  matched_data_rows     = $matched
  unmatched_data_rows   = $unmatchedData.Count
  unmatched_stock_rows  = $unmatchedStock.Count
}

$summary | ConvertTo-Json -Depth 10
