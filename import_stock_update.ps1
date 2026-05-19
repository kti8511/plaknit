param(
  [Parameter(Position=0)]
  [string]$StockFile
)

$ErrorActionPreference = 'Stop'

function Normalize([object]$value) {
  $text = ''
  if ($null -ne $value) { $text = [string]$value }
  $text = [regex]::Replace($text, '\s+', '')
  return $text.ToUpperInvariant()
}

function Normalize-Name([object]$value) {
  $text = Normalize $value
  $text = [regex]::Replace($text, '^\[[^\]]+\]', '')
  return [regex]::Replace($text, '[-_()\[\]/,]', '')
}

$script:SizePattern = '(?i)(?:^|[^A-Z0-9])(4XL|3XL|2XL|XXXL|XXL|XL|XS|FREE|F|S|M|L)(?=$|[^A-Z0-9])'
$script:SizeSuffixPattern = '(?i)(4XL|3XL|2XL|XXXL|XXL|XL|XS|FREE|F|S|M|L)$'
$script:ColorAliases = [ordered]@{
  BLACK = 'BLACK'; BLK = 'BLACK'; '블랙' = 'BLACK'
  WHITE = 'WHITE'; WHT = 'WHITE'; '화이트' = 'WHITE'
  NAVY = 'NAVY'; NVY = 'NAVY'; '네이비' = 'NAVY'
  CHARCOAL = 'CHARCOAL'; '차콜' = 'CHARCOAL'
  GREY = 'GREY'; GRAY = 'GREY'; GRY = 'GREY'; LGR = 'GREY'; MGREY = 'GREY'; LIGHTGREY = 'GREY'; LIGHTGRAY = 'GREY'; '라이트그레이' = 'GREY'; MELANGEGREY = 'GREY'; '멜란지그레이' = 'GREY'
  KHAKI = 'KHAKI'; LKHAKI = 'KHAKI'; '카키' = 'KHAKI'
  BEIGE = 'BEIGE'; '베이지' = 'BEIGE'
  BROWN = 'BROWN'; '브라운' = 'BROWN'
}
$script:ColorSuffixes = @($script:ColorAliases.Keys | Sort-Object Length -Descending)

function Normalize-Size([object]$value) {
  $text = ''
  if ($null -ne $value) { $text = [string]$value }
  $text = $text.ToUpperInvariant().Replace('(', ' ').Replace(')', ' ')
  $matches = [regex]::Matches($text, $script:SizePattern)
  if ($matches.Count -eq 0) { return '' }
  $size = $matches[$matches.Count - 1].Groups[1].Value.ToUpperInvariant()
  if ($size -eq 'F') { return 'FREE' }
  return $size
}

function Get-ProductKey([object]$value) {
  $text = Normalize-Name $value
  if (-not $text) { return '' }
  $text = [regex]::Replace($text, $script:SizeSuffixPattern, '')
  foreach ($color in $script:ColorSuffixes) {
    $text = [regex]::Replace($text, ([regex]::Escape($color) + '$'), '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
  }
  return $text
}

function Get-ColorKey([object[]]$values) {
  $text = ''
  foreach ($value in $values) {
    if ($null -ne $value) { $text += (Normalize-Name $value) }
  }
  $colors = @()
  foreach ($alias in $script:ColorAliases.Keys) {
    if ($text.Contains($alias)) {
      $canonical = $script:ColorAliases[$alias]
      if ($colors -notcontains $canonical) { $colors += $canonical }
    }
  }
  return (@($colors | Sort-Object) -join '|')
}

function Get-LatestStockFile([string]$desktopPath) {
  $patterns = @(
    '재고조회(기본)_*.xlsx',
    '재고조회_*.xlsx',
    '*재고조회*.xlsx'
  )

  $files = @()
  foreach ($pattern in $patterns) {
    $files += Get-ChildItem -LiteralPath $desktopPath -Filter $pattern -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -notlike '~$*' }
  }

  if (-not $files -or $files.Count -eq 0) {
    throw "재고조회(기본) 엑셀 파일을 바탕화면에서 찾지 못했습니다. (Desktop: $desktopPath)"
  }

  return ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}

function Convert-ExcelColumnToIndex([string]$colLetters) {
  $idx = 0
  foreach ($ch in $colLetters.ToUpperInvariant().ToCharArray()) {
    if ($ch -lt 'A' -or $ch -gt 'Z') { break }
    $idx = ($idx * 26) + ([int][char]$ch - [int][char]'A' + 1)
  }
  return $idx - 1
}

function Get-CellText($cell, [string[]]$sharedStrings) {
  if ($null -eq $cell) { return $null }

  $t = $cell.GetAttribute('t')

  if ($t -eq 's') {
    $v = $cell.SelectSingleNode('./*[local-name()="v"]')
    if ($null -eq $v) { return $null }
    $i = [int]$v.InnerText
    if ($i -ge 0 -and $i -lt $sharedStrings.Count) { return $sharedStrings[$i] }
    return $null
  }

  if ($t -eq 'inlineStr') {
    $tNode = $cell.SelectSingleNode('./*[local-name()="is"]/*[local-name()="t"]')
    if ($null -eq $tNode) { return $null }
    return $tNode.InnerText
  }

  $v2 = $cell.SelectSingleNode('./*[local-name()="v"]')
  if ($null -eq $v2) { return $null }
  return $v2.InnerText
}

function Read-StockRowsFromXlsx([string]$xlsxPath) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

  $zip = [System.IO.Compression.ZipFile]::OpenRead($xlsxPath)
  try {
    $shared = @()
    $sharedEntry = $zip.GetEntry('xl/sharedStrings.xml')
    if ($null -ne $sharedEntry) {
      $sr = New-Object System.IO.StreamReader($sharedEntry.Open())
      try { $xmlText = $sr.ReadToEnd() } finally { $sr.Dispose() }
      [xml]$sharedXml = $xmlText
      $siNodes = $sharedXml.SelectNodes('//*[local-name()="sst"]/*[local-name()="si"]')
      foreach ($si in $siNodes) {
        $tNodes = $si.SelectNodes('.//*[local-name()="t"]')
        $sb = New-Object System.Text.StringBuilder
        foreach ($tn in $tNodes) { [void]$sb.Append($tn.InnerText) }
        $shared += $sb.ToString()
      }
    }

    $sheetEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
    if ($null -eq $sheetEntry) { throw "sheet1.xml을 찾지 못했습니다: $xlsxPath" }
    $sr2 = New-Object System.IO.StreamReader($sheetEntry.Open())
    try { $sheetText = $sr2.ReadToEnd() } finally { $sr2.Dispose() }
    [xml]$sheetXml = $sheetText

    $rowNodes = $sheetXml.SelectNodes('//*[local-name()="worksheet"]//*[local-name()="sheetData"]/*[local-name()="row"]')
    if ($rowNodes.Count -lt 1) { throw "시트에 데이터가 없습니다: $xlsxPath" }

    function Get-RowValues($rowNode) {
      $cells = $rowNode.SelectNodes('./*[local-name()="c"]')
      $maxCol = -1
      $cellMap = @{}
      foreach ($c in $cells) {
        $r = $c.GetAttribute('r')
        if (-not $r) { continue }
        $colLetters = ([regex]::Match($r, '^[A-Za-z]+')).Value
        if (-not $colLetters) { continue }
        $col = Convert-ExcelColumnToIndex $colLetters
        if ($col -gt $maxCol) { $maxCol = $col }
        $cellMap[$col] = (Get-CellText $c $shared)
      }
      $vals = New-Object object[] ($maxCol + 1)
      for ($i = 0; $i -le $maxCol; $i++) { $vals[$i] = $cellMap[$i] }
      return $vals
    }

    $headers = @(Get-RowValues $rowNodes[0] | ForEach-Object { if ($null -eq $_) { '' } else { ([string]$_).Trim() } })
    $idx = @{}
    for ($i = 0; $i -lt $headers.Count; $i++) {
      $h = $headers[$i]
      if ($h -and -not $idx.ContainsKey($h)) { $idx[$h] = $i }
    }

    $stockColumn = '총재고'
    if ($idx.ContainsKey('출고가능')) { $stockColumn = '출고가능' }
    elseif ($idx.ContainsKey('현재고')) { $stockColumn = '현재고' }

    $required = @('출고상품명', '바코드', $stockColumn)
    $missing = @($required | Where-Object { -not $idx.ContainsKey($_) })
    if ($missing.Count -gt 0) { throw "필수 컬럼 누락: $($missing -join ', ')" }

    $out = New-Object System.Collections.Generic.List[object]
    for ($ri = 1; $ri -lt $rowNodes.Count; $ri++) {
      $vals = Get-RowValues $rowNodes[$ri]
      if (-not ($vals | Where-Object { $null -ne $_ -and [string]$_ -ne '' })) { continue }

      $barcode = ''
      $outboundName = ''
      $stockRaw = $null

      if ($idx['바코드'] -lt $vals.Count) {
        if ($null -ne $vals[$idx['바코드']]) { $barcode = ([string]$vals[$idx['바코드']]).Trim() }
      }
      if ($idx['출고상품명'] -lt $vals.Count) {
        if ($null -ne $vals[$idx['출고상품명']]) { $outboundName = ([string]$vals[$idx['출고상품명']]).Trim() }
      }
      if ($idx[$stockColumn] -lt $vals.Count) { $stockRaw = $vals[$idx[$stockColumn]] }

      if (-not $barcode -and -not $outboundName) { continue }

      $qty = 0
      if ($null -ne $stockRaw) {
        $s = ([string]$stockRaw).Replace(',', '').Trim()
        [double]$d = 0
        if ([double]::TryParse($s, [ref]$d)) { $qty = [int][math]::Truncate($d) }
      }

      $out.Add([pscustomobject]@{ barcode = $barcode; outbound_name = $outboundName; stock_qty = $qty }) | Out-Null
    }

    return ,$out
  } finally {
    $zip.Dispose()
  }
}

function Write-Json([object]$obj, [string]$path, [bool]$indented) {
  $depth = 100
  if ($indented) {
    $json = $obj | ConvertTo-Json -Depth $depth
  } else {
    $json = $obj | ConvertTo-Json -Depth $depth -Compress
  }
  [System.IO.File]::WriteAllText($path, $json, [System.Text.Encoding]::UTF8)
}

function Update-IndexHtmlRawRows([string]$htmlPath, [string]$rowsJson, [string]$nowText) {
  $text = Get-Content -LiteralPath $htmlPath -Raw -Encoding UTF8

  $pattern = '(?m)^\s*const\s+rawRows\s*=.*;\r?\n'
  $replacement = "const rawRows = $rowsJson;`n"
  if (-not [regex]::IsMatch($text, $pattern)) { throw "rawRows 치환 대상 패턴을 찾지 못했습니다: $htmlPath" }
  $newText = [regex]::Replace($text, $pattern, $replacement)

  $chipPattern = '(<div class="status-chip"><div class="live-dot"></div>)([^<]+)(</div>)'
  $newText2 = [regex]::Replace($newText, $chipPattern, ('$1' + $nowText + '$3'))

  Set-Content -LiteralPath $htmlPath -Value $newText2 -Encoding UTF8 -NoNewline
}

function Set-ObjProp($obj, [string]$name, $value) {
  if ($null -eq $obj.PSObject.Properties[$name]) {
    Add-Member -InputObject $obj -MemberType NoteProperty -Name $name -Value $value -Force
  } else {
    $obj.$name = $value
  }
}

function Get-Count([object]$value) {
  if ($null -eq $value) { return 0 }
  if ($value -is [System.Array]) { return $value.Length }
  if ($value -is [System.Collections.ICollection]) { return $value.Count }
  return 1
}

$root = Split-Path -Parent $PSCommandPath
$dataFile = Join-Path $root 'data.json'
$unmatchedFile = Join-Path $root 'stock_unmatched.json'
$indexFile = Join-Path $root 'index.html'
$publicIndexFile = Join-Path (Join-Path $root 'public') 'index.html'
$desktop = [Environment]::GetFolderPath('Desktop')

if (-not $StockFile) { $StockFile = Get-LatestStockFile $desktop }
if (-not (Test-Path -LiteralPath $StockFile)) { throw "stock_file 경로가 존재하지 않습니다: $StockFile" }

$dataText = Get-Content -LiteralPath $dataFile -Raw -Encoding UTF8
$data = $dataText | ConvertFrom-Json

$stockRows = Read-StockRowsFromXlsx $StockFile

$mergedByName = @{}
foreach ($r in $stockRows) {
  $nameKey = Normalize $r.outbound_name
  if (-not $nameKey) { continue }
  if (-not $mergedByName.ContainsKey($nameKey)) {
    $mergedByName[$nameKey] = [pscustomobject]@{ barcode = $r.barcode; outbound_name = $r.outbound_name; stock_qty = [int]$r.stock_qty }
  } else {
    $mergedByName[$nameKey].stock_qty = [int]$mergedByName[$nameKey].stock_qty + [int]$r.stock_qty
    if (-not $mergedByName[$nameKey].barcode -and $r.barcode) { $mergedByName[$nameKey].barcode = $r.barcode }
  }
}
$stockRows = @($mergedByName.Values)

$byBarcode = @{}
$byName = @{}
$byProductSize = @{}
$productColors = @{}
$stockNamePairs = New-Object System.Collections.Generic.List[object]
foreach ($r in $stockRows) {
  if ($r.barcode) { $byBarcode[(Normalize $r.barcode)] = $r }
  if ($r.outbound_name) { $byName[(Normalize $r.outbound_name)] = $r }
  $productKey = Get-ProductKey $r.outbound_name
  if ($productKey) {
    $stockColor = Get-ColorKey @($r.outbound_name, $r.barcode)
    if (-not $productColors.ContainsKey($productKey)) { $productColors[$productKey] = @() }
    if ($stockColor -and $productColors[$productKey] -notcontains $stockColor) {
      $productColors[$productKey] = @($productColors[$productKey]) + $stockColor
    }
    $stockSize = Normalize-Size $r.outbound_name
    if ($stockSize) {
      $productSizeKey = "$productKey|$stockSize"
      if (-not $byProductSize.ContainsKey($productSizeKey)) {
        $byProductSize[$productSizeKey] = [pscustomobject]@{
          barcode = $r.barcode
          outbound_name = $r.outbound_name
          stock_qty = 0
          used_keys = @()
        }
      }
      $aggregate = $byProductSize[$productSizeKey]
      $aggregate.stock_qty = [int]$aggregate.stock_qty + [int]$r.stock_qty
      if (-not $aggregate.barcode -and $r.barcode) { $aggregate.barcode = $r.barcode }
      foreach ($usedKey in @((Normalize $r.barcode), (Normalize $r.outbound_name))) {
        if ($usedKey -and $aggregate.used_keys -notcontains $usedKey) {
          $aggregate.used_keys = @($aggregate.used_keys) + $usedKey
        }
      }
    }
  }
  if ($r.outbound_name) {
    $sn = Normalize-Name $r.outbound_name
    if ($sn.Length -ge 6) { $stockNamePairs.Add(@($sn, $r)) | Out-Null }
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
  )

  $stockRow = $null
  $matchKey = $null
  $itemColor = ''
  if ($null -ne $item.color) { $itemColor = ([string]$item.color).Trim() }
  $itemProductKey = Get-ProductKey $item.name
  if (-not $itemProductKey) { $itemProductKey = Get-ProductKey $item.standard_name }
  $itemSize = Normalize-Size $item.size
  if (-not $itemSize) { $itemSize = Normalize-Size $item.standard_name }
  if ($itemProductKey -and $itemSize) {
    $colors = @()
    if ($productColors.ContainsKey($itemProductKey)) { $colors = @($productColors[$itemProductKey]) }
    $productSizeKey = "$itemProductKey|$itemSize"
    if ((-not $itemColor -or $colors.Count -le 1) -and $byProductSize.ContainsKey($productSizeKey)) {
      $stockRow = $byProductSize[$productSizeKey]
      $matchKey = $productSizeKey
    }
  }
  foreach ($k in $candidates) {
    if ($null -ne $stockRow) { break }
    if ($k -and $byBarcode.ContainsKey($k)) { $stockRow = $byBarcode[$k]; $matchKey = $k; break }
  }
  if ($null -eq $stockRow) {
    foreach ($k in $candidates) {
      if ($k -and $byName.ContainsKey($k)) { $stockRow = $byName[$k]; $matchKey = $k; break }
    }
  }
  if ($null -eq $stockRow) {
    $nameCandidates = @(
      (Normalize-Name $item.standard_name),
      (Normalize-Name $item.name)
    ) | Where-Object { $_ -and $_.Length -ge 6 }

    $matchedByName = @()
    foreach ($cand in $nameCandidates) {
      $matchedByName = @(
        $stockNamePairs |
          Where-Object { ($cand -and (($_[0] -like "*$cand*") -or ($cand -like "*$($_[0])*"))) } |
          ForEach-Object { $_[1] }
      )
      if ($matchedByName.Count -gt 0) { break }
    }
    if ($matchedByName.Count -gt 0) {
      $stockRow = [pscustomobject]@{ barcode = ''; outbound_name = $matchedByName[0].outbound_name; stock_qty = $matchedByName[0].stock_qty }
      $matchKey = Normalize-Name $stockRow.outbound_name
    }
  }

  if ($null -ne $stockRow) {
    Set-ObjProp $item 'stock_qty' $stockRow.stock_qty
    Set-ObjProp $item 'stock_barcode' $stockRow.barcode
    Set-ObjProp $item 'stock_name' $stockRow.outbound_name
    $matched++
    $usedKeys = @()
    if ($stockRow.PSObject.Properties['used_keys']) { $usedKeys = @($stockRow.used_keys) }
    if ($usedKeys.Count -gt 0) {
      foreach ($usedKey in $usedKeys) {
        if ($usedKey) { [void]$usedStockKeys.Add($usedKey) }
      }
    } elseif ($matchKey) {
      [void]$usedStockKeys.Add($matchKey)
    }
  } else {
    Set-ObjProp $item 'stock_qty' 0
    Set-ObjProp $item 'stock_barcode' ''
    Set-ObjProp $item 'stock_name' ''
    $unmatchedData.Add([pscustomobject]@{ retailer=$item.retailer; match_sku=$item.match_sku; standard_name=$item.standard_name; name=$item.name; color=$item.color; size=$item.size }) | Out-Null
  }
}

$unmatchedStock = @(
  foreach ($r in $stockRows) {
    $bc = Normalize $r.barcode
    $nm = Normalize $r.outbound_name
    if (-not $usedStockKeys.Contains($bc) -and -not $usedStockKeys.Contains($nm)) { $r }
  }
)

$null = Write-Json -obj $data -path $dataFile -indented $true
$summaryObj = [pscustomobject]@{
  stock_file = $StockFile
  data_rows = (Get-Count $data)
  stock_rows = (Get-Count $stockRows)
  matched_data_rows = $matched
  unmatched_data_rows = $unmatchedData.ToArray()
  unmatched_stock_rows = $unmatchedStock
}
$null = Write-Json -obj $summaryObj -path $unmatchedFile -indented $true

$rowsJson = $data | ConvertTo-Json -Depth 100 -Compress

$nowText = (Get-Date).ToString('yyyy-MM-dd HH:mm')
Update-IndexHtmlRawRows $indexFile $rowsJson $nowText
Update-IndexHtmlRawRows $publicIndexFile $rowsJson $nowText

$summary = [pscustomobject]@{
  stock_file = $StockFile
  data_rows = (Get-Count $data)
  stock_rows = (Get-Count $stockRows)
  matched_data_rows = $matched
  unmatched_data_rows = (Get-Count $unmatchedData)
  unmatched_stock_rows = (Get-Count $unmatchedStock)
}
$summary | ConvertTo-Json -Depth 5
