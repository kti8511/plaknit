Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataPath = Join-Path $repoRoot 'data.json'
$indexPath = Join-Path $repoRoot 'index.html'

$data = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json
$json = $data | ConvertTo-Json -Depth 100 -Compress

$html = Get-Content -LiteralPath $indexPath -Raw -Encoding UTF8
$pattern = '(?s)const\s+rawRows\s*=\s*\[.*?\];'
$replacement = "const rawRows = $json;"

if($html -notmatch $pattern) {
  throw "Could not find 'const rawRows = [...]' block in index.html"
}

$updated = [regex]::Replace($html, $pattern, $replacement, 1)
Set-Content -LiteralPath $indexPath -Value $updated -Encoding UTF8

Write-Output "index.html updated"

