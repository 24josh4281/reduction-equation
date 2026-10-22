$ErrorActionPreference = 'Stop'
# Python: Slide Master repo .venv (5 levels up), the standalone package .venv
# (2 levels up), then Python on PATH.
$candidates = @(
  (Join-Path $PSScriptRoot '..\..\..\..\..\.venv\Scripts\python.exe'),
  (Join-Path $PSScriptRoot '..\..\.venv\Scripts\python.exe')
)
$python = $null
foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { $python = (Resolve-Path $c).Path; break } }
if (-not $python) {
  $cmd = Get-Command python -ErrorAction SilentlyContinue
  if (-not $cmd) { $cmd = Get-Command py -ErrorAction SilentlyContinue }
  if (-not $cmd) { throw 'Python을 찾지 못했습니다. setup.ps1을 먼저 실행하세요.' }
  $python = $cmd.Source
}
$env:PYTHONUTF8 = '1'
Write-Host '감축사업 수식·모니터링 인자 편집기를 시작합니다. 브라우저가 자동으로 열립니다.'
Write-Host '종료하려면 이 창에서 Ctrl+C 를 누르세요.'
& $python (Join-Path $PSScriptRoot 'server.py') @args
