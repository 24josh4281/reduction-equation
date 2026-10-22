$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$py = Get-Command py -ErrorAction SilentlyContinue
if ($py) { & py -3 -m venv .venv } else { & python -m venv .venv }
& .\.venv\Scripts\python.exe -m pip install --upgrade pip --quiet
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt --quiet
$env:PYTHONUTF8 = '1'
& .\.venv\Scripts\python.exe scripts\monitoring_plan.py check examples\plan_example.txt
Write-Host ''
Write-Host '설치 완료. 편집기 실행: powershell -ExecutionPolicy Bypass -File scripts\equation_editor\start_equation_editor.ps1'
