# 감축사업 수식 편집기를 Tailscale 주소(https://<PC>.<tailnet>.ts.net:<포트>/)로 연다.
# 서버는 127.0.0.1에서만 듣고, tailscale serve가 tailnet 안의 HTTPS 요청만 넘겨 준다.
#   serve_tailscale.ps1                 지금 실행 (창을 닫으면 편집기 종료, Tailscale 설정은 남음)
#   serve_tailscale.ps1 -Install        로그인할 때마다 자동 실행되도록 작업 스케줄러에 등록
#   serve_tailscale.ps1 -Uninstall      자동 실행과 Tailscale 공개 해제
param(
  [int]$HttpsPort = 10001,
  [int]$LocalPort = 8095,
  [switch]$Install,
  [switch]$Uninstall
)
$ErrorActionPreference = 'Stop'
$TaskName = 'ReductionEquationsEditor'

$tailscale = (Get-Command tailscale -ErrorAction SilentlyContinue).Source
if (-not $tailscale) {
  foreach ($c in @("$env:ProgramFiles\Tailscale\tailscale.exe", "${env:ProgramFiles(x86)}\Tailscale\tailscale.exe")) {
    if (Test-Path -LiteralPath $c) { $tailscale = $c; break }
  }
}
if (-not $tailscale) { throw 'Tailscale을 찾지 못했습니다. https://tailscale.com/download 에서 설치하고 로그인하세요.' }

if ($Uninstall) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
  & $tailscale serve --https=$HttpsPort off
  Write-Host '자동 실행과 Tailscale 공개를 해제했습니다.'
  return
}

# Python: Slide Master repo .venv (5 levels up), the standalone package .venv (2 levels up), then PATH.
$candidates = @(
  (Join-Path $PSScriptRoot '..\..\..\..\..\.venv\Scripts\python.exe'),
  (Join-Path $PSScriptRoot '..\..\.venv\Scripts\python.exe')
)
$python = $null
foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { $python = (Resolve-Path $c).Path; break } }
if (-not $python) {
  $cmd = Get-Command python -ErrorAction SilentlyContinue
  if (-not $cmd) { throw 'Python을 찾지 못했습니다. setup.ps1을 먼저 실행하세요.' }
  $python = $cmd.Source
}
& $python -c "import pptx" 2>$null
if ($LASTEXITCODE -ne 0) { throw "python-pptx가 없는 Python입니다 ($python). setup.ps1을 먼저 실행하세요." }
$server = Join-Path $PSScriptRoot 'server.py'

if ($Install) {
  $action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`" -HttpsPort $HttpsPort -LocalPort $LocalPort"
  $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
  $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
  Start-ScheduledTask -TaskName $TaskName
  Write-Host "자동 실행을 등록하고 시작했습니다 (작업 이름: $TaskName)."
} else {
  # Publish first, then serve in the foreground so closing the window stops the editor.
  & $tailscale serve --bg --https=$HttpsPort "http://127.0.0.1:$LocalPort" | Out-Null
}

$dns = $null
try { $dns = ((& $tailscale status --json | ConvertFrom-Json).Self.DNSName).TrimEnd('.') } catch {}
if ($dns) { Write-Host "편집기 주소: https://${dns}:$HttpsPort/  (같은 Tailscale 네트워크에서만 열립니다)" }

if (-not $Install) {
  $env:PYTHONUTF8 = '1'
  & $python $server --port $LocalPort --no-browser
}
