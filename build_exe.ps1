param([string]$PythonPath = 'python')

$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    & $PythonPath -m PyInstaller --noconfirm --clean --onefile --windowed --name '基金金额自动测算工具' fund_calculator.py
    if ($LASTEXITCODE -ne 0) { throw "打包失败，退出码：$LASTEXITCODE" }
    Write-Host "已生成：$PSScriptRoot\dist\基金金额自动测算工具.exe"
} finally {
    Pop-Location
}
