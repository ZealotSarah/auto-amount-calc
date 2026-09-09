$ErrorActionPreference = 'Stop'

$python = 'C:\Users\71457\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
& $python -m PyInstaller --noconfirm --clean --onefile --windowed --name '基金金额自动测算工具' fund_calculator.py

Write-Host "已生成：$PWD\dist\基金金额自动测算工具.exe"
