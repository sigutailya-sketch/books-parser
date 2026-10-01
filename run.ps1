# Запуск парсера одной командой из PowerShell:
#   .\run.ps1                        — весь каталог
#   .\run.ps1 -c Mystery --details   — любые параметры передаются парсеру как есть
#
# При первом запуске скрипт сам создаст виртуальное окружение и установит библиотеки.
# Если PowerShell запрещает запуск скриптов, выполните один раз:
#   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "Первый запуск: создаю окружение и ставлю библиотеки..."
    py -3 -m venv .venv
    .\.venv\Scripts\python.exe -m pip install --disable-pip-version-check -q -r requirements.txt
}

.\.venv\Scripts\python.exe parser.py @args
