$ErrorActionPreference = 'Stop'

Write-Host "Instalando PyInstaller..." -ForegroundColor Cyan
& c:/Users/PC/Downloads/desbravador_central_v0_2/.venv/Scripts/python.exe -m pip install pyinstaller

Write-Host "Gerando executavel..." -ForegroundColor Cyan
& c:/Users/PC/Downloads/desbravador_central_v0_2/.venv/Scripts/python.exe -m PyInstaller --noconfirm --onefile --name desbravador-central --distpath .\dist --workpath .\build --specpath .\build --add-data "backend/templates;templates" --add-data "extension;extension" --add-data "README.md;." .\backend\app.py

Write-Host "" 
Write-Host "Executavel criado em:" -ForegroundColor Green
Write-Host ".\dist\desbravador-central.exe" -ForegroundColor Yellow
