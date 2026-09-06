import os
import sys
import subprocess
import tempfile
import time
from pathlib import Path

sys.path.insert(0, r'c:\Users\PC\Downloads\desbravador_central_v0_2\backend')
import app

print('IMPRESSORAS:', app.get_printers())
path = Path(tempfile.gettempdir()) / 'teste_impressao.txt'
path.write_text('Teste de impressão do backend\nData: 2026-08-10\n', encoding='utf-8')
print('ARQUIVO:', path)
printer_name = None
printers = app.get_printers()
if printers:
    printer_name = printers[0]
    print('USANDO IMPRESSORA:', printer_name)

try:
    app.imprimir_arquivo(path, printer_name=printer_name)
    print('CHAMADA DE IMPRESSAO ENVIADA')
except Exception as exc:
    print('ERRO NA IMPRESSAO:', exc)

for attempt in range(3):
    print('AGUARDANDO FILA...', attempt + 1)
    time.sleep(2)
    try:
        out = subprocess.check_output([
            'powershell', '-NoProfile', '-Command',
            'Get-PrintJob -ComputerName . | Select-Object -First 5 | Format-Table -AutoSize'
        ], text=True)
        print(out)
    except Exception as exc:
        print('ERRO AO CONSULTAR FILA:', exc)
