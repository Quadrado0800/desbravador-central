# Central de Impressões v0.2

## Instalação

    py -m venv .venv
    .venv\Scripts\activate
    pip install -r requirements.txt
    .venv\Scripts\python.exe app.py

Abra http://127.0.0.1:5000

## Arquivos opcionais

No `app.py`:

    FICHA_HOSPEDES_PATH = r"C:\caminho\Ficha-Hospedes.docx"
    INFORMATIVOS_PATH = r"C:\caminho\Informativos.docx"

Ative com:

    PRINT_FICHA_HOSPEDES = True
    PRINT_INFORMATIVOS = True

A confirmação fica em:

    PRINT_CONFIRMATION = True

Os três podem ser combinados.

## Regra herdada

Mais de 2 quartos = 1 ciclo.
Até 2 quartos = 1 ciclo por quarto.

Cada ciclo pode imprimir confirmação, Ficha-Hospedes e Informativos, conforme as opções acima.

Cookies ficam somente na RAM e o servidor escuta apenas em 127.0.0.1.
O Word é necessário para imprimir os DOCX opcionais.
