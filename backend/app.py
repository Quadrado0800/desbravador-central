import ctypes
import importlib.util
import os
import sys
import time
from pathlib import Path
from datetime import datetime, timezone
import requests
from flask import Flask, jsonify, render_template, request
from werkzeug.utils import secure_filename
from pathlib import Path
import subprocess
from win32com.client import Dispatch
import fitz
import re
import pythoncom
BASE_DIR = Path(__file__).resolve().parent
SUMATRA_PATH = BASE_DIR / "tools" / "SumatraPDF.exe"

def get_win32_support():
    info={
        "available": False,
        "python_executable": sys.executable,
        "reason": None,
        "pywin32_installed": False,
        "Dispatch": None,
        "SW_SHOWNORMAL": None,
        "shell": None,
        "GetDefaultPrinter": None,
    }
    try:
        info["pywin32_installed"] = bool(
            importlib.util.find_spec("win32api") and importlib.util.find_spec("win32com")
        )
    except Exception as exc:
        info["reason"] = f"find_spec error: {exc}"

    if not info["pywin32_installed"]:
        info["reason"] = info["reason"] or "pywin32 não foi encontrado no interpretador atual."
        return info

    try:
        from win32com.shell import shell
        from win32con import SW_SHOWNORMAL
        from win32com.client import Dispatch
        info.update({
            "available": True,
            "Dispatch": Dispatch,
            "SW_SHOWNORMAL": SW_SHOWNORMAL,
            "shell": shell,
        })
    except Exception as exc:
        info["reason"] = f"{type(exc).__name__}: {exc}"

    return info

app=Flask(__name__)
BASE_DIR=Path(__file__).resolve().parent
DOWNLOADS_DIR=Path.home()/"Downloads"
TEMP_DIR=DOWNLOADS_DIR/"temp"
TEMP_DIR.mkdir(parents=True, exist_ok=True)


def get_temp_dir():
    return TEMP_DIR

# Configure these optional files.
DESKTOP = Path(os.path.join(os.environ["USERPROFILE"], "Desktop"))
FICHA_HOSPEDES_PATH = DESKTOP / "Ficha_Hospedagem.docx"
INFORMATIVOS_PATH = DESKTOP / "Informativos_Hospedagem.docx"
INFORMATIVOS_ALL_UHS_PATH = DESKTOP / "Informativos_All_Uhs"
PRINT_CONFIRMATION=True
PRINT_FICHA_HOSPEDES=False
PRINT_INFORMATIVOS=False
PRINT_HOLD_SECONDS=5


def get_print_settings(overrides=None):
    overrides = overrides or {}
    return {
        "print_confirmation": overrides.get("print_confirmation", PRINT_CONFIRMATION),
        "print_ficha": overrides.get("print_ficha", PRINT_FICHA_HOSPEDES),
        "print_informativos": overrides.get("print_informativos", PRINT_INFORMATIVOS),
    }

BASE_URL="https://desbravadorweb.com.br"
_desbravador_cookies=[]

# ============================================================
# RELATÓRIOS DO DESBRAVADOR
# ============================================================

RELATORIO_CHECKIN_URL = (
    f"{BASE_URL}/relatorios/"
    "relatorioCheckinPrevisao/imprimir"
)

RELATORIO_GOVERNANCA_URL = (
    f"{BASE_URL}/relatorios/"
    "relatorioGovernanca/imprimir"
)

def build_session():
    session=requests.Session()
    for cookie in _desbravador_cookies:
        kwargs={}
        if cookie.get("domain"): kwargs["domain"]=cookie["domain"]
        if cookie.get("path"): kwargs["path"]=cookie["path"]
        session.cookies.set(cookie["name"],cookie["value"],**kwargs)
    return session

def baixar_relatorio(session, url, params, nome_arquivo):
    """
    Baixa um relatório do Desbravador e salva temporariamente como PDF.
    """

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/pdf,*/*",
        "Referer": BASE_URL,
    }

    response = session.get(
        url,
        params=params,
        headers=headers,
        timeout=60
    )

    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "").lower()

    if "pdf" not in content_type and not response.content.startswith(b"%PDF"):
        raise RuntimeError(
            "O Desbravador não retornou um PDF."
        )

    path = get_temp_dir() / nome_arquivo
    path.write_bytes(response.content)

    return path

def baixar_relatorio_checkin_dia(session):
    """
    Gera o relatório detalhado de previsão de check-in
    para o dia atual.
    """

    data = datetime.now().strftime("%d/%m/%Y")

    params = [
        ("rel", "true"),
        ("tipoRelatorio", "DETALHADO"),
        ("pessoaTitular.id", ""),
        ("pessoaTitular.razaoNome", ""),
        ("dataInicio", data),
        ("dataFim", data),
        ("ordenacao", "RESERVA"),
        ("exibirHospedeClassificacao", "on"),
        ("listarObservacaoDosHospedes", "on"),
        ("listarObservacaoPublicaDaHospedagem", "on"),
    ]

    return baixar_relatorio(
        session,
        RELATORIO_CHECKIN_URL,
        params,
        "relatorio_checkin_dia.pdf"
    )

def imprimir_relatorio_checkin_dia():
    """
    Gera e imprime o relatório de check-in do dia.
    """

    if not _desbravador_cookies:
        raise RuntimeError("Sessão não conectada.")


    session = build_session()
    pdf = None

    try:
        pdf = baixar_relatorio_checkin_dia(session)

        imprimir_pdf(pdf)

        return {
            "ok": True,
            "relatorio": "checkin",
            "data": datetime.now().strftime("%d/%m/%Y")
        }

    finally:
        if pdf and pdf.exists():
            try:
                time.sleep(PRINT_HOLD_SECONDS)
                pdf.unlink()
            except OSError:
                pass

def baixar_relatorio_governanca(session, andares, dias_entre_trocas=2):
    """
    Gera um relatório de Governança para os andares informados.
    """

    data = datetime.now().strftime("%d/%m/%Y")

    params = [
        ("rel", "true"),
        ("dataPesquisa", data),

        ("situacoes", "LIVRE"),
        ("situacoes", "OCUPADA"),
        ("situacoes", "MANUTENCAO"),
        ("situacoes", "LIMPEZA"),

        ("checkinPrevisto", "on"),

        ("observacaoHospedagemPublica", "on"),
        ("exibeHospedes", "on"),

        ("diasEntreTrocas", str(dias_entre_trocas)),
        ("tipoAgrupamento", "NENHUM"),
    ]

    # Adiciona os andares
    for andar in andares:
        params.append(("andares", str(andar)))

    nome = (
        "relatorio_governanca_"
        + "_".join(str(a) for a in andares)
        + ".pdf"
    )

    return baixar_relatorio(
        session,
        RELATORIO_GOVERNANCA_URL,
        params,
        nome
    )


def detectar_andares_governanca():
    """
    Detecta a estrutura dos andares a partir dos quartos
    carregados pelo Desbravador.

    Se houver quarto começando com 0:
        0, 1, 2, 3

    Caso contrário:
        1, 2, 3, 4
    """

    reservas = listar_reservas_atuais()

    quartos = []

    for reserva in reservas:
        quartos.extend(reserva.get("quartos", []))

    if not quartos:
        raise RuntimeError(
            "Não foi possível detectar os andares: "
            "nenhum quarto foi encontrado."
        )

    for quarto in quartos:

        # Remove espaços e caracteres não numéricos
        numero = re.sub(r"\D", "", str(quarto))

        if numero and numero[0] == "0":
            return [0, 1, 2, 3]

    return [1, 2, 3, 4]

def imprimir_relatorio_governanca():
    """
    Imprime:

    1. Relatório de todos os andares
    2. Relatório individual de cada andar
    """

    if not _desbravador_cookies:
        raise RuntimeError("Sessão não conectada.")

    body = request.get_json(silent=True) or {}

    dias_entre_trocas = int(
        body.get("dias_entre_trocas", 2)
    )

    if dias_entre_trocas < 0:
        dias_entre_trocas = 0

    session = build_session()

    arquivos = []
    impressos = []

    try:

        # Detecta a estrutura da pousada
        andares = detectar_andares_governanca()

        print(f"[governança] Andares detectados: {andares}")

        # ----------------------------------------------------
        # 1. TODOS OS ANDARES
        # ----------------------------------------------------

        pdf_todos = baixar_relatorio_governanca(
            session,
            andares,
            dias_entre_trocas
        )

        arquivos.append(pdf_todos)

        imprimir_pdf(pdf_todos)

        impressos.append(
            "Governança - Todos os andares"
        )

        # ----------------------------------------------------
        # 2. CADA ANDAR INDIVIDUALMENTE
        # ----------------------------------------------------

        for andar in andares:

            print(
                f"[governança] Imprimindo andar {andar}..."
            )

            pdf = baixar_relatorio_governanca(
                session,
                [andar],
                dias_entre_trocas
            )

            arquivos.append(pdf)

            imprimir_pdf(pdf)

            impressos.append(
                f"Governança - Andar {andar}"
            )

        return {
            "ok": True,
            "relatorio": "governanca",
            "andares": andares,
            "impressos": impressos
        }

    finally:

        for arquivo in arquivos:

            if arquivo.exists():

                try:
                    time.sleep(PRINT_HOLD_SECONDS)
                    arquivo.unlink()

                except OSError:
                    pass



def get_reservas(session):
    url=f"{BASE_URL}/reserva/search"
    params={
        "length":70,"page":1,"valueSearch":"",
        "data.tipoPeriodo.value":"SELECIONE",
        "data.dataInicio.value":"","data.dataFim.value":"",
        "uh.tipoNomeUh.value":"SELECIONE","uh.nomeUh.value":"",
        "numeroReserva.value":"","titular.value":"",
        "situacao.value":"SELECIONE","localizador.value":"",
        "checkinsDoDia.value":"true","reservaOrigem.value":"",
        "channel.value":"true","simboloAdiantamento.value":"R$",
        "advancedFilterActive.value":"false","membership.value":""
    }
    headers={
        "User-Agent":"Mozilla/5.0",
        "Accept":"application/json, text/plain, */*",
        "X-Requested-With":"XMLHttpRequest",
        "Referer":f"{BASE_URL}/#/reserva/"
    }
    r=session.get(url,params=params,headers=headers,timeout=30)
    r.raise_for_status()
    return r.json()

def extrair_dados(data):
    reservas=[]
    for item in data.get("data",[]):
        hospedagens=item.get("details",{}).get("hospedagens",[])
        quartos=[h["abreviaturaUh"] for h in hospedagens if h.get("abreviaturaUh")]
        reservas.append({"id":item["id"],"quartos":quartos})
    return reservas

def listar_reservas_atuais():
    return extrair_dados(get_reservas(build_session()))

def baixar_pdf(session,reserva_id):
    url=f"{BASE_URL}/reserva/manter/gerarRelatorioConfirmacaoReserva/{reserva_id}"
    r=session.get(url,params={"rel":"true","mostraObservacao":"false"},timeout=60)
    r.raise_for_status()
    if "pdf" not in r.headers.get("Content-Type","").lower() and not r.content.startswith(b"%PDF"):
        raise RuntimeError("Resposta não parece ser um PDF.")
    path=get_temp_dir()/f"temp_{secure_filename(str(reserva_id))}.pdf"
    path.write_bytes(r.content)
    return path

def filtrar_pdf_confirmacao(pdf_path):
    pdf_path = Path(pdf_path)

    documento = fitz.open(pdf_path)

    for numero_pagina in range(len(documento)):

        pagina = documento[numero_pagina]

        ocorrencias = pagina.search_for("Informações importantes")

        if ocorrencias:

            inicio = ocorrencias[0]

            print(
                f'"Informações importantes" encontrada na página '
                f'{numero_pagina + 1}'
            )

            # Mantém a página no tamanho e orientação originais.
            # O corte é feito no sistema visual da página.
            area_remover = fitz.Rect(
                pagina.rect.x0,
                inicio.y0,
                pagina.rect.x1,
                pagina.rect.y1
            )

            pagina.add_redact_annot(
                area_remover,
                fill=(1, 1, 1)
            )

            pagina.apply_redactions()

            # Remove todas as páginas posteriores
            if numero_pagina + 1 < len(documento):
                documento.delete_pages(
                    from_page=numero_pagina + 1,
                    to_page=len(documento) - 1
                )

            break

    # Salva em um novo arquivo
    arquivo_filtrado = (
        pdf_path.parent /
        f"{pdf_path.stem}_filtrado.pdf"
    )

    documento.save(arquivo_filtrado)
    documento.close()

    return arquivo_filtrado

def _normalize_printer_name(item):
    if isinstance(item, (tuple, list)):
        for idx in (2, 1, 0):
            if idx < len(item) and isinstance(item[idx], str):
                candidate = item[idx].strip()
                if candidate:
                    if idx == 2 and "," in candidate:
                        candidate = candidate.split(",")[0].strip()
                    if candidate:
                        return candidate
    elif isinstance(item, str):
        return item.strip()
    return None


def get_printers():
    if os.name != "nt":
        return []
    try:
        import win32print
        printers=[]
        for item in win32print.EnumPrinters(2):
            name = _normalize_printer_name(item)
            if name:
                printers.append(name)
        return sorted(set(printers))
    except Exception as exc:
        print(f"[print] get_printers falhou: {exc}")
        return []


def imprimir_via_shell32(path, printer_name=None):
    if os.name != "nt":
        return None

    try:
        shell32 = getattr(ctypes.windll, "shell32", None)
        execute = getattr(shell32, "ShellExecuteW", None)
        if execute is None:
            return None

        if printer_name:
            printer_name = str(printer_name).strip()
            result = execute(None, "printto", str(path.resolve()), printer_name, str(path.parent), 0)
            print(f"[print] ShellExecuteW printer={printer_name} result={result}")
            return result

        return execute(None, "print", str(path.resolve()), None, str(path.parent), 1)
    except Exception as exc:
        print(f"[print] ShellExecuteW falhou: {exc}")
        return None


def set_default_printer(printer_name):
    if os.name != "nt" or not printer_name:
        return False

    try:
        import win32print
        if hasattr(win32print, "SetDefaultPrinter"):
            win32print.SetDefaultPrinter(printer_name)
            return True
    except Exception as exc:
        print(f"[print] Não foi possível definir a impressora padrão: {exc}")

    return False


def imprimir_arquivo(path, printer_name=None):
    path=Path(path)
    if not path.exists(): raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    if os.name == "nt":
        try:
            if printer_name:
                result = imprimir_via_shell32(path, printer_name=printer_name)
                if result is not None and result > 32:
                    return
            os.startfile(str(path.resolve()), "print")
            return
        except Exception as exc:
            print(f"[print] fallback os.startfile falhou: {exc}")

    raise RuntimeError("Impressão não disponível neste ambiente.")

    raise RuntimeError(
        f"Impressão não disponível neste ambiente. Detalhes: {support['reason'] or 'sem suporte Windows'}"
    )
def imprimir_pdf(path):
    if not SUMATRA_PATH.exists():
        raise FileNotFoundError(
            f"SumatraPDF não encontrado: {SUMATRA_PATH}"
        )

    path = Path(path).resolve()

    if not path.exists():
        raise FileNotFoundError(
            f"PDF não encontrado: {path}"
        )

    comando = [
        str(SUMATRA_PATH),
        "-print-to-default",
        str(path)
    ]

    resultado = subprocess.run(
        comando,
        capture_output=True,
        text=True,
        timeout=60
    )

    if resultado.returncode != 0:
        raise RuntimeError(
            f"SumatraPDF retornou código {resultado.returncode}.\n"
            f"{resultado.stderr}"
        )

def imprimir_docx(path, printer_name=None):
    path = Path(path).resolve()

    if not path.exists():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {path}"
        )
    pythoncom.CoInitialize()
    word = Dispatch("Word.Application")
    document = None

    try:
        word.Visible = False
        word.DisplayAlerts = False

        document = word.Documents.Open(
            str(path),
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False,
            NoEncodingDialog=True
        )

        document.PrintOut(
            Background=False
        )

    finally:
        if document is not None:
            try:
                document.Close(
                    SaveChanges=False
                )
            except Exception:
                pass

        try:
            word.Quit(
                SaveChanges=False
            )
        except Exception:
            pass
        pythoncom.CoUninitialize()

def opcional(path,nome):
    if not path: raise RuntimeError(f"{nome} está ativado, mas sem caminho configurado.")
    p=Path(path)
    if not p.exists(): raise FileNotFoundError(f"{nome} não encontrado: {p}")
    return p

def obter_informativo(quarto):
    """
    Procura o informativo pelo número do quarto,
    independentemente do restante do nome do arquivo.
    """

    if not INFORMATIVOS_ALL_UHS_PATH.exists():
        return INFORMATIVOS_PATH

    numero_quarto = re.sub(r"\D", "", str(quarto))

    if not numero_quarto:
        return INFORMATIVOS_PATH

    # Primeiro procura o número EXATO no nome
    for arquivo in INFORMATIVOS_ALL_UHS_PATH.glob("*.docx"):
        numeros_no_nome = re.findall(r"\d+", arquivo.stem)

        if numero_quarto in numeros_no_nome:
            return arquivo

    # Depois aceita diferença apenas de zeros à esquerda
    numero_sem_zeros = str(int(numero_quarto))

    for arquivo in INFORMATIVOS_ALL_UHS_PATH.glob("*.docx"):
        numeros_no_nome = re.findall(r"\d+", arquivo.stem)

        if numero_sem_zeros in numeros_no_nome:
            return arquivo

    return INFORMATIVOS_PATH

def imprimir_checkin(reserva, printer_name=None, print_settings=None):
    if not _desbravador_cookies: raise RuntimeError("Sessão não conectada.")
    settings = get_print_settings(print_settings)
    hold_seconds = int((print_settings or {}).get("print_hold_seconds", PRINT_HOLD_SECONDS))
    quartos=reserva["quartos"]
    session=build_session()
    quantidade=1 if len(quartos)>2 else max(len(quartos),1)
    impressos=[]
    for _ in range(quantidade):
        if settings["print_ficha"]:
            imprimir_docx(opcional(FICHA_HOSPEDES_PATH,"Ficha_Hospedagem"), printer_name=printer_name)
            impressos.append("Ficha_Hospedagem")
        if settings["print_informativos"]:
            quarto = quartos[_] if _ < len(quartos) else quartos[0]
            informativo = obter_informativo(quarto)

            imprimir_docx(
                opcional(informativo, "Informativo"),
                printer_name=printer_name
            )

            impressos.append(f"Informativo - Quarto {quarto}")
        if settings["print_confirmation"]:
            pdf = None
            pdf_filtrado = None

            try:
                # 1. Baixa a confirmação original
                pdf = baixar_pdf(session, reserva["id"])

                # 2. Filtra as páginas indesejadas
                pdf_filtrado = filtrar_pdf_confirmacao(pdf)

                # 3. Imprime o PDF filtrado
                imprimir_pdf(pdf_filtrado)

                impressos.append(f"Confirmação #{reserva['id']}")

            finally:
                # Remove o PDF original
                if pdf and pdf.exists():
                    try:
                        time.sleep(hold_seconds)
                        pdf.unlink()
                    except OSError:
                        pass

                # Remove o PDF filtrado
                if pdf_filtrado and pdf_filtrado.exists():
                    try:
                        pdf_filtrado.unlink()
                    except OSError:
                        pass
    return {"reserva_id":reserva["id"],"quartos":quartos,"impressos":impressos}

@app.get("/")
def index():
    return render_template("index.html",
        connected=bool(_desbravador_cookies),
        print_confirmation=PRINT_CONFIRMATION,
        print_ficha=PRINT_FICHA_HOSPEDES,
        print_informativos=PRINT_INFORMATIVOS)

@app.post("/api/session")
def receive_session():
    global _desbravador_cookies
    body=request.get_json(silent=True) or {}
    cookies=body.get("cookies")
    if not isinstance(cookies,list) or not cookies:
        return jsonify(error="Lista de cookies ausente ou vazia."),400
    clean=[{"name":c["name"],"value":c["value"],
            "domain":c.get("domain"),"path":c.get("path","/")}
           for c in cookies if c.get("name") and "value" in c]
    if not clean: return jsonify(error="Nenhum cookie válido recebido."),400
    _desbravador_cookies=clean
    return jsonify(ok=True,cookie_count=len(clean),
                   received_at=datetime.now(timezone.utc).isoformat())

@app.get("/api/session/status")
def status():
    return jsonify(connected=bool(_desbravador_cookies),
                   cookie_count=len(_desbravador_cookies))

@app.post("/api/reservas")
def reservas():
    try:
        rs=listar_reservas_atuais()
        return jsonify(ok=True,reservas=rs,count=len(rs))
    except Exception as e:
        return jsonify(ok=False,error=str(e)),500

@app.get("/api/printers")
def printers():
    return jsonify(ok=True,printers=get_printers())

@app.post("/api/imprimir/<reserva_id>")
def imprimir(reserva_id):
    try:
        rs=listar_reservas_atuais()
        reserva=next((r for r in rs if str(r["id"])==str(reserva_id)),None)
        if reserva is None: return jsonify(ok=False,error="Reserva não encontrada."),404
        payload=request.get_json(silent=True) or {}
        printer_name=payload.get("printer") or request.args.get("printer")
        return jsonify(ok=True,**imprimir_checkin(reserva, printer_name=printer_name, print_settings=payload))
    except Exception as e:
        return jsonify(ok=False,error=str(e)),500

@app.post("/api/imprimir-lote")
def imprimir_lote():
    try:
        body=request.get_json(silent=True) or {}
        ids=body.get("ids")
        if not isinstance(ids,list) or not ids:
            return jsonify(ok=False,error="Lista de reservas ausente ou vazia."),400

        rs=listar_reservas_atuais()
        mapa={str(r["id"]):r for r in rs}
        impressos=[]
        erros=[]

        for reserva_id in ids:
            chave=str(reserva_id)
            reserva=mapa.get(chave)
            if reserva is None:
                erros.append({"id":chave,"error":"Reserva não encontrada."})
                continue
            try:
                printer_name=body.get("printer")
                impressos.append(imprimir_checkin(reserva, printer_name=printer_name, print_settings=body))
            except Exception as e:
                erros.append({"id":chave,"error":str(e)})

        return jsonify(ok=not erros,printed=impressos,errors=erros,count=len(impressos))
    except Exception as e:
        return jsonify(ok=False,error=str(e)),500

@app.post("/api/relatorio/checkin")
def relatorio_checkin():
    try:
        resultado = imprimir_relatorio_checkin_dia()

        return jsonify(resultado)

    except Exception as e:
        return jsonify(
            ok=False,
            error=str(e)
        ), 500

@app.post("/api/relatorio/governanca")
def relatorio_governanca():
    try:
        resultado = imprimir_relatorio_governanca()

        return jsonify(resultado)

    except Exception as e:
        return jsonify(
            ok=False,
            error=str(e)
        ), 500

@app.get("/health")
def health(): return jsonify(ok=True)

if __name__=="__main__":
    app.run(host="127.0.0.1",port=5000,debug=False)
