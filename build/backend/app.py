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
FICHA_HOSPEDES_PATH=""
INFORMATIVOS_PATH=""

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

def build_session():
    session=requests.Session()
    for cookie in _desbravador_cookies:
        kwargs={}
        if cookie.get("domain"): kwargs["domain"]=cookie["domain"]
        if cookie.get("path"): kwargs["path"]=cookie["path"]
        session.cookies.set(cookie["name"],cookie["value"],**kwargs)
    return session

def get_reservas(session):
    url=f"{BASE_URL}/reserva/search"
    params={
        "length":15,"page":1,"valueSearch":"",
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


def imprimir_docx(path, printer_name=None):
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
            print(f"[print] DOCX fallback falhou: {exc}")

    raise RuntimeError("Impressão de DOCX não disponível neste ambiente.")

    raise RuntimeError(
        f"Impressão de DOCX não disponível neste ambiente. Detalhes: {support['reason'] or 'sem suporte Windows'}"
    )

def opcional(path,nome):
    if not path: raise RuntimeError(f"{nome} está ativado, mas sem caminho configurado.")
    p=Path(path)
    if not p.exists(): raise FileNotFoundError(f"{nome} não encontrado: {p}")
    return p

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
            imprimir_docx(opcional(FICHA_HOSPEDES_PATH,"Ficha-Hospedes"), printer_name=printer_name)
            impressos.append("Ficha-Hospedes")
        if settings["print_informativos"]:
            imprimir_docx(opcional(INFORMATIVOS_PATH,"Informativos"), printer_name=printer_name)
            impressos.append("Informativos")
        if settings["print_confirmation"]:
            pdf=None
            try:
                pdf=baixar_pdf(session,reserva["id"])
                imprimir_arquivo(pdf, printer_name=printer_name)
                impressos.append(f"Confirmação #{reserva['id']}")
            finally:
                if pdf and pdf.exists():
                    try:
                        time.sleep(hold_seconds)
                        pdf.unlink()
                    except OSError: pass
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

@app.get("/health")
def health(): return jsonify(ok=True)

if __name__=="__main__":
    app.run(host="127.0.0.1",port=5000,debug=False)
