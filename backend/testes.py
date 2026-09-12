import requests
import re


COOKIES = {
    "serie": "6148",
    "SESSION": "YjBiOGJjYWUtNThmNS00ODJkLTg4ZTAtYjM4MjhlYjg2ZTJj",
    "SERVERID": "dsl-zd|aqSxH|aqSu/",
}

session = requests.Session()

def identificar_pousada(session):
    serie = session.cookies.get("serie")

    if not serie:
        raise ValueError("Cookie 'serie' não encontrado.")

    url = f"https://desbravadorweb.com.br/acesso/{serie}"

    response = session.get(url)
    response.raise_for_status()

    match = re.search(
        r"nomeTenant:\s*`([^`]+)`",
        response.text
    )

    if not match:
        raise ValueError("nomeTenant não encontrado.")

    return {
        "serie": serie,
        "nome": match.group(1)
    }

identificar_pousada(session)