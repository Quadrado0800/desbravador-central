import customtkinter as ctk
from win32com.client import Dispatch
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
SUMATRA_PATH = BASE_DIR / "tools" / "SumatraPDF.exe"
from time import sleep

def imprimir_docx(path):
    path = Path(path).resolve()

    if not path.exists():
        raise FileNotFoundError(
            f"Arquivo não encontrado: {path}"
        )

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



app = ctk.CTk()
app._set_appearance_mode("dark")
app.geometry("800x400")
app.title("Impressão de Informativos!")

CaixaTexto = ctk.CTkTextbox(master=app, width=700, font=("arial",16,"bold") )
CaixaTexto.insert("0.0", "Olá! \nDigite logo abaixo os apartamentos aos quais você precisa imprimir os informativos: \n\n(para imprimir multiplos, basta escrever uma lista separando por virgula)")
CaixaTexto.configure(state="disabled")
CaixaTexto.pack(padx=20, pady=20)

CaixaEntrada = ctk.CTkEntry(master=app, placeholder_text="001, 102, 203, 404, ...")
CaixaEntrada.pack(padx=20, pady=20)  

def imprimir_infos():
    #função pra imprimir infos:
    ListaPura = CaixaEntrada.get().replace(" ", "").replace("  ", "").split(",")
    desktop = Path.home() / "Desktop"
    infos_folder = desktop / "Informativos_All_Uhs"
    ListaPaths = [f"{infos_folder}/info {x} - OLD.docx" for x in ListaPura]

    for I in ListaPaths:
        imprimir_docx(I)
        sleep(5)
    

CaixaButton = ctk.CTkButton(master=app,width=150, height=48, corner_radius=8, command=imprimir_infos, text="Imprimir", font=("roboto", 18, "bold"))
CaixaButton.pack(padx=20, pady=20)


app.mainloop()