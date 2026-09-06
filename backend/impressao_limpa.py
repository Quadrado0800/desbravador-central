import fitz
from pathlib import Path


PDF_ORIGINAL = Path("confirmacao_de_reserva_exemplo_com_2_aptos.pdf")
PDF_FILTRADO = Path("confirmacao_filtrada.pdf")


def filtrar_confirmacao(pdf_original, pdf_saida):

    documento = fitz.open(pdf_original)

    # Percorre as páginas procurando "Informações importantes"
    for numero_pagina in range(len(documento)):

        pagina = documento[numero_pagina]

        ocorrencias = pagina.search_for("Informações importantes")

        if ocorrencias:

            # Encontramos o início das informações que queremos remover
            inicio = ocorrencias[0]

            print(
                f'"Informações importantes" encontrada na página '
                f'{numero_pagina + 1}'
            )

            # O PDF possui rotação.
            # Por isso usamos a MediaBox original para definir
            # a área que será removida.
            area_remover = fitz.Rect(
                inicio.x0,
                0,
                pagina.mediabox.width,
                pagina.mediabox.height
            )

            # Cria uma área de ocultação.
            pagina.add_redact_annot(
                area_remover,
                fill=(1, 1, 1)
            )

            # Aplica a ocultação
            pagina.apply_redactions()

            # Remove todas as páginas posteriores
            if numero_pagina + 1 < len(documento):
                documento.delete_pages(
                    from_page=numero_pagina + 1,
                    to_page=len(documento) - 1
                )

            break

    documento.save(pdf_saida)

    documento.close()

    print()
    print(f"PDF filtrado salvo em:")
    print(pdf_saida.resolve())


if __name__ == "__main__":

    filtrar_confirmacao(
        PDF_ORIGINAL,
        PDF_FILTRADO
    )