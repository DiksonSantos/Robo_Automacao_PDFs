"""
renomear_utils.py

Renomeia o PDF já processado usando o número da Nota de Empenho extraído,
movendo-o para a pasta de arquivos processados. Isso permite rastrear
depois qual PDF gerou qual linha da planilha.
"""

import os
import random
import string


def renomear_pdf(pdf_path, numero_empenho_bruto, pasta_destino):
    """
    numero_empenho_bruto: string bruta capturada pelo extrator_empenho
    (ex.: 'Nota de Empenho: 000123/2023' ou o sentinela 'Nota_Emp: 0000000000000').

    Se o número do empenho não foi identificado, um sufixo aleatório é
    adicionado ao nome do arquivo para evitar sobrescrever outros PDFs
    igualmente "não identificados".
    """
    numero_empenho = "".join(filter(str.isdigit, numero_empenho_bruto))

    if numero_empenho_bruto == "Nota_Emp: 0000000000000":
        sufixo_aleatorio = "".join(
            random.choices(string.ascii_letters + string.digits, k=6)
        )
        novo_nome_arquivo = f"{numero_empenho}_{sufixo_aleatorio}.PDF"
    else:
        novo_nome_arquivo = f"{numero_empenho}.PDF"

    caminho_novo_arquivo = os.path.join(pasta_destino, novo_nome_arquivo)
    os.rename(pdf_path, caminho_novo_arquivo)
    return caminho_novo_arquivo
