"""
extrator_fiscal.py

Funções de extração dos campos fiscais/administrativos:
- Fonte de Recurso        (coluna I)
- Credor(A)                (coluna G)
- Data do Empenho          (coluna J)
- Data da Nota Fiscal      (coluna K)
"""

import re

CORRECOES_FONTE_RECURSO = {
    "Préprios": "Próprios",
    "ATENGAO BASICA": "ATENÇÃO BÁSICA",
    "SANITARIA": "SANITÁRIA",
    "Prdéprios": "Próprios",
}

CORRECOES_CREDOR = {
    "Enderego": "Endereço",
    "Endereco": "Endereço",
    "PANIFICAGAO": "PANIFICAÇÃO",
    "EIREL!": "EIRELI",
}

# Palavras-chave usadas para localizar a data da nota fiscal no texto OCR.
# A nota fiscal costuma ter formatos variados, então buscamos várias
# âncoras textuais possíveis (inclusive variações causadas por erro de OCR).
PALAVRAS_CHAVE_NOTA_FISCAL = [
    "NOTA FISCAL",
    "Nome/Razao",
    "DANFE",
    "Eletrénica",
    "Eletrénico",
    "Fletrénica",
]


def capturar_fonte_recurso(caminho_arquivo):
    """Captura a linha 'Fonte de Recurso:' e corrige erros comuns de OCR."""
    with open(caminho_arquivo, "r") as arquivo:
        for linha in arquivo:
            if "Fonte de Recurso:" in linha:
                corrigida = linha
                for errado, certo in CORRECOES_FONTE_RECURSO.items():
                    corrigida = corrigida.replace(errado, certo)
                return corrigida
    return "- FONTE_Rec_0"


def capturar_credor(caminho_arquivo):
    """Captura a linha 'Credor(A)' e corrige erros comuns de OCR."""
    with open(caminho_arquivo, "r") as arquivo:
        for linha in arquivo:
            if "Credor(A)" in linha:
                corrigida = linha
                for errado, certo in CORRECOES_CREDOR.items():
                    corrigida = corrigida.replace(errado, certo)
                return corrigida
    return "S/_CREDOR(A)"


def capturar_data_empenho(caminho_arquivo):
    """Procura o padrão 'Em: dd/mm/aaaa', típico da data de emissão do empenho."""
    with open(caminho_arquivo, "r") as arquivo:
        for linha in arquivo:
            match = re.search(r"\bEm: \d{2}/\d{2}/\d{4}\b", linha)
            if match:
                return match.group()
    return ": 00/00/0000 Não Encontrada"


def capturar_data_nota_fiscal(output_file):
    """
    Procura, na ordem definida em PALAVRAS_CHAVE_NOTA_FISCAL, a primeira
    data (dd/mm/aaaa) que aparece logo após a palavra-chave encontrada.
    """
    with open(output_file, "r") as arquivo:
        texto = arquivo.read()
        for palavra_chave in PALAVRAS_CHAVE_NOTA_FISCAL:
            posicao = texto.find(palavra_chave)
            if posicao != -1:
                trecho_apos_palavra = texto[posicao + len(palavra_chave):]
                match = re.search(r"\d{2}/\d{2}/\d{4}", trecho_apos_palavra)
                if match:
                    return match.group()
    return "DATA_Da_NOTA_FISCAL_NÃO_ENCONTRADA"
