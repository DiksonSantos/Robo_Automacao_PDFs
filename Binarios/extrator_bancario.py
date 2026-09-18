"""
extrator_bancario.py

Identifica qual banco emitiu o comprovante (Banco do Brasil ou Caixa
Econômica Federal) e extrai os dados bancários correspondentes:
- Data do pagamento   (coluna L)
- Agência             (coluna M)
- Conta               (coluna N)
- Valor do comprovante (coluna O)

Reaproveita os módulos já existentes e específicos de cada banco:
- Data_Comprov_BB.py       -> data do pagamento (padrão Banco do Brasil)
- Valor_Comprovante_CX.py  -> valor do comprovante (padrão Caixa)
"""

import re

from Data_Comprov_BB import encontrar_data_apos_palavra_chave, chave
from Valor_Comprovante_CX import extrair_valor

PADRAO_AG_CONTA_BB = re.compile(
    r"CLIENTE: [^\n]*\n(?:[^\n]*\n)?AGENCIA: ([^\s]+) CONTA: (\d{1,}\.\d{3}-\d{1,})"
)


def _capturar_valor_comprovante_bb(output_file):
    """Valor do comprovante no padrão Banco do Brasil (última ocorrência da linha)."""
    palavras_chave = ["VALOR: R$", "VALOR COBRADO", "VALOR TOTAL"]
    valores_encontrados = []

    with open(output_file, "r", encoding="utf-8") as arquivo:
        for linha in arquivo:
            for palavra_chave in palavras_chave:
                if linha.startswith(palavra_chave):
                    valores_encontrados.append(linha.split(palavra_chave)[-1].strip())

    if not valores_encontrados:
        return "Sem_Valor_Comprov"

    return valores_encontrados[-1].replace("€", "0")


def _extrair_dados_banco_do_brasil(output_file):
    data_pagamento = encontrar_data_apos_palavra_chave(output_file, chave)
    valor_comprovante = _capturar_valor_comprovante_bb(output_file)

    agencia = "Agencia_Não_Captada"
    conta = "Conta_Não_Encontrada"

    with open(output_file, "r") as arquivo:
        for linha in arquivo:
            if "CLIENTE:" in linha:
                match = PADRAO_AG_CONTA_BB.search(
                    linha + arquivo.readline() + arquivo.readline()
                )
                if match:
                    agencia = match.group(1).replace("@", "0")
                    conta = match.group(2)[0:8]
                break

    return {
        "banco": "Banco do Brasil",
        "data_pagamento": data_pagamento,
        "agencia": agencia,
        "conta": conta,
        "valor_comprovante": valor_comprovante,
    }


def _extrair_dados_caixa(output_file):
    def corrigir_formato(texto):
        return texto.replace("@", "0").replace("S", "$")

    agencia = conta = valor_comprovante = data_debito = None
    linha_encontrada = False

    with open(output_file, "r") as arquivo:
        for linha in arquivo:
            linha = corrigir_formato(linha)

            if "Conta Origem:" in linha:
                partes = re.findall(r"\d+", linha)
                if len(partes) >= 3:
                    agencia = partes[0]
                    conta = partes[2] + "-" + partes[-1]
                elif partes:
                    agencia = partes[0]

            elif "PREFEITURA M CHA PRETA" in linha:
                linha_encontrada = True

            elif linha_encontrada:
                partes = re.findall(r"\d+", linha.strip())
                if len(partes) >= 3:
                    agencia = partes[0]
                    conta = partes[2].strip()

            match_data_hora = re.search(r"(\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2})", linha)
            if match_data_hora:
                data_debito = match_data_hora.group(1).split(" - ")[0]
            else:
                match_data = re.search(r"(\d{2}/\d{2}/\d{4})", linha)
                if match_data:
                    data_debito = match_data.group(1)

            if valor_comprovante is None:
                valor_comprovante = extrair_valor(output_file)

    return {
        "banco": "Caixa Econômica Federal",
        "data_pagamento": data_debito or "Sem_Data_Do_Debito",
        "agencia": agencia or "Agência_não_encontrada",
        "conta": conta or "Conta_não_encontrada",
        "valor_comprovante": valor_comprovante or "Sem_Comprovante",
    }


def identificar_banco_e_extrair(output_file):
    """
    Lê o texto OCR do comprovante, identifica se é Banco do Brasil ou
    Caixa Econômica Federal, e retorna um dicionário com os dados bancários.

    Retorna None se nenhum dos dois bancos for identificado no texto
    (esse é o ponto de extensão citado no LEIA-ME: novos bancos exigiriam
    uma nova função _extrair_dados_<banco> aqui).
    """
    with open(output_file, "r") as arquivo:
        for linha in arquivo:
            if "- BANCO DO BRASIL -" in linha or "Banco do Brasil" in linha:
                return _extrair_dados_banco_do_brasil(output_file)
            elif "GovConta Caixa" in linha or "Conta Origem:" in linha:
                return _extrair_dados_caixa(output_file)
    return None
