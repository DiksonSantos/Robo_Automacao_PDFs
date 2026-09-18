"""
planilha_utils.py

Funções auxiliares para tratar strings antes de escrevê-las na planilha,
e para escrever os valores extraídos nas colunas corretas do Excel.
"""


def tratar_string(string):
    """
    Separa o código da unidade orçamentária do seu nome.
    Ex.: '02.01 - SECRETARIA DE FINANÇAS' -> 'SECRETARIA DE FINANÇAS'
    """
    partes = string.split("-")
    if len(partes) == 2:
        return partes[1].strip()
    elif len(partes) > 2:
        return partes[1].strip() + " " + partes[2].strip()
    else:
        return partes[1] if string.endswith("-") else string.split("-")[0].strip()


def escrever_na_planilha(sheet, colunas, valores):
    """
    Escreve cada valor na primeira linha vazia da coluna correspondente
    (a partir da linha 2, já que a linha 1 é o cabeçalho).
    """
    for coluna, valor in zip(colunas, valores):
        if valor is None:
            continue

        linha_vazia = 2
        celula_destino = f"{coluna}{linha_vazia}"

        while sheet[celula_destino].value is not None:
            linha_vazia += 1
            celula_destino = f"{coluna}{linha_vazia}"

        sheet[celula_destino] = valor
