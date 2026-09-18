"""
main.py — Soft_13 (versão modularizada)

Automatiza a leitura de comprovantes de empenho, notas fiscais e
comprovantes bancários (Banco do Brasil / Caixa Econômica Federal) em PDF,
extrai os dados relevantes via OCR + regex, e lança tudo em despesa.xlsx.

Este arquivo é o ORQUESTRADOR: mantém o estado (listas globais) e chama,
na ordem certa, as funções de extração que vivem nos módulos especializados:

    ocr_utils.py        -> converte PDF em texto (OCR)
    extrator_empenho.py -> valor/número do empenho, órgão, elemento de despesa
    extrator_fiscal.py  -> fonte de recurso, credor, data do empenho/NF
    extrator_bancario.py-> identifica o banco e extrai dados bancários
    planilha_utils.py   -> formatação de strings e escrita no Excel
    renomear_utils.py   -> renomeia o PDF processado
"""

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import time
import datetime
import re
import os

import openpyxl

from ocr_utils import extrair_texto_pdf
from extrator_empenho import (
    capturar_valor_empenho,
    capturar_numero_empenho,
    capturar_unidade_orcamentaria,
    capturar_elemento_despesa,
)
from extrator_fiscal import (
    capturar_fonte_recurso,
    capturar_credor,
    capturar_data_empenho,
    capturar_data_nota_fiscal,
)
from extrator_bancario import identificar_banco_e_extrair
from planilha_utils import tratar_string, escrever_na_planilha
from renomear_utils import renomear_pdf

Hora_inicial = datetime.datetime.now().strftime("%H:%M")
tempo_inicial = time.time()

# ------------------------------------------------------------------------
# LISTAS GLOBAIS DE ESTADO
# Guardam os dados extraídos do PDF que está sendo processado no momento.
# São preenchidas por processar_pdf() e limpas (.clear()) ao final de cada
# arquivo, para que dados de um comprovante nunca "vazem" para o próximo.
# ------------------------------------------------------------------------
ORGAO = []
OBJETO = []
CREDOR = []
VALOR_EMP = []
RECURSO = []
DATA_EMP = []
DATA_NF = []
DATA_PG = []
AG = []
CONTA = []
VALOR_COMPR = []
N_EMPRENHO = []

# Pasta para onde vão os PDFs já processados e renomeados (ajustar por máquina)
CAMINHO_PASTA_PDF = "/home/dikson/PycharmProjects/Alagoas/Modularizado_Set_2026/Renamed"

# Ordem das colunas na planilha despesa.xlsx
COLUNAS = ["D", "E", "F", "I", "H", "G", "J", "K", "L", "M", "N", "O"]


def processar_pdf(pdf_path, sheet):
    """Processa um único PDF: OCR -> extração de campos -> escrita -> renomeação."""

    if not os.path.exists(pdf_path) or not pdf_path.endswith(".PDF"):
        print(f"Arquivo PDF não encontrado ou extensão incorreta: {pdf_path}")
        return

    try:
        output_file = extrair_texto_pdf(pdf_path)

        # ---------------- CAMPOS DO EMPENHO ----------------
        try:
            VALOR_EMP.insert(0, capturar_valor_empenho(output_file))
        except Exception as e:
            print(f"Ocorreu um erro em VALOR DO EMPENHO: {e}")

        try:
            N_EMPRENHO.append(capturar_numero_empenho(output_file))
        except Exception as e:
            print(f"Ocorreu um erro ao processar 'Nota de Empenho': {e}")

        try:
            ORGAO.append(capturar_unidade_orcamentaria(output_file))
        except Exception as e:
            print(f"Ocorreu um erro ao processar 'Unidade Orçamentária': {e}")

        try:
            OBJETO.append(capturar_elemento_despesa(output_file))
        except Exception as e:
            print(f"Ocorreu um erro ao captar Elemento De Despesa: {e}")

        # ---------------- CAMPOS FISCAIS ----------------
        try:
            RECURSO.append(capturar_fonte_recurso(output_file))
        except Exception as e:
            print(f"Ocorreu um erro ao captar FONTE DE RECURSOS: {e}")

        try:
            CREDOR.append(capturar_credor(output_file))
        except Exception as e:
            print(f"Ocorreu um erro ao processar CREDOR(A): {e}")

        try:
            DATA_EMP.append(capturar_data_empenho(output_file))
        except Exception as e:
            print(f"Ocorreu um erro na Data do Empenho: {e}")

        try:
            DATA_NF.append(capturar_data_nota_fiscal(output_file))
        except Exception as e:
            print(f"Ocorreu um erro ao procurar Data NF: {e}")

        # ---------------- DADOS BANCÁRIOS ----------------
        try:
            dados_bancarios = identificar_banco_e_extrair(output_file)
            if dados_bancarios:
                if dados_bancarios["data_pagamento"] not in DATA_PG:
                    DATA_PG.insert(0, dados_bancarios["data_pagamento"])
                if dados_bancarios["agencia"] not in AG:
                    AG.insert(0, dados_bancarios["agencia"])
                if dados_bancarios["conta"] not in CONTA:
                    CONTA.append(dados_bancarios["conta"])
                if dados_bancarios["valor_comprovante"] not in VALOR_COMPR:
                    VALOR_COMPR.append(dados_bancarios["valor_comprovante"])
        except Exception as e:
            print(f"Erro em Dados Bancários: {e}")

        # Fallbacks para quando nenhum banco reconhecido foi identificado
        if not DATA_PG:
            DATA_PG.append("Falha_PG")
        if not AG:
            AG.append("Falha_AG")
        if not CONTA:
            CONTA.append("Falha_CT")
        if not VALOR_COMPR:
            VALOR_COMPR.append("Comprovante_Ausente")

        # ---------------- MONTAGEM DAS COLUNAS FINAIS ----------------
        H = VALOR_EMP[0]
        D = "".join(filter(str.isdigit, N_EMPRENHO[0]))
        E = tratar_string(ORGAO[0])
        F = OBJETO[0]

        # Fonte de recurso: mantém apenas o texto após o último código numérico
        linha_recurso = RECURSO[0]
        padrao = re.search(r"[^0-9.]+$", linha_recurso)
        I = padrao.group().strip() if padrao else "Sem_Fonte_De_Recursos"

        G = (
            CREDOR[0]
            .split("Endereço")[0]
            .strip()
            .replace("Credor(A):", "")
            .replace("Credor(A);", "")
            .strip()
        )
        J = DATA_EMP[0].split(":")[1].strip()
        K = DATA_NF[0]
        L = DATA_PG[0]
        M = AG[0]
        N = list(CONTA)[0]
        O = VALOR_COMPR[0]

        final_linha = [D, E, F, I, H, G, J, K, L, M, N, O]
        escrever_na_planilha(sheet, COLUNAS, final_linha)

        print(f"Dados do arquivo {pdf_path} adicionados à planilha.")

        # ---------------- RENOMEAÇÃO DO PDF ----------------
        if N_EMPRENHO:
            novo_caminho = renomear_pdf(pdf_path, N_EMPRENHO[0], CAMINHO_PASTA_PDF)
            print(f"Arquivo {pdf_path} renomeado para {novo_caminho}")

    except Exception as e:
        print(f"Ocorreu um erro ao processar o PDF {pdf_path}: {e}")

    finally:
        # Limpa todas as listas para o próximo PDF do lote
        ORGAO.clear()
        OBJETO.clear()
        CREDOR.clear()
        VALOR_EMP.clear()
        RECURSO.clear()
        DATA_EMP.clear()
        DATA_NF.clear()
        DATA_PG.clear()
        AG.clear()
        CONTA.clear()
        VALOR_COMPR.clear()
        N_EMPRENHO.clear()


def selecionar_arquivos_pdf():
    """Abre o seletor de arquivos e processa todos os PDFs escolhidos em lote."""
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Selecione arquivos PDF", filetypes=[("PDF Files", "*.PDF")]
    )

    if file_paths:
        workbook = openpyxl.load_workbook("despesa.xlsx")
        sheet = workbook.active

        for pdf_path in file_paths:
            processar_pdf(pdf_path, sheet)

        workbook.save("despesa.xlsx")
        workbook.close()


if __name__ == "__main__":
    selecionar_arquivos_pdf()

    tempo_decorrido = time.time() - tempo_inicial
    Hora_final = datetime.datetime.now().strftime("%H:%M")

    messagebox.showinfo(
        "Processo(s) Finalizado(s) !",
        f"Tempo levado: {tempo_decorrido:.2f} segundos\n"
        f"Iniciado em: {Hora_inicial}\n"
        f"Finalizado em: {Hora_final}",
    )
    print("TODAS AS TAREFAS CONCLUÍDAS 🌒")
