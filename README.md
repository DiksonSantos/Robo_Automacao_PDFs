# Automação de Extração de Dados de Comprovantes Bancários (OCR)

## O problema
[Nome/tipo do cliente, ex: "um órgão público" ou "um escritório de contabilidade de uma prefeitura"] processava manualmente dezenas de comprovantes de pagamento em PDF (Banco do Brasil e Caixa Econômica Federal) todo mês, digitando à mão data, valor, credor e número de empenho numa planilha de controle — um processo lento e sujeito a erro de digitação.

## A solução
Desenvolvi uma ferramenta desktop em Python que:

- Recebe os PDFs (inclusive digitalizados/escaneados) através de uma interface simples (Tkinter)
- Converte cada página em imagem e aplica OCR (Tesseract + OpenCV) para extrair o texto
- Usa expressões regulares para identificar automaticamente: data do pagamento, valor, órgão, credor e número de empenho
- Grava tudo já estruturado direto numa planilha Excel (openpyxl), pronta para conferência e arquivamento

## Resultado
- Fim da digitação manual linha a linha desses dados

## Tecnologias
Python · Tesseract OCR · OpenCV · pdf2image · openpyxl · Tkinter



---
*Projeto real, desenvolvido sob demanda para um cliente em 2023.*
