"""
ocr_utils.py

Responsável por converter o PDF (imagem digitalizada) em texto, via OCR.

Fluxo: PDF -> imagens (uma por página) -> escala de cinza -> Tesseract OCR
       -> arquivo de texto intermediário (texto_extraido.txt)
"""

from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import cv2
import numpy as np


def extrair_texto_pdf(pdf_path, output_file="texto_extraido.txt"):
    """
    Converte cada página do PDF em imagem, aplica OCR e salva o texto
    extraído em um arquivo .txt intermediário.

    Retorna o caminho do arquivo de texto gerado, que será lido pelos
    módulos de extração (extrator_empenho, extrator_fiscal, extrator_bancario).
    """
    pages_as_images = convert_from_path(pdf_path)

    with open(output_file, "w", encoding="utf-8") as text_file:
        for page_number, page_image in enumerate(pages_as_images):
            # Escala de cinza melhora a taxa de acerto do OCR
            gray_image = cv2.cvtColor(np.array(page_image), cv2.COLOR_RGB2GRAY)
            # lang='por': os comprovantes são em português — sem isso, o Tesseract
            # usa o idioma padrão (inglês), prejudicando o reconhecimento de
            # acentos e palavras específicas (ex.: "Empenho", "Fonte de Recurso").
            texto = pytesseract.image_to_string(Image.fromarray(gray_image), lang="por")
            text_file.write(f"Texto da página {page_number + 1}:\n{texto}\n\n")

    return output_file



"""
# FUNÇÃO SEM O USO DE ENUMERATE:

for item in enumerate(pages_as_images):
    page_number = item[0]
    page_image = item[1]
    ...

"""
