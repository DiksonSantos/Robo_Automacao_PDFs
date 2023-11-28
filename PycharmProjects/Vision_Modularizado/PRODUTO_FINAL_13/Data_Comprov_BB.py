# CAPTURANDO DATA DO PAGAMENTO (COMPROVANTE BANCÁRIO BB)       COL  'L' :

import re


chave = [
    "DATA DA TRANSFERENCIA ",
    "DATA DO PAGAMENTO",
    "DEBITO EM",
    "Data de Débito",
]

global output_file
output_file = 'texto_extraido.txt'

caminho_arquivo = output_file

def encontrar_data_apos_palavra_chave(caminho_arquivo, chave):
    output_file = 'texto_extraido.txt'
    datas_encontradas = set()  # Usar um set para armazenar datas únicas

    with open(caminho_arquivo, "r") as arquivo:
        linhas = arquivo.readlines()

        for linha in linhas:
            for word in chave:
                if word in linha:
                    padrao_data = re.compile(r"(\d{1,2}/\d{1,2}/\d{4})")
                    # Expressão regular para capturar a data no formato dd/mm/aaaa

                    data_encontrada = re.findall(padrao_data, linha)
                    # Usando findall para encontrar todas as datas na linha

                    for data in data_encontrada:
                        data = data.replace('@','0')
                        datas_encontradas.add(
                            data
                        )  # Adiciona a data ao conjunto

    # Adiciona as datas únicas encontradas à lista final_2
    for data in datas_encontradas:
        # final_2.insert(Indice_Oito, data)
        # print('DATA COMPROVANTE: ',data)
        #DATA_PG.insert(0, data)
        return data

    # Verifica se alguma data foi encontrada e, se não, insere "Data_Não_Encontrada"
    if not datas_encontradas:
        # final_2.insert(Indice_Oito, 'Data_Do_Comprovante_Bancário_Não_Encontrada')
        #DATA_PG.insert(0, "Data_Do_Comprovante_Bancário_Não_Encontrada")
        return "Data_Do_Comprovante_Bancário_Não_Encontrada"
    elif datas_encontradas is None:
        return "Data_Do_Comprovante_Bancário_Não_Encontrada"



