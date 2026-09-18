import re
#
# def extrair_primeiro_valor_apos_palavra(arquivo_path, palavra):
#     with open(arquivo_path, 'r', encoding='utf-8') as arquivo:
#         conteudo = arquivo.read()
#
#         # Procura pela palavra específica
#         #match_palavra = re.search(f'{palavra}\s*([\d\.,]+)', conteudo)
#         match_palavra = re.search(r'R\$ ([\d,.]+)', conteudo)
#
#         if match_palavra:
#             return match_palavra.group(1)
#         else:
#             return 'Valor não encontrado após a palavra específica.'
#
# # Exemplo de uso:
# #arquivo_path = 'texto_extraido.txt'
# arquivo_path = 'Datas_Extraidas_CAIXA.txt'
# palavra_especifica = 'Valor'  # Substitua pela palavra desejada
#
# valor_encontrado = extrair_primeiro_valor_apos_palavra(arquivo_path, palavra_especifica)
#
# print(f"Valor: {valor_encontrado}")
# #
#
# ###############################
#
#

#VALOR_COMPR = []


"""
import re
# VALOR DO COMPROVANTE CAIXA:
def extrair_valor(arquivo_path):
    with open(arquivo_path, 'r', encoding='utf-8') as arquivo:
        conteudo = arquivo.read()

        # Procura pela primeira forma: 'VALOR' em linhas separadas
        match_linhas_separadas = re.search(r'VALOR[\n\s]+R\$ ([\d,.]+)', conteudo)
        if match_linhas_separadas:
            valor = match_linhas_separadas.group(1)
            return valor

        # Procura pela segunda forma: 'VALOR' na mesma linha
        match_mesma_linha = re.search(r'VALOR\s+R\$ ([\d,.]+)', conteudo)
        if match_mesma_linha:
            valor = match_mesma_linha.group(1)
            VALOR_COMPR.insert(0, valor)
            return valor
            
			if not match_mesma_linha:
				VALOR_COMPR.insert(0, 'Sem_Valor_Comprovante')
        # Nenhum valor encontrado
        return VALOR_COMPR[0]#'Valor não encontrado.'

# Exemplo de uso:
#arquivo_path = 'texto_extraido.txt'
arquivo_path = 'Datas_Extraidas_CAIXA.txt'

valor_encontrado = extrair_valor(arquivo_path)

print(f"Valor: {valor_encontrado}")
"""

#global VALOR_COMPR



# VALOR DO COMPROVANTE CAIXA:
def extrair_valor(arquivo_path):
    with open(arquivo_path, 'r', encoding='utf-8') as arquivo:
        conteudo = arquivo.read()
        # Procura pela primeira forma: 'VALOR' em linhas separadas
        match_linhas_separadas = re.search(r'VALOR[\n\s]+R\$ ([\d,.]+)', conteudo)
        if match_linhas_separadas:
            valor = match_linhas_separadas.group(1)
            return valor
        # Procura pela segunda forma: 'VALOR' na mesma linha
        match_mesma_linha = re.search(r'VALOR\s+R\$ ([\d,.]+)', conteudo)
        if match_mesma_linha:
            valor = match_mesma_linha.group(1)

            return valor                 
        # Nenhum valor encontrado
        return 'Valor não encontrado'



if __name__ == "__main__":
	# Exemplo de uso:

	VALOR_COMPR = []
	
	#arquivo_path = 'Datas_Extraidas_CAIXA.txt'
	arquivo_path = 'texto_extraido.txt'
	
	valor_encontrado = extrair_valor(arquivo_path)
	if valor_encontrado:
		VALOR_COMPR.insert(0, valor_encontrado)
	
	
	#print(f"Valor: {valor_encontrado}")
	
	print(VALOR_COMPR[0])
	
