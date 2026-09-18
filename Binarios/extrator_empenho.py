"""
extrator_empenho.py

Funções de extração dos campos ligados à Nota de Empenho:
- Valor do Empenho              (coluna H)
- Número da Nota de Empenho     (coluna D)
- Unidade Orçamentária / Órgão  (coluna E)
- Elemento de Despesa           (coluna F)

Cada função lê o texto OCR (texto_extraido.txt) e devolve o valor já
tratado (com correções de acentuação típicas de erro de OCR), ou um
valor-sentinela caso o campo não seja encontrado.
"""

# Lista de possíveis "Elementos de Despesa" reconhecidos nos documentos.
# Mantida aqui porque é um dado de apoio específico desta extração
# (diferente das listas de ESTADO do processamento, que ficam no main.py).
ELEMENTOS_DE_DESPESA = [
    " OUTROS SERVICOS DE TERCEIROS - PESSOA JURIDICA",
    "OUTROS SERVIGOS DE TERCEIROS - PESSOA JURIDICA",
    " OUTROS SERVIGOS DE TERCEIROS - PESSOA FISICA",
    "MATERIAL DE CONSUMO",
    "DIARIAS - PESSOAL CIVIL",
    "OBRIGAGOES PATRONAIS",
    "VENCIMENTO E VANTAGENS FIXAS - PESSOA CIVIL",
    "DESPESAS DE EXERCICIO ANTERIORES",
    "APOSENTADORIA, RESERVA REMUNERADA E REFORMAS",
    "PENSOES",
    "DIARIAS -",
    "PENSÃO ALIMENTICIA",
    "CAMARA VEREADORES",
    "PRINCIPAL DA DIVIDA POR CONTRATO",
    "OBRIGAÇÕES TRIBUTÁRIAS E CONTRIBUTIVAS",
    "PASSAGENS E DESPESAS COM LOCOMOGAO",
    "VENCIMENTOS E VANTAGENS FIXAS - SERVIDORES",
    "CONTRATAGAO POR TEMPO DETERMINADO",
    "MATERIAL, BEM OU SERVIGOS PARA DISTRIBUIÇÃO GRATUITA",
    "SERVIÇOS DE CONSULTORIA",
    "EQUIPAMENTOS E MATERIAL PERMANENTE",
    "ARRENDAMENTO MERCANTIL",
    "EMPRESTIMOS CONSIGNADOS",
    "OUTROS BENEDICIOS PREVIDENCIARIOS",
    "SENTENGAS JUDICIAIS",
    "AQUISIGAO DE IMOVEIS",
    "OUTROS AUXILIOS FINANCEIROS A PESSOA FÍSICA",
    "JUROS SOBRE DIVIDA POR CONTRATO",
    "PREMIAGOES CULTURAIS",
]

# Correções de acentuação comuns cometidas pelo OCR nesses termos
CORRECOES_ELEMENTO_DESPESA = {
    "SERVIGOS": "SERVIÇOS",
    "PREMIAGOES": "PREMIAÇÕES",
    "AQUISIGAO": "AQUISIÇÃO",
    "SENTENGAS": "SENTENÇAS",
    "CONTRATAGAO": "CONTRATAÇÃO",
    "LOCOMOGAO": "LOCOMOÇÃO",
    "OBRIGAGOES": "OBRIGAÇÕES",
}

CORRECOES_UNIDADE_ORCAMENTARIA = {
    "Unidade Orgamentaria": "Unidade Orçamentária",
    "FINANCAS": "FINANÇAS",
    "ADMINISTRAGAO": "ADMINISTRAÇÃO",
    "EDUCAGAO": "EDUCAÇÃO",
    "SAUDE": "SAÚDE",
    "ASSISTENCIA": "ASSISTÊNCIA",
    "PROMOGAO E CULTURA": "PROMOÇÃO E CULTURA",
}


def capturar_valor_empenho(caminho_arquivo):
    """
    Tenta dois padrões possíveis no texto extraído, nesta ordem:
    1) 'Data do Contrato: R$ ... R$ ... R$ ...'  -> usa o 3º valor da linha
    2) 'Valor do Empenho: R$ ...'                -> valor na mesma linha
    """
    with open(caminho_arquivo, "r") as arquivo:
        for linha in arquivo:
            if "Data do Contrato:" in linha:
                linha_corrigida = linha.replace("S", "$")
                if "R$" in linha_corrigida:
                    return linha_corrigida.split("R$")[2].strip()
                break

    with open(caminho_arquivo, "r") as arquivo:
        for linha in arquivo:
            if "Valor do Empenho: R$" in linha:
                return linha.split("R$")[1].strip()

    return "R$ Val_NULO_"


def capturar_numero_empenho(caminho_arquivo):
    """Captura a linha contendo 'Nota de Empenho'."""
    with open(caminho_arquivo, "r") as arquivo:
        for linha in arquivo:
            if "Nota de Empenho" in linha:
                return linha
    return "Nota_Emp: 0000000000000"


def capturar_unidade_orcamentaria(caminho_arquivo):
    """Captura a linha do órgão/unidade orçamentária, corrigindo erros de OCR."""
    with open(caminho_arquivo, "r") as arquivo:
        for linha in arquivo:
            if "Unidade Or" in linha:
                linha_corrigida = linha
                for errado, certo in CORRECOES_UNIDADE_ORCAMENTARIA.items():
                    linha_corrigida = linha_corrigida.replace(errado, certo)
                return linha_corrigida
    return "Un_Orçament - Un_Or_Não_Encontrada"


def capturar_elemento_despesa(caminho_arquivo):
    """Procura, linha a linha, por qualquer frase da lista ELEMENTOS_DE_DESPESA."""
    with open(caminho_arquivo, "r") as arquivo:
        for linha in arquivo:
            for frase in ELEMENTOS_DE_DESPESA:
                if frase in linha:
                    corrigida = frase
                    for errado, certo in CORRECOES_ELEMENTO_DESPESA.items():
                        corrigida = corrigida.replace(errado, certo)
                    return corrigida
    return "Elem_Desp_Não_Encontrado"
