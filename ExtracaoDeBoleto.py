import PyPDF2
import re
import pandas as pd
import os
from openpyxl import load_workbook
from collections import defaultdict

# Caminho do pdf que vai extrair os dados e o caminho do excel para onde os dados irão 
caminho_do_excel = "mascara.xlsx"
caminho_do_pdf = "Fatura2.pdf"


# Funções auxiliares para extrair textos, separar palavras, etc.
def extrair_texto_pdf(caminho_arquivo):
    texto = " "
    with open(caminho_arquivo, 'rb') as arquivo:
        leitor_pdf = PyPDF2.PdfReader(arquivo)
        num_paginas = len(leitor_pdf.pages)
        for pagina in range(num_paginas):
            texto += f" {leitor_pdf.pages[pagina].extract_text()}"
    return texto

def separar_palavras(texto):
    # Usa regex para encontrar uma transição de dígito ou vírgula para letra e insere um espaço
    texto_modificado = re.sub(r'(?<=\d),(?=\D)', r', ', texto)
    texto_modificado = re.sub(r'(?<=\d)(?=[A-Z])', r' ', texto_modificado)
    return texto_modificado

# Extraí o texto do pdf escolhido

texto_extraido = extrair_texto_pdf(caminho_do_pdf)

# Corta o texto em várias partes
texto_emlista = texto_extraido.split("\n")

# Separa o texto do valor a ser pago no boleto e o endereço (o scanner deu erro pra separar esses 2 elementos)
texto_original_doValoraPagar = texto_emlista[13]
texto_modificado_doValoraPagar = f"{separar_palavras(texto_original_doValoraPagar)}"
lista_resultadoDoValoraPagar = list((texto_modificado_doValoraPagar.split())[2:4])
resultadoDoValoraPagar = [' '.join(lista_resultadoDoValoraPagar)]

# Lista que se refere ao mês e o ano da fatura
lista_do_mesAno = [texto_modificado_doValoraPagar.split()[0]]

# Lista que se refere a data de vencimento da fatura
lista_do_Vencimento = [texto_modificado_doValoraPagar.split()[1]]

# Ajustar o dados que estavam juntos após o scanner
texto_original_da_leitura = texto_emlista[17]
texto_modificado_da_leitura = f'{separar_palavras(texto_original_da_leitura)}'
lista_resultado_daleitura = list((texto_modificado_da_leitura.split()))
def ajustar_data_com_codigo(lista_resultado_daleitura):
    # Nova lista para armazenar os resultados
    nova_lista = []
    for item in lista_resultado_daleitura:
        # Verifica se o item corresponde ao padrão com código seguido de data
        if re.match(r'^\d{11}/\d{2}/\d{4}$', item):
            # Insere um espaço entre o código e a data
            item_modificado = item[:9] + " " + item[9:]
            nova_lista.append(item_modificado.split(" ")[1])
        else:
            nova_lista.append(item)
    return nova_lista
lista_modificada_da_leitura = ajustar_data_com_codigo(lista_resultado_daleitura[-4:])

# Dicionário para armanezar o total e as leituras
dict_leitura = {
    'Leitura Anterior': lista_modificada_da_leitura[0],
    'Leitura Atual': lista_modificada_da_leitura[1],
    'N° de dias': lista_modificada_da_leitura[2],
    'Proxima Leitura': lista_modificada_da_leitura[3],
}

dict_vencimento_mes_ano = {
    'Mês/Ano': lista_do_mesAno[0],
    'Vencimento': lista_do_Vencimento[0],
}

dict_total_a_pagar = {
    'Total a Pagar': resultadoDoValoraPagar[0],
}

# Separa os valores do medidor e armazena numa lista 
texto_original_doMedidor = texto_emlista[32]
medidor_lista = list()
for texto in texto_emlista:
    for i in range(len(texto)):
        if texto[i:i+16] == "ENERGIA ATIVA - ":
            i+=16
            medidor_lista.append(texto.split())

for frase_medidor in medidor_lista:
    medidor = list()
    medidor.append(frase_medidor[0])
    medidor.append((" ".join(frase_medidor[-9:-5])))
    medidor.append(frase_medidor[5])  
    medidor.append(frase_medidor[6])  
    medidor.append(frase_medidor[7])  
    medidor.append(frase_medidor[8])  
    medidor.append(frase_medidor[9])  

# Dicionário para armanezar os valores do medidor
dict_medida = {
    'Medidor': medidor[0],
    'Grandezas': medidor[1],
    'Postos Tarifários': medidor[2],
    'Leitura Anterior': medidor[3],
    'Leitura Atual': medidor[4],
    'Const. Medidor': medidor[5],
    'Consumo kWh/kw': medidor[6],
}

# Cria uma lista para a energia e armazena os valores
energia_lista = []

for texto in texto_emlista:
    for i in range(len(texto)):
        if texto[i:i+10] == "Energia At":
            i+=10
            energia_lista.append(texto.split()) #[[],[]]

cip_lista = []
for texto in texto_emlista:
    for i in range(len(texto)):
        if texto[i:i+10] == "CIP ILUM P":
            i+=10
            cip_lista.append(texto.split()) 

# Cria uma lista para a energia e armazena os valores
for frase_cip in cip_lista:
    cip = list()
    cip.append((" ".join(frase_cip[0:-5]),frase_cip[-5:]))     
    dist_cip = dict(cip)
    for i in range(3):
        dist_cip["CIP ILUM PUB PREF MUNICIPAL"].insert(0,"") 
    
    dist_cip["CIP ILUM PUB PREF MUNICIPAL"].append("") 

lista_dist_energia = list()
for frase in energia_lista:
    #[(chave,valor),(chave,valor)]
    tmp = list()
    copia = frase[-9:]
    frase[-9] = copia[-4]
    frase[-6] = copia[-9]
    frase[-4] = copia[-6]
    frase[-5] = copia[-7]
    frase[-7] = copia[-5]
    tmp.append((" ".join(frase[0:-9]),frase[-9:]))
    dist_descricao = dict(tmp)
    lista_dist_energia.append(dist_descricao)
lista_dist_energia.append(dist_cip)

# Coloca o sinal de negativo na frente, pois o scan deixa de acordo com a fatura da enel
nova_lista_dist_energia = []
for item in lista_dist_energia:
    novo_dict = {}
    for chave, valores in item.items():
        nova_lista_valores = []
        for valor in valores:
            if valor.endswith('-'):
                # Move o '-' para o começo e remove do final
                novo_valor = '-' + valor[:-1]
                nova_lista_valores.append(novo_valor)
            else:
                nova_lista_valores.append(valor)
        novo_dict[chave] = nova_lista_valores
    nova_lista_dist_energia.append(novo_dict)

# Coloca os valores como float e não str

def parse_number(value):
    if isinstance(value, str):  # Verifica se o valor é uma string
        value = value.strip()  # Remove espaços em branco nas bordas
        if value == '':
            return ''  # Retorna a string vazia diretamente
        if '%' in value:
            # Remove o símbolo de porcentagem e substitui vírgulas por pontos
            value = value.replace('%', '').replace(',', '.')
            # Converte para float e divide por 100 para obter a forma decimal
            return float(value) / 100
        else:
            # Remove pontos usados como separadores de milhar, e substitui vírgula por ponto decimal
            value = re.sub(r'\.(?=\d{3}(?:,|$))', '', value)  # Remove pontos em números como 1.000,50
            value = value.replace(',', '.')
            return float(value)  # Converte a string para float
    return value  # Retorna o valor não alterado se não for uma string

def transform_dict_values(dictionaries):
    for dictionary in dictionaries:
        for key, values in dictionary.items():
            dictionary[key] = [values[0]] + [parse_number(value) for value in values[1:]]

# Aplicar a transformação
transform_dict_values(nova_lista_dist_energia)

def converter_valor_para_float(valor_str):
    # Remove o símbolo 'R$' e espaços em branco
    valor_str = valor_str.replace('R$', '').strip()
    # Substitui vírgula por ponto para o formato decimal correto
    valor_str = valor_str.replace('.', '')  # Remove pontos usados para milhares
    valor_str = valor_str.replace(',', '.')  # Troca vírgula por ponto para decimal
    # Converte a string para float
    valor_float = float(valor_str)
    return valor_float

# Aplica a conversão para todos os valores no dicionário
for chave, valor in dict_total_a_pagar.items():
    dict_total_a_pagar[chave] = converter_valor_para_float(valor)

"""
""" 



# Dicionário para armazenar os itens agrupados por quantidade de energia
agrupados_por_kWh = defaultdict(list)

# Agrupar os itens com base na quantidade de kWh
for dicionario in nova_lista_dist_energia:
    for chave, valor in dicionario.items():
        quantidade = valor[1]  # A quantidade de kWh está na segunda posição da lista
        agrupados_por_kWh[quantidade].append({chave: valor})

# Converter o resultado em uma lista para fácil visualização ou uso futuro
lista_agrupada = list(agrupados_por_kWh.values())

# Criar um dicionário para armazenar cada parte dos arrays multidimensionais
dict_resultante = {}

# Iterar sobre a lista agrupada e adicionar cada parte a um dicionário separado
for indice, parte in enumerate(lista_agrupada):
    dict_resultante[f'parte_{indice + 1}'] = parte

# Lista nova para separar os valores quando for colocar na planilha
dict_parte_2_separada = {}

# Extrair 'parte_2' do dicionário original e colocar no novo dicionário
if 'parte_2' in dict_resultante:
    dict_parte_2_separada['parte_2'] = dict_resultante.pop('parte_2')


wb = load_workbook(caminho_do_excel)

# Selecionar a planilha na qual você deseja inserir os novos dados
planilha = wb.active

intervalo = planilha['I7':'R15']

# Limpar os dados no intervalo especificado
for linha in intervalo:
    for celula in linha:
        celula.value = None
# Suponha que você queira inserir os novos dados a partir da célula I7
linha_excel = 7

# Iterar sobre os valores do dicionário e inserir os dados na planilha
for dados in dict_resultante.values():
    for item in dados:
        for chave, valor in item.items():
            planilha[f'I{linha_excel}'] = chave  # Adiciona a chave na célula atual
            coluna_excel = 'J'  # Começar a partir da coluna J
            for v in valor:
                planilha[f'{coluna_excel}{linha_excel}'] = v  # Adiciona cada valor na mesma linha
                coluna_excel = chr(ord(coluna_excel) + 1)  # Avança para a próxima coluna
        linha_excel += 1  # Avança para a próxima linha

# Salvar as alterações no arquivo Excel



intervalo = planilha['J21':'P21']

# Limpar os dados no intervalo especificado
for linha in intervalo:
    for celula in linha:
        celula.value = None

# Array multidimensional com os novos dados


# Suponha que você queira inserir os novos dados a partir da célula I6
coluna_excel = 'J'
linha_excel = 21

for chave, valor in dict_medida.items():
    planilha[f'{coluna_excel}{linha_excel}'] = valor  # Adiciona o valor na célula atual
    coluna_excel = chr(ord(coluna_excel) + 1)  # Avança para a próxima coluna

#disct_LEITURA - adicionar no espaço da leitura


intervalo = planilha['J24':'M24']

# Limpar os dados no intervalo especificado
for linha in intervalo:
    for celula in linha:
        celula.value = None

# Array multidimensional com os novos dados


# Suponha que você queira inserir os novos dados a partir da célula I6
coluna_excel = 'J'
linha_excel = 24

for chave, valor in dict_leitura.items():
    planilha[f'{coluna_excel}{linha_excel}'] = valor  # Adiciona o valor na célula atual
    coluna_excel = chr(ord(coluna_excel) + 1)  # Avança para a próxima coluna


intervalo = planilha['J27':'L27']

# Limpar os dados no intervalo especificado
for linha in intervalo:
    for celula in linha:
        celula.value = None

# Array multidimensional com os novos dados


# Suponha que você queira inserir os novos dados a partir da célula I6
coluna_excel = 'J'
linha_excel = 27

for chave, valor in dict_vencimento_mes_ano.items():
    planilha[f'{coluna_excel}{linha_excel}'] = valor  # Adiciona o valor na célula atual
    coluna_excel = chr(ord(coluna_excel) + 1)  # Avança para a próxima coluna


intervalo = planilha['P23':'P23']

# Limpar os dados no intervalo especificado
for linha in intervalo:
    for celula in linha:
        celula.value = None

# Array multidimensional com os novos dados


# Suponha que você queira inserir os novos dados a partir da célula I6
coluna_excel = 'P'
linha_excel = 23

for chave, valor in dict_total_a_pagar.items():
    planilha[f'{coluna_excel}{linha_excel}'] = valor  # Adiciona o valor na célula atual
    coluna_excel = chr(ord(coluna_excel) + 1)  # Avança para a próxima coluna

#dict_parte_2_separada

intervalo = planilha['I31':'R32']

# Limpar os dados no intervalo especificado
for linha in intervalo:
    for celula in linha:
        celula.value = None



# Suponha que você queira inserir os novos dados a partir da célula I6

linha_excel = 31
for dados in dict_parte_2_separada.values():
    for item in dados:
        for chave, valor in item.items():
            planilha[f'I{linha_excel}'] = chave  # Adiciona a chave na célula atual
            coluna_excel = 'J'  # Começar a partir da coluna J
            for v in valor:
                planilha[f'{coluna_excel}{linha_excel}'] = v  # Adiciona cada valor na mesma linha
                coluna_excel = chr(ord(coluna_excel) + 1)  # Avança para a próxima coluna
        linha_excel += 1  # Avança para a próxima linha

wb.save(caminho_do_excel)