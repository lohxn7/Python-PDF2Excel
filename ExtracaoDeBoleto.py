import PyPDF2
import re
import pandas as pd
import os
from openpyxl import load_workbook
from collections import defaultdict


caminho_do_excel = "mascara.xlsx"
caminho_do_pdf = "Fatura1.pdf"



def extrair_texto_pdf(caminho_arquivo):
    texto = " "
    with open(caminho_arquivo, 'rb') as arquivo:
        leitor_pdf = PyPDF2.PdfReader(arquivo)
        num_paginas = len(leitor_pdf.pages)
        for pagina in range(num_paginas):
            texto += f" {leitor_pdf.pages[pagina].extract_text()}"
    return texto

def separar_palavras(texto):
    
    texto_modificado = re.sub(r'(?<=\d),(?=\D)', r', ', texto)
    texto_modificado = re.sub(r'(?<=\d)(?=[A-Z])', r' ', texto_modificado)
    return texto_modificado



texto_extraido = extrair_texto_pdf(caminho_do_pdf)


texto_emlista = texto_extraido.split("\n")


texto_original_doValoraPagar = texto_emlista[13]
texto_modificado_doValoraPagar = f"{separar_palavras(texto_original_doValoraPagar)}"
lista_resultadoDoValoraPagar = list((texto_modificado_doValoraPagar.split())[2:4])
resultadoDoValoraPagar = [' '.join(lista_resultadoDoValoraPagar)]


lista_do_mesAno = [texto_modificado_doValoraPagar.split()[0]]


lista_do_Vencimento = [texto_modificado_doValoraPagar.split()[1]]


texto_original_da_leitura = texto_emlista[17]
texto_modificado_da_leitura = f'{separar_palavras(texto_original_da_leitura)}'
lista_resultado_daleitura = list((texto_modificado_da_leitura.split()))
def ajustar_data_com_codigo(lista_resultado_daleitura):
    
    nova_lista = []
    for item in lista_resultado_daleitura:
        
        if re.match(r'^\d{11}/\d{2}/\d{4}$', item):
            
            item_modificado = item[:9] + " " + item[9:]
            nova_lista.append(item_modificado.split(" ")[1])
        else:
            nova_lista.append(item)
    return nova_lista
lista_modificada_da_leitura = ajustar_data_com_codigo(lista_resultado_daleitura[-4:])


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


dict_medida = {
    'Medidor': medidor[0],
    'Grandezas': medidor[1],
    'Postos Tarifários': medidor[2],
    'Leitura Anterior': medidor[3],
    'Leitura Atual': medidor[4],
    'Const. Medidor': medidor[5],
    'Consumo kWh/kw': medidor[6],
}


energia_lista = []

for texto in texto_emlista:
    for i in range(len(texto)):
        if texto[i:i+10] == "Energia At":
            i+=10
            energia_lista.append(texto.split()) 

cip_lista = []
for texto in texto_emlista:
    for i in range(len(texto)):
        if texto[i:i+10] == "CIP ILUM P":
            i+=10
            cip_lista.append(texto.split()) 


for frase_cip in cip_lista:
    cip = list()
    cip.append((" ".join(frase_cip[0:-5]),frase_cip[-5:]))     
    dist_cip = dict(cip)
    for i in range(3):
        dist_cip["CIP ILUM PUB PREF MUNICIPAL"].insert(0,"") 
    
    dist_cip["CIP ILUM PUB PREF MUNICIPAL"].append("") 

lista_dist_energia = list()
for frase in energia_lista:
    
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


nova_lista_dist_energia = []
for item in lista_dist_energia:
    novo_dict = {}
    for chave, valores in item.items():
        nova_lista_valores = []
        for valor in valores:
            if valor.endswith('-'):
                
                novo_valor = '-' + valor[:-1]
                nova_lista_valores.append(novo_valor)
            else:
                nova_lista_valores.append(valor)
        novo_dict[chave] = nova_lista_valores
    nova_lista_dist_energia.append(novo_dict)



def parse_number(value):
    if isinstance(value, str):  
        value = value.strip()  
        if value == '':
            return ''  
        if '%' in value:
            
            value = value.replace('%', '').replace(',', '.')
            
            return float(value) / 100
        else:
            
            value = re.sub(r'\.(?=\d{3}(?:,|$))', '', value)  
            value = value.replace(',', '.')
            return float(value)  
    return value  

def transform_dict_values(dictionaries):
    for dictionary in dictionaries:
        for key, values in dictionary.items():
            dictionary[key] = [values[0]] + [parse_number(value) for value in values[1:]]


transform_dict_values(nova_lista_dist_energia)

def converter_valor_para_float(valor_str):
    
    valor_str = valor_str.replace('R$', '').strip()
    
    valor_str = valor_str.replace('.', '')  
    valor_str = valor_str.replace(',', '.')  
    
    valor_float = float(valor_str)
    return valor_float


for chave, valor in dict_total_a_pagar.items():
    dict_total_a_pagar[chave] = converter_valor_para_float(valor)

"""
""" 




agrupados_por_kWh = defaultdict(list)


for dicionario in nova_lista_dist_energia:
    for chave, valor in dicionario.items():
        quantidade = valor[1]  
        agrupados_por_kWh[quantidade].append({chave: valor})


lista_agrupada = list(agrupados_por_kWh.values())


dict_resultante = {}


for indice, parte in enumerate(lista_agrupada):
    dict_resultante[f'parte_{indice + 1}'] = parte


dict_parte_2_separada = {}


if 'parte_2' in dict_resultante:
    dict_parte_2_separada['parte_2'] = dict_resultante.pop('parte_2')


wb = load_workbook(caminho_do_excel)


planilha = wb.active

intervalo = planilha['I7':'R15']


for linha in intervalo:
    for celula in linha:
        celula.value = None

linha_excel = 7


for dados in dict_resultante.values():
    for item in dados:
        for chave, valor in item.items():
            planilha[f'I{linha_excel}'] = chave  
            coluna_excel = 'J'  
            for v in valor:
                planilha[f'{coluna_excel}{linha_excel}'] = v  
                coluna_excel = chr(ord(coluna_excel) + 1)  
        linha_excel += 1  





intervalo = planilha['J21':'P21']


for linha in intervalo:
    for celula in linha:
        celula.value = None





coluna_excel = 'J'
linha_excel = 21

for chave, valor in dict_medida.items():
    planilha[f'{coluna_excel}{linha_excel}'] = valor  
    coluna_excel = chr(ord(coluna_excel) + 1)  




intervalo = planilha['J24':'M24']


for linha in intervalo:
    for celula in linha:
        celula.value = None





coluna_excel = 'J'
linha_excel = 24

for chave, valor in dict_leitura.items():
    planilha[f'{coluna_excel}{linha_excel}'] = valor  
    coluna_excel = chr(ord(coluna_excel) + 1)  


intervalo = planilha['J27':'L27']


for linha in intervalo:
    for celula in linha:
        celula.value = None





coluna_excel = 'J'
linha_excel = 27

for chave, valor in dict_vencimento_mes_ano.items():
    planilha[f'{coluna_excel}{linha_excel}'] = valor  
    coluna_excel = chr(ord(coluna_excel) + 1)  


intervalo = planilha['P23':'P23']


for linha in intervalo:
    for celula in linha:
        celula.value = None





coluna_excel = 'P'
linha_excel = 23

for chave, valor in dict_total_a_pagar.items():
    planilha[f'{coluna_excel}{linha_excel}'] = valor  
    coluna_excel = chr(ord(coluna_excel) + 1)  



intervalo = planilha['I31':'R32']


for linha in intervalo:
    for celula in linha:
        celula.value = None





linha_excel = 31
for dados in dict_parte_2_separada.values():
    for item in dados:
        for chave, valor in item.items():
            planilha[f'I{linha_excel}'] = chave  
            coluna_excel = 'J'  
            for v in valor:
                planilha[f'{coluna_excel}{linha_excel}'] = v  
                coluna_excel = chr(ord(coluna_excel) + 1)  
        linha_excel += 1  

wb.save(caminho_do_excel)