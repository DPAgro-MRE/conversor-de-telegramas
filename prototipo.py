from tkinter import *
from tkinter import filedialog
from openpyxl import Workbook
import os
import math
import re
import pytz
import csv
import PyPDF2
import pandas as pd
from datetime import datetime

def gerarCsv(nome_arquivo, dados):
    try:
        with open(nome_arquivo+".csv", mode='w', newline='', encoding='utf-8-sig') as arquivo_csv:
            escritor = csv.writer(arquivo_csv)
            escritor.writerow(["data_hora_entrada", "data_documento", "tipo_documento", "numero", "ano", "remetente", "documento", "indice", "prioridade", "carater", "distribuicao", "primdistribuicao", "redistribuicao", "primredistribuicao", "refdoc", "processos_sei", "teor", "corpo", "resumo", "paises_ois", "pasta", "apenas_cumpre_instrucoes"])
            escritor.writerows(dados)
            caminhoCsv = os.path.abspath(nome_arquivo+".csv")
        print(f"Arquivo CSV criado em: {caminhoCsv}")
    except Exception as e:
        print(f"Erro ao criar o arquivo CSV: {e}")

def gerarExcel(nome_arquivo, dados):
    try:
        workbook = Workbook()
        planilha = workbook.active
        planilha.title = "Coleção"
        planilha.append(["data_hora_entrada", "data_documento", "tipo_documento", "numero", "ano", "remetente", "documento", "indice", "prioridade", "carater", "distribuicao", "primdistribuicao", "redistribuicao", "primredistribuicao", "refdoc", "processos_sei", "teor", "corpo", "resumo", "paises_ois", "pasta", "apenas_cumpre_instrucoes"])
        for linha in dados:
            planilha.append(linha)
        workbook.save(nome_arquivo+".xlsx")
        caminhoExcel = os.path.abspath(nome_arquivo+".xlsx")
        print(f"Arquivo Excel criado em: {caminhoExcel}")
    except Exception as e:
        print(f"Erro ao criar o arquivo Excel: {e}")

def gerarTxt(textoPdf, NumTel):
    try:
        nome_arquivo_txt = (f"TEL{NumTel}.txt") 
        with open(nome_arquivo_txt, 'w', encoding='utf-8-sig') as arquivo_txt:
            if type(textoPdf) is list:
                for pagina in textoPdf:
                    arquivo_txt.write(str(pagina)+"\n")
            elif type(textoPdf) is str:
                arquivo_txt.write(textoPdf)
            caminhoTxt = os.path.abspath(nome_arquivo_txt)
        print(f"Texto salvo em: {caminhoTxt}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo .txt: {e}")
        
def Extracao(filepath, geraCsv, geraExcel):
    TELsValidos = [] #Lista com as páginas ordenadas dos telegramas de caráter Ostensivo e Reservado, desconsiderando os Secretos.

    #Parte inicial do programa, recebe o PDF e extrai o texto das páginas.
    try:
        #filepath = "C:/Users/eduardo.p.sousa/Downloads/TELDODIAT03122024.pdf"
        with open(filepath, 'rb') as arquivoPdf:
            leitor = PyPDF2.PdfReader(arquivoPdf)
            texto_pdf = ""
            for pagina in leitor.pages:
                texto_pdf += pagina.extract_text()
                if ("Para:" in texto_pdf) or ("CARAT=Secreto" in texto_pdf) or ('Distribuído' not in texto_pdf):
                    pass
                else:
                    TELsValidos.append(pagina.extract_text())
                texto_pdf = ""
    except Exception as e:
        print(f"Erro ao ler o PDF: {e}")

    #Segunda parte do programa, separa os telegramas válidos por grupos de página.
    TEL = []
    Dados = []

    for pagina in TELsValidos:
        numero_pagina = (((pagina.split("\n"))[0]).split(" "))[1]
        #Utiliza padrões de regex para extrair as informações das páginas.
        if int(numero_pagina[0])//int(numero_pagina[2]) == 1:
            TEL.append(pagina)
            TEL = '\n'.join(TEL)
            match_redistribuicao = re.search(r"Redistribuído para\s*(.*?)\s* em \d{2}/\d{2}/\d{4}", TEL) 
            if match_redistribuicao:
                Redistribuicao = match_redistribuicao.group(1) #CHECK
                prim_redistribuicao = Redistribuicao.split("/")[0] #CHECK
            else:
                Redistribuicao = "NA"
                prim_redistribuicao = "NA"
            match_distribuicao = re.search(r'DISTR=([A-Za-z/]+)', TEL)
            if match_distribuicao.group(1).lower().startswith("dpagro") or Redistribuicao.lower().startswith("dpagro"):
                match_corpo = re.search(r'Nr\. \d+.*', TEL, re.DOTALL)
                if match_corpo:
                    corpo = match_corpo.group(0)  # Pega o conteúdo extraído
                else:
                    corpo = ""
                    corpo = re.sub(r'Nr\. \d+\r?\n?', '', corpo)
                    corpo = re.sub(r'\r\n\r\n', '<pArAg>', corpo)
                    corpo = re.sub(r'\r\n', ' ', corpo)
                    corpo = re.sub(r'<pArAg>', '\r\n\r\n', corpo)
                Distribuicao = match_distribuicao.group(1) #CHECK
                match_remetente_e_data = re.search(r"De (.*?) para Exteriores em (\d{2}/\d{2}/\d{4})", TEL)
                Remetente = match_remetente_e_data.group(1)     #CHECK
                data_expedicao = match_remetente_e_data.group(2) #CHECK
                match_numero_tel = re.search(r"De: .*? Recebido em: (\d{2}/\d{2}/\d{4}) (\d{2}:\d{2}:\d{2}) N.°: (\d{5})", TEL)
                data_recebimento = match_numero_tel.group(1) #CHECK
                hora_entrada = match_numero_tel.group(2) #CHECK
                numero_tel = int(match_numero_tel.group(3)) #CHECK
                match_carater = re.search(r'CARAT=([A-Za-z]+)', TEL)
                match_prioridade = re.search(r'PRIOR=([A-Za-z]+)', TEL)
                Indice = (re.findall('//([\s\S]*?)//', TEL))[0].replace("\n", " ")
                Carater = match_carater.group(1) #CHECK
                Prioridade = match_prioridade.group(1) #CHECK
                primeira_distribuicao = "DPAGRO"
                data_e_hora = datetime.strptime(f"{data_recebimento} {hora_entrada}", "%d/%m/%Y %H:%M:%S") #CHECK
                Data = datetime.strptime(data_expedicao, "%d/%m/%Y") #CHECK
                Ano = Data.year #CHECK
                Documento = f"TEL {numero_tel}/{Ano}/{Remetente}" #CHECK
                match_resumo = re.search(r'RESUMO=\n(.*?)(\n\s*\n)', TEL, re.DOTALL)
                if match_resumo:
                    Resumo = match_resumo.group(1)
                else:
                    Resumo = ""
                match_instrucoes = re.search(r"(cumpre |cumpri )(instrução|instruções)", TEL, re.IGNORECASE)
                if match_instrucoes:
                    Instrucoes = "Sim"
                else:
                    Instrucoes = "Não"
                match_ref_doc = re.search(r'REF/ADIT=(.*)', TEL)
                if match_ref_doc:
                    refdoc = re.sub(r'(TEL [0-9]+|DET [0-9]+) ([0-9]{4})', r'\1/\2/<posto>', match_ref_doc.group(1))
                    refdoc = re.sub(r'(TEL [0-9]+|DET [0-9]+),', r'\1/<ano>/<posto>,', refdoc)
                    refdoc = re.sub(r'(TEL [0-9]+|DET [0-9]+)$', r'\1/<ano>/<posto>', refdoc)
                    refdoc = refdoc.replace('<ano>', str(Ano)).replace('<posto>', Remetente)
                else:
                    refdoc = "NA"
                match_processos = re.search(r"(\d{5}\.\d{6}/\d{4}-\d{2})", TEL)
                if match_processos:
                    Processos = match_processos.group(0)
                else:
                    Processos = "NA"
                Dados.append([data_e_hora, Data.date(), "TEL", numero_tel, Ano, Remetente, Documento, Indice, Prioridade, Carater, Distribuicao, primeira_distribuicao, Redistribuicao, prim_redistribuicao, refdoc, Processos, "teor", "corpo", Resumo, "paises_ois", "pasta", Instrucoes])
            TEL = []
        elif int(numero_pagina[0])//int(numero_pagina[2]) != 1:
            TEL.append(pagina)
    if geraExcel == 1:
        gerarExcel("TELEGRAMAS", Dados)
    if geraCsv == 1:
        gerarCsv("TELEGRAMAS", Dados)

Extracao(filedialog.askopenfilename(title="Selecione um arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")]), 1, 1)