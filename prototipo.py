from tkinter import *
from tkinter import filedialog
from openpyxl import Workbook
import os
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
        
def Extracao():
    TELsValidos = [] #Lista com as páginas ordenadas dos telegramas de caráter Ostensivo e Reservado, desconsiderando os Secretos.
    #Parte inicial do programa, recebe o PDF e extrai o texto das páginas.
    try:
        diretorio = filedialog.askopenfilename(title="Selecione um arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        #diretorio = "C:/Users/eduardo.p.sousa/Downloads/TELDODIAT03122024.pdf"
        with open(diretorio, 'rb') as arquivoPdf:
            leitor = PyPDF2.PdfReader(arquivoPdf)
            textoPdf = ""
            for pagina in leitor.pages:
                textoPdf += pagina.extract_text()
                textosPdf = textoPdf.split("\n") #Separação da página por quebra de linhas
                if ("Para:" in textosPdf[1]) or ("CARAT=Secreto" in textosPdf[2]) or ('Distribuído' not in textoPdf):
                    pass
                else:
                    TELsValidos.append(pagina.extract_text())
                textoPdf = ""
    except Exception as e:
        print(f"Erro ao ler o PDF: {e}")
    arquivoPdf.close()

    #Segunda parte do programa, separa os telegramas válidos por grupos de página.
    TEL = []
    Dados = []
    for pagina in TELsValidos:
        numero_pagina = (((pagina.split("\n"))[0]).split(" "))
        numero_pagina = numero_pagina[1]
        texto_Tel = ""
        #Utiliza padrões de regex para extrair as informações das páginas.
        if int(numero_pagina[0])//int(numero_pagina[2]) == 1:
            TEL.append(pagina)
            for i in range(len(TEL)):
                if i == 0:
                    matchRedistribuicao = re.search(r"Redistribuído para\s*(.*?)\s* em \d{2}/\d{2}/\d{4}", TEL[i]) 
                    if matchRedistribuicao:
                        Redistribuicao = matchRedistribuicao.group(1) #CHECK
                        PrimRedistribuicao = Redistribuicao.split("/")[0]
                    else:
                        Redistribuicao = "NA"
                        PrimRedistribuicao = "NA"
                    matchDistr = re.search(r'DISTR=([A-Za-z/]+)', TEL[i])
                    if not matchDistr.group(1).startswith("DPAGRO") and not matchDistr.group(1).startswith("DPAgro"):
                        break
                    DISTR = matchDistr.group(1) #CHECK
                    matchRemetenteData = re.search(r"De (.*?) para Exteriores em (\d{2}/\d{2}/\d{4})", TEL[i])
                    Remetente = matchRemetenteData.group(1)     #CHECK
                    DataExpedicao = matchRemetenteData.group(2) #CHECK
                    matchNumTel = re.search(r"De: .*? Recebido em: (\d{2}/\d{2}/\d{4}) (\d{2}:\d{2}:\d{2}) N.°: (\d{5})", TEL[i])
                    DataRecebimento = matchNumTel.group(1) #CHECK
                    HoraEntrada = matchNumTel.group(2) #CHECK
                    NumTel = int(matchNumTel.group(3)) #CHECK
                    matchCarat = re.search(r'CARAT=([A-Za-z]+)', TEL[i])
                    matchPrior = re.search(r'PRIOR=([A-Za-z]+)', TEL[i])
                    indice = (re.findall('//([\s\S]*?)//', TEL[i]))[0].replace("\n", " ")
                    CARAT = matchCarat.group(1) #CHECK
                    PRIOR = matchPrior.group(1) #CHECK
                    PRIMDISTR = "DPAGRO"
                    DataHora = datetime.strptime(f"{DataRecebimento} {HoraEntrada}", "%d/%m/%Y %H:%M:%S") #CHECK
                    Data = datetime.strptime(DataExpedicao, "%d/%m/%Y") #CHECK
                    Ano = Data.year #CHECK
                    Documento = f"TEL {NumTel}/{Ano}/{Remetente}" #CHECK
                    matchInstrucoes = re.search(r"(cumpre |cumpri )(instrução|instruções)", TEL[i], re.IGNORECASE)
                    if matchInstrucoes:
                        Instrucoes = "Sim"
                    else:
                        Instrucoes = "Não"
                    Dados.append([DataHora, Data.date(), "TEL", NumTel, Ano, Remetente, Documento, indice, PRIOR, CARAT, DISTR, PRIMDISTR, Redistribuicao, PrimRedistribuicao, "refdoc", "processos_sei", "teor", "corpo","resumo","paises_ois","pasta", Instrucoes])
            #Começar as operações com o telegrama
            TEL.clear()
        elif int(numero_pagina[0])//int(numero_pagina[2]) != 1:
            TEL.append(pagina)
    gerarExcel("TELEGRAMAS", Dados)

Extracao()
'''
#Função utilizada para gerar o Csv utilizando os dados em formato de lista aninhada.
def gerarCsv(nome_arquivo, dados):
    try:
        with open(nome_arquivo+".csv", mode='w', newline='', encoding='utf-8-sig') as arquivo_csv:
            escritor = csv.writer(arquivo_csv)
            escritor.writerow(["data_hora_entrada", "data_documento", "tipo_documento", "numero", "ano", "remetente", "documento", "indice", "prioridade", "carater", "distribuicao", "primdistribuicao", "redistribuicao", "refdoc", "processos_sei", "teor", "corpo", "resumo", "paises_ois", "pasta", "apenas_cumpre_instrucoes"])
            escritor.writerows(dados)
            caminhoCsv = os.path.abspath(nome_arquivo+".csv")
        print(f"Arquivo CSV criado em: {caminhoCsv}")
    except Exception as e:
        print(f"Erro ao criar o arquivo CSV: {e}")

#Função utilizada para gerar a planilha Excel também utilizando os dados em formato de lista aninhada.
def gerarExcel(nome_arquivo, dados):
    try:
        workbook = Workbook()
        planilha = workbook.active
        planilha.title = "Coleção"
        planilha.append(["data_hora_entrada", "data_documento", "tipo_documento", "numero", "ano", "remetente", "documento", "indice", "prioridade", "carater", "distribuicao", "primdistribuicao", "redistribuicao", "refdoc", "processos_sei", "teor", "corpo", "resumo", "paises_ois", "pasta", "apenas_cumpre_instrucoes"])
        for linha in dados:
            planilha.append(linha)
        workbook.save(nome_arquivo+".xlsx")
        caminhoExcel = os.path.abspath(nome_arquivo+".xlsx")
        print(f"Arquivo Excel criado em: {caminhoExcel}")
    except Exception as e:
        print(f"Erro ao criar o arquivo Excel: {e}")

#Função para gerar um .txt sendo informado uma string ou lista, não permanecerá no final do projeto e serve apenas para testes.
def gerarTxt(textoPdf):
    try:
        nome_arquivo_txt = os.path.basename(diretorio).replace(".pdf","") + ".txt"
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
'''