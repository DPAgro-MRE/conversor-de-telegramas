from tkinter import *
from tkinter import filedialog
from openpyxl import Workbook
import os
import re
import csv
import PyPDF2

TELsValidos = [] #Lista com as páginas ordenadas dos telegramas de caráter Ostensivo e Reservado, desconsiderando os Secretos.

#Parte inicial do programa, recebe o PDF e extrai o texto das páginas.
try:
    diretorio = filedialog.askopenfilename(title="Selecione um arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")])
    #diretorio = "C:/Users/eduardo.p.sousa/Downloads/GetFileAttachment.pdf"
    with open(diretorio, 'rb') as arquivoPdf:
        leitor = PyPDF2.PdfReader(arquivoPdf)
        textoPdf = ""
        for pagina in leitor.pages:
            textoPdf += pagina.extract_text()
            textoPdf = textoPdf.split("\n") #Separação da página por quebra de linhas
            if ("Para:" in textoPdf[1]) or ("CARAT=Secreto" in textoPdf[2]) or (len(textoPdf)<7):
                pass
            else:
                TELsValidos.append(pagina.extract_text())
            textoPdf = ""
except Exception as e:
    print(f"Erro ao ler o PDF: {e}")
arquivoPdf.close()

def gerarTxt(textoPdf):
    try:
        nome_arquivo_txt = os.path.basename(diretorio).replace(".pdf","") +f"{cont}" ".txt" 
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

#Segunda parte do programa, separa os telegramas válidos por grupos de página.
TEL = []
Dados = []
cont = 0 #temp
for pagina in TELsValidos:
    numero_pagina = (((pagina.split("\n"))[0]).split(" "))
    numero_pagina = numero_pagina[1]
    texto_Tel = ""
    #Pega primeiramente as informações que são garantidas de aparecer em ordem nos telegramas.
    if int(numero_pagina[0])//int(numero_pagina[2]) == 1:
        TEL.append(pagina)
        for i in range(len(TEL)):
            pagina_tel = TEL[i].split("\n")
            #pagina_tel[10].startswith("DISTR=DPAGRO")
            if i == 0:
                matchRemetenteData = re.search(r"De (.*?) para Exteriores em (\d{2}/\d{2}/\d{4})", pagina_tel[3])
                Remetente = matchRemetenteData.group(1)     #CHECK
                DataExpedicao = matchRemetenteData.group(2) #CHECK
                matchNumTel = re.search(r"De: .*? Recebido em: (\d{2}/\d{2}/\d{4}) (\d{2}:\d{2}:\d{2}) N.°: (\d{5})", pagina_tel[1])
                DataRecebimento = matchNumTel.group(1) #CHECK
                HoraEntrada = matchNumTel.group(2) #CHECK
                NumTel = matchNumTel.group(3) #CHECK
                CARAT = (pagina_tel[6].replace("CARAT=", ""))
                PRIOR = (pagina_tel[9].replace("PRIOR=", ""))
                DISTR = (pagina_tel[10].replace("DISTR=", ""))
                cont += 1 #temp
                gerarTxt(TEL)
        #Começar as operações com o telegrama
        TEL.clear()
    elif int(numero_pagina[0])//int(numero_pagina[2]) != 1:
        TEL.append(pagina)

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
#gerarTxt(TELsValidos)

'''
Função para gerar um .txt sendo informado uma string ou lista, não permanecerá no final do projeto e serve apenas para testes.

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