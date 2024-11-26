from tkinter import *
from tkinter import filedialog
from openpyxl import Workbook
import os
import csv
import PyPDF2

TELsValidos = [] #Lista com as páginas ordenadas dos telegramas de caráter Ostensivo e Reservado, desconsiderando os Secretos.

#Parte inicial do programa, recebe o PDF e extrai o texto das páginas.
try:
    diretorio = filedialog.askopenfilename(title="Selecione um arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")])
    with open(diretorio, 'rb') as arquivoPdf:
        leitor = PyPDF2.PdfReader(arquivoPdf)
        textoPdf = ""
        for pagina in leitor.pages:
            textoPdf += pagina.extract_text()
            textoPdf = textoPdf.split("\n") #Separação da página por linhas
            if ("Para:" in textoPdf[1]) or ("CARAT=Secreto" in textoPdf[2]):
                pass
            else:
                TELsValidos.append(pagina.extract_text())
            textoPdf = ""
except Exception as e:
    print(f"Erro ao ler o PDF: {e}")
arquivoPdf.close()

#Segunda parte do programa, separa os telegramas válidos por grupos de página.


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

'''
Função para gerar um .txt sendo informado uma string ou lista, não permanecerá no final do projeto

def gerarTxt(textoPdf):
    try:
        nome_arquivo_txt = diretorio + ".txt"
        with open(nome_arquivo_txt, 'w', encoding='utf-8-sig') as arquivo_txt:
            if type(textoPdf) is list:
                for pagina in textoPdf:
                    arquivo_txt.write(str(pagina)+"\n")
            elif type(textoPdf) is str:
                arquivo_txt.write(textoPdf)
            caminhoTxt = os.path.abspath(diretorio)
        print(f"Texto salvo em: {caminhoTxt}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo .txt: {e}")
'''
