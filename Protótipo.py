from tkinter import *
from tkinter import filedialog
from openpyxl import Workbook
import os
import csv
import PyPDF2

#Parte inicial do programa, recebe o PDF e extrai o texto das páginas.
try:
    diretorio = filedialog.askopenfilename(title="Selecione um arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")])
    with open(diretorio, 'rb') as arquivoPdf:
        leitor = PyPDF2.PdfReader(arquivoPdf)
        textoPdf = ""
        for pagina in leitor.pages:
            textoPdf += pagina.extract_text()
except Exception as e:
    print(f"Erro ao ler o PDF: {e}")
arquivoPdf.close()

#Função utilizada para gerar o Csv utilizando os dados em formato de lista aninhada.
def gerarCsv(nome_arquivo, dados):
    try:
        with open(nome_arquivo, mode='w', newline='', encoding='utf-8-sig') as arquivo_csv:
            escritor = csv.writer(arquivo_csv)
            escritor.writerow(["data_hora_entrada", "data_documento", "tipo_documento", "numero", "ano", "remetente", "documento", "indice", "prioridade", "carater", "distribuicao", "primdistribuicao", "redistribuicao", "refdoc", "processos_sei", "teor", "corpo", "resumo", "paises_ois", "pasta", "apenas_cumpre_instrucoes"])
            escritor.writerows(dados)
            caminhoCsv = os.path.abspath(nome_arquivo)
        print(f"Arquivo CSV criado em: {caminhoCsv}")
    except Exception as e:
        print(f"Erro ao criar o arquivo CSV: {e}")

gerarCsv("DPAGRO.csv",[[1,2,3]])

'''
Função para gerar um .txt sendo informado uma string, não permanecerá no final do projeto

def gerarTxt(textoPdf):
    try:
        nome_arquivo_txt = diretorio + "pagina.txt"
        with open(nome_arquivo_txt, 'w', encoding='utf-8-sig') as arquivo_txt:
            arquivo_txt.write(textoPdf)
            caminhoTxt = os.path.abspath(textoPdf)
        print(f"Texto salvo em: {caminhoTxt}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo.txt: {e}")
gerarTxt(textoPdf)
'''
