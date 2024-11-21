from tkinter import filedialog
from tkinter import *
import PyPDF2
#import sys
#sys.setdefaultencoding("utf-8")
caminho = filedialog.askopenfilename()
arquivo = open(caminho, 'r')
print(caminho)
print(arquivo.read())
arquivo.close()