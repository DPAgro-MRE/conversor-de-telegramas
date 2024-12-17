# Conversor de Telegramas
Programa de conversão de telegramas recebidos no formato pdf para o formato csv e xlsx, com o intuito de adaptar e inserir os dados na base de dados relacional do Portal DPAgro.

<div align="center">
  <img alt="Python" src="https://img.shields.io/badge/Python-3776AB?logo=python&logoColor=fff&style=for-the-badge">
</div>

<div align="center">
    <img alt="Tamanho do código do projeto" src="https://img.shields.io/github/languages/code-size/DPAgro-MRE/conversor-de-telegramas" />
</div>

## Descrição
Este repositório foi criado com o propósito de manter a transparência nas atividades técnicas da Divisão de Política Agrícola, aderir ao conceito de **desenvolvimento comunitário** e melhorar o processo de revisão e manutenção de código.


O sistema é capaz de:  
a) processar as informações referentes a telegramas, obtidas através da leitura de arquivos no formato pdf;  
b) disponibilizar os dados processados em um arquivo no formato xlsx e/ou csv.

## Dependências
- [Custom TKinter](https://customtkinter.tomschimansky.com/)
- [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)
- [PyPDF2](https://pypdf2.readthedocs.io/en/3.x/)


## Configurações iniciais
1. Ao clonar o repositório para as configurações iniciais das bibliotecas do projeto é necessário a execução do seguinte comando:
```sh
python -m venv venv
```

2. Ativar o ambiente virtual:
```sh
venv\Scripts\activate
```
3. Instalar as dependências necessárias do projeto executando o comando:
```sh
pip install -r requirements.txt
```
<h2>💻 Autores</h2>

<table>
  <tr>
    <td align="center"><a href="https://github.com/pedrosilv1514" target="_blank"><img style="border-radius: 50%;" src="https://github.com/pedrosilv1514.png" width="100px;" alt="Pedro Henrique"/><br /><sub><b>Pedro Henrique</b></sub></a><br/></td>
    <td align="center"><a href="https://github.com/Digs-LS" target="_blank"><img style="border-radius: 50%;" src="https://github.com/Digs-LS.png" width="100px;" alt="Diego Lucas"/><br /><sub><b>Diego Lucas</b></sub></a><br/></td>
    <td align="center"><a href="https://github.com/eduardopsousa" target="_blank"><img style="border-radius: 50%;" src="https://github.com/eduardopsousa.png" width="100px;" alt="Eduardo Pereira"/><br /><sub><b>Eduardo Pereira</b></sub></a><br/></td>
    <td align="center"><a href="https://github.com/Sara-Andrade" target="_blank"><img style="border-radius: 50%;" src="https://github.com/Sara-Andrade.png" width="100px;" alt="Sara Andrade"/><br /><sub><b>Sara Andrade</b></sub></a><br/></td>
</table>
