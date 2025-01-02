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
        pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        caminho_csv = os.path.join(pasta_downloads, nome_arquivo + ".csv")
        with open(caminho_csv, mode='w', newline='', encoding='utf-8-sig') as arquivo_csv:
            escritor = csv.writer(arquivo_csv)
            escritor.writerow(["data_hora_entrada", "data_documento", "tipo_documento", "numero", "ano", "remetente", "documento", "indice", "prioridade", "carater", "distribuicao", "primdistribuicao", "redistribuicao", "primredistribuicao", "refdoc", "processos_sei", "teor", "corpo", "resumo", "paises_ois", "pasta", "apenas_cumpre_instrucoes"])
            escritor.writerows(dados)
        print(f"Arquivo CSV criado em: {caminho_csv}")
    except Exception as e:
        print(f"Erro ao criar o arquivo CSV: {e}")

def gerarExcel(nome_arquivo, dados):
    try:
        pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        caminho_arquivo = os.path.join(pasta_downloads, nome_arquivo + ".xlsx")
        workbook = Workbook()
        planilha = workbook.active
        planilha.title = "Coleção"
        planilha.append(["data_hora_entrada", "data_documento", "tipo_documento", "numero", "ano", "remetente", "documento", "indice", "prioridade", "carater", "distribuicao", "primdistribuicao", "redistribuicao", "primredistribuicao", "refdoc", "processos_sei", "teor", "corpo", "resumo", "paises_ois", "pasta", "apenas_cumpre_instrucoes"])
        for linha in dados:
            planilha.append(linha)
        workbook.save(caminho_arquivo)
        print(f"Arquivo Excel criado em: {caminho_arquivo}")
    except Exception as e:
        print(f"Erro ao criar o arquivo Excel: {e}")

def gerarTxt(textoPdf, NumTel):
    #Função de debug, não utilizada para gerar o .xlsx ou .csv.
    try:
        nome_arquivo_txt = (f"TEL{NumTel}.txt") 
        with open(nome_arquivo_txt, 'w', encoding='utf-8-sig') as arquivo_txt:
            if type(textoPdf) is list:
                for pagina in textoPdf:
                    arquivo_txt.write(str(pagina)+"\n")
            elif type(textoPdf) is str:
                arquivo_txt.write(textoPdf)
            caminhoTxt = os.path.join(os.path.expanduser("~"), "Downloads")
        print(f"Texto salvo em: {caminhoTxt}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo .txt: {e}")

def Extracao(filepath, geraCsv, geraExcel):
    dicioDetecta = {'Afeganistão': 'Afeganistão', 'África do Sul': 'África do Sul', 'ALADI': 'aladi', 'Albânia': 'Albânia', 'Alemanha': 'Alemanha', 'Andorra': 'Andorra', 'Angola': 'Angola', 'ANGUILA': 'ANGUILA', 'Antígua e Barbuda': 'Antígua e Barbuda', 'Arábia Saudita': 'Arábia Saudita', 'Argélia': 'Argélia', 'Argentina': 'Argentina', 'Armênia': 'Armênia', 'Aruba': 'Aruba', 'Austrália': 'Austrália', 'Áustria': 'Áustria', 'Azerbaijão': 'Azerbaijão', 'Bahamas': 'Bahamas', 'Bahrein': 'Bahrein', 'Bangladesh': 'Bangladesh', 'Barbados': 'Barbados', 'Belarus': 'Belarus', 'Bélgica': 'Bélgica', 'Belize': 'Belize', 'Benin': 'Benin', 'BERMUDA': 'BERMUDA', 'Bolívia': 'Bolívia', 'BONAIRE, SINT EUSTATIUS E SABA': 'BONAIRE, SINT EUSTATIUS E SABA', 'Bósnia e Herzegovina': 'Bósnia e Herzegovina', 'Botsuana': 'Botsuana', 'Brunei': 'Brunei', 'Bulgária': 'Bulgária', 'Burkina Faso': 'Burkina Faso', 'Burundi': 'Burundi', 'Butão': 'Butão', 'Cabo Verde': 'Cabo Verde', 'Camboja': 'Camboja', 'Cameroun (Camarões)': 'Cameroun (Camarões)', 'Cameroun': 'Cameroun (Camarões)', 'Canadá': 'Canadá', 'CARICOM': 'CARICOM', 'Catar': 'Catar', 'Cazaquistão': 'Cazaquistão', 'CDB': 'CDB', 'CEEA': 'CEEA', 'CFC/FCPB': 'CFC/FCPB', 'Chade': 'Chade', 'Chile': 'Chile', 'China': 'China', 'Chipre': 'Chipre', 'CIPV': 'CIPV', 'Codex Alimentarius': 'Codex Alimentarius', 'COI': 'COI', 'Colômbia': 'Colômbia', 'Comores': 'Comores', 'Congo': 'Congo', 'Coreia do Norte': 'Coreia do Norte', 'Coreia do Sul': 'Coreia do Sul', 'Costa Rica': 'Costa Rica', "Côte D'Ivoire (Costa do Marfim)": "Côte D'Ivoire (Costa do Marfim)", "Côte D'Ivoire": "Côte D'Ivoire (Costa do Marfim)", 'Costa do Marfim': "Côte D'Ivoire (Costa do Marfim)", 'CPLP': 'cplp', 'Croácia': 'Croácia', 'Cuba': 'Cuba', 'CURAÇAO': 'CURAÇAO', 'Dinamarca': 'Dinamarca', 'Djibouti (Djibuti)': 'Djibouti (Djibuti)', 'Djibouti': 'Djibouti (Djibuti)', 'Djibuti': 'Djibouti (Djibuti)', 'Dominica': 'Dominica', 'Egito': 'Egito', 'El Salvador': 'El Salvador', 'Emirados Árabes': 'Emirados Árabes', 'Emirados Árabes Unidos': 'Emirados Árabes', 'EAU': 'Emirados Árabes', 'Equador': 'Equador', 'Eritreia': 'Eritreia', 'Eslováquia': 'Eslováquia', 'Eslovênia': 'Eslovênia', 'Espanha': 'Espanha', 'Estados Unidos': 'Estados Unidos', 'EUA': 'Estados Unidos', 'Estônia': 'Estônia', 'Etiópia': 'Etiópia', 'FAO': 'FAO', 'Fiji': 'Fiji', 'Filipinas': 'Filipinas', 'Finlândia': 'Finlândia', 'França': 'França', 'Gabão': 'Gabão', 'Gâmbia': 'Gâmbia', 'Gana': 'Gana', 'Geórgia': 'Geórgia', 'GEÓRGIA DO SUL E AS ILHAS SANDWICH': 'GEÓRGIA DO SUL E AS ILHAS SANDWICH', 'GIBRALTAR': 'GIBRALTAR', 'Granada': 'Granada', 'Grécia': 'Grécia', 'GROENLÂNDIA': 'GROENLÂNDIA', 'GUADALUPE': 'GUADALUPE', 'GUAM': 'GUAM', 'Guatemala': 'Guatemala', 'GUERNSEY': 'GUERNSEY', 'Guiana': 'Guiana', 'GUIANA FRANCESA': 'GUIANA FRANCESA', 'Guiné Equatorial': 'Guiné Equatorial', 'Guiné-Bissau': 'Guiné-Bissau', 'Guiné-Conacri': 'Guiné-Conacri', 'Guiné': 'Guiné-Conacri', 'Haiti': 'Haiti', 'Honduras': 'Honduras', 'Hong Kong': 'Hong Kong', 'Hungria': 'Hungria', 'ICAC': 'ICAC', 'Iêmen': 'Iêmen', 'ILHA DE BOUVET': 'ILHA DE BOUVET', 'ILHA DE MAN': 'ILHA DE MAN', 'ILHA DO NATAL': 'ILHA DO NATAL', 'ILHA NORFOLK': 'ILHA NORFOLK', 'ILHAS CAYMAN': 'ILHAS CAYMAN', 'ILHAS COCOS': 'ILHAS COCOS', 'ILHAS COOK': 'ILHAS COOK', 'ILHAS DE ÅLAND': 'ILHAS DE ÅLAND', 'ILHAS FAROE': 'ILHAS FAROE', 'ILHAS HEARD E ILHAS McDONALD': 'ILHAS HEARD E ILHAS McDONALD', 'ILHAS MALVINAS (FALKLAND)': 'ILHAS MALVINAS (FALKLAND)', 'ILHAS MARIANAS DO NORTE': 'ILHAS MARIANAS DO NORTE', 'Ilhas Marshall': 'Ilhas Marshall', 'ILHAS OUTONARES MENORES DOS ESTADOS UNIDOS': 'ILHAS OUTONARES MENORES DOS ESTADOS UNIDOS', 'Ilhas Salomão': 'Ilhas Salomão', 'ILHAS VIRGENS (BRITÂNICAS)': 'ILHAS VIRGENS (BRITÂNICAS)', 'ILHAS VIRGENS (EUA)': 'ILHAS VIRGENS (EUA)', 'Índia': 'Índia', 'Indonésia': 'Indonésia', 'Irã': 'Irã', 'Iraque': 'Iraque', 'Irlanda': 'Irlanda', 'Islândia': 'Islândia', 'Israel': 'Israel', 'Itália': 'Itália', 'Jamaica': 'Jamaica', 'Japão': 'Japão', 'JERSEY': 'JERSEY', 'Jordânia': 'Jordânia', 'Kiribati': 'Kiribati', 'Kuaite (Kuwait)': 'Kuaite (Kuwait)', 'Kuaite': 'Kuaite (Kuwait)', 'Kuwait': 'Kuaite (Kuwait)', 'Laos': 'Laos', 'Lesoto': 'Lesoto', 'Letônia': 'Letônia', 'Líbano': 'Líbano', 'Libéria': 'Libéria', 'Líbia': 'Líbia', 'Liechtenstein': 'Liechtenstein', 'Lituânia': 'Lituânia', 'Luxemburgo': 'Luxemburgo', 'MACAU': 'MACAU', 'Macedônia do Norte': 'Macedônia do Norte', 'Madagascar': 'Madagascar', 'Malásia': 'Malásia', 'Malauí (Malawi)': 'Malauí (Malawi)', 'Malauí': 'Malauí (Malawi)', 'Malawi': 'Malauí (Malawi)', 'Maldivas': 'Maldivas', 'Mali': 'Mali', 'Malta': 'Malta', 'Marrocos': 'Marrocos', 'MARTINICA': 'MARTINICA', 'Maurício': 'Maurício', 'Mauritânia': 'Mauritânia', 'MAYOTTE': 'MAYOTTE', 'MERCOSUL': 'MERCOSUL', 'México': 'México', 'Mianmar (Myanmar)': 'Mianmar (Myanmar)', 'Mianmar': 'Mianmar (Myanmar)', 'Myanmar': 'Mianmar (Myanmar)', 'Micronésia': 'Micronésia', 'Moçambique': 'Moçambique', 'Moldova (Moldávia)': 'Moldova (Moldávia)', 'Moldova': 'Moldova (Moldávia)', 'Moldávia': 'Moldova (Moldávia)', 'Mônaco': 'Mônaco', 'Mongólia': 'Mongólia', 'Montenegro': 'Montenegro', 'MONTSERRAT': 'MONTSERRAT', 'Namíbia': 'Namíbia', 'Nauru': 'Nauru', 'Nepal': 'Nepal', 'Nicarágua': 'Nicarágua', 'Níger': 'Níger', 'Nigéria': 'Nigéria', 'NIUE': 'NIUE', 'Noruega': 'Noruega', 'NOVA CALEDÔNIA': 'NOVA CALEDÔNIA', 'Nova Zelândia': 'Nova Zelândia', 'OCDE': 'OCDE', 'OEA': 'OEA', 'OIAçúcar': 'OIAçúcar', 'Organização Internacional do Açúcar': 'OIAçúcar', 'OICacau/ICCO': 'OICacau/ICCO', 'OICacau': 'OICacau/ICCO', 'ICCO': 'OICacau/ICCO', 'OICafé': 'OICafé', 'Organização Internacional do Café': 'OICafé', 'OIE': 'OIE', 'OIV': 'OIV', 'Omã': 'Omã', 'OMC': 'OMC', 'OMS': 'OMS', 'ONU': 'ONU', 'Nações Unidas': 'ONU', 'Países Baixos': 'Países Baixos', 'Palau': 'Palau', 'Palestina': 'Palestina', 'Panamá': 'Panamá', 'Papua Nova Guiné': 'Papua Nova Guiné', 'Paquistão': 'Paquistão', 'Paraguai': 'Paraguai', 'Peru': 'Peru', 'PITCAIRN': 'PITCAIRN', 'POLINÉSIA FRANCESA': 'POLINÉSIA FRANCESA', 'Polônia': 'Polônia', 'PORTO RICO': 'PORTO RICO', 'Portugal': 'Portugal', 'Quênia': 'Quênia', 'Quirguistão': 'Quirguistão', 'Reino Unido': 'Reino Unido', 'Rep. Centro-Africana': 'Rep. Centro-Africana', 'Rep. Dem. do Congo': 'Rep. Dem. do Congo', 'República Dominicana': 'República Dominicana', 'Rep. Dom.': 'República Dominicana', 'RepDom': 'República Dominicana', 'República Tcheca': 'República Tcheca', 'RÉUNION': 'RÉUNION', 'Romênia': 'Romênia', 'Ruanda': 'Ruanda', 'Rússia': 'Rússia', 'SAARA OCIDENTAL': 'SAARA OCIDENTAL', 'SAINT BARTHÉLEMY': 'SAINT BARTHÉLEMY', 'SAINT HELENA, ASCENSION E TRISTAN DA CUNHA': 'SAINT HELENA, ASCENSION E TRISTAN DA CUNHA', 'Samoa': 'Samoa', 'SAMOA AMERICANA': 'SAMOA AMERICANA', 'San Marino': 'San Marino', 'Santa Lúcia': 'Santa Lúcia', 'Santa Sé (Vaticano)': 'Santa Sé (Vaticano)', 'Santa Sé': 'Santa Sé (Vaticano)', 'Vaticano': 'Santa Sé (Vaticano)', 'São Cristóvão e Névis': 'São Cristóvão e Névis', 'SÃO MARTINHO (PARTE FRANCESA)': 'SÃO MARTINHO (PARTE FRANCESA)', 'SÃO PIERRE E MIQUELON': 'SÃO PIERRE E MIQUELON', 'São Tomé e Principe': 'São Tomé e Principe', 'São Vicente e Granadinas': 'São Vicente e Granadinas', 'Seichelles (Seychelles)': 'Seichelles (Seychelles)', 'Seichelles': 'Seichelles (Seychelles)', 'Seychelles': 'Seichelles (Seychelles)', 'Seicheles': 'Seichelles (Seychelles)', 'Senegal': 'Senegal', 'Serra Leoa': 'Serra Leoa', 'Sérvia': 'Sérvia', 'Singapura (Cingapura)': 'Singapura (Cingapura)', 'Singapura': 'Singapura (Cingapura)', 'Cingapura': 'Singapura (Cingapura)', 'SINT MAARTEN (PARTE HOLANDESA)': 'SINT MAARTEN (PARTE HOLANDESA)', 'Síria': 'Síria', 'Somália': 'Somália', 'Sri Lanka': 'Sri Lanka', 'Suazilândia': 'Suazilândia', 'Sudão': 'Sudão', 'Sudão do Sul': 'Sudão do Sul', 'Suécia': 'Suécia', 'Suíça': 'Suíça', 'Suriname': 'Suriname', 'SVALBARD E JAN MAYEN': 'SVALBARD E JAN MAYEN', 'Tadjiquistão': 'Tadjiquistão', 'Tailândia': 'Tailândia', 'Taiwan': 'Taiwan', 'Tanzânia': 'Tanzânia', 'TERRITÓRIO OCEANO BRITÂNICO (THE)': 'TERRITÓRIO OCEANO BRITÂNICO (THE)', 'TERRITÓRIOS DO SUL FRANCÊS': 'TERRITÓRIOS DO SUL FRANCÊS', 'Timor-Leste (Timor Leste)': 'Timor-Leste (Timor Leste)', 'Timor-Leste': 'Timor-Leste (Timor Leste)', 'Timor Leste': 'Timor-Leste (Timor Leste)', 'Togo': 'Togo', 'TOKELAU': 'TOKELAU', 'Tonga': 'Tonga', 'Trinidad e Tobago': 'Trinidad e Tobago', 'Tunísia': 'Tunísia', 'Turcomenistão': 'Turcomenistão', 'TURKS E CAICOS ISLANDS': 'TURKS E CAICOS ISLANDS', 'Turquia': 'Turquia', 'Tuvalu': 'Tuvalu', 'Ucrânia': 'Ucrânia', 'Uganda': 'Uganda', 'União Europeia': 'União Europeia', 'UE': 'União Europeia', 'Uruguai': 'Uruguai', 'Uzbequistão': 'Uzbequistão', 'Vanuatu': 'Vanuatu', 'VÁRIOS PAÍSES': 'VÁRIOS PAÍSES', 'Venezuela': 'Venezuela', 'Vietnã': 'Vietnã', 'WALLIS E FUTUNA': 'WALLIS E FUTUNA', 'Zâmbia': 'Zâmbia', 'Zimbábue': 'Zimbábue'}
    # ^ Dicionário com os países que aparecem no índice dos telegramas, sendo a Key do dicionário o nome do país a ser identificado no texto, e o seu Value o nome que ficará na coluna paises_ois da planilha.
    
    dicioPasta = {'Afeganistão': 'afeganistao', 'África do Sul': 'africa_do_sul', 'aladi': 'aladi', 'Albânia': 'albania', 'Alemanha': 'alemanha', 'Andorra': 'andorra', 'Angola': 'angola', 'ANGUILA': 'anguila', 'Antígua e Barbuda': 'antigua_e_barbuda', 'Arábia Saudita': 'arabia_saudita', 'Argélia': 'argelia', 'Argentina': 'argentina', 'Armênia': 'armenia', 'Aruba': 'aruba', 'Austrália': 'australia', 'Áustria': 'austria', 'Azerbaijão': 'azerbaijao', 'Bahamas': 'bahamas', 'Bahrein': 'bahrein', 'Bangladesh': 'bangladesh', 'Barbados': 'barbados', 'Belarus': 'belarus', 'Bélgica': 'belgica', 'Belize': 'belize', 'Benin': 'benin', 'BERMUDA': 'bermuda', 'Bolívia': 'bolivia', 'BONAIRE, SINT EUSTATIUS E SABA': 'bonaire_sint_eustatius_e_saba', 'Bósnia e Herzegovina': 'bosnia_e_herzegovina', 'Botsuana': 'botsuana', 'Brunei': 'brunei', 'Bulgária': 'bulgaria', 'Burkina Faso': 'burkina_faso', 'Burundi': 'burundi', 'Butão': 'butao', 'Cabo Verde': 'cabo_verde', 'Camboja': 'camboja', 'Cameroun (Camarões)': 'cameroun_camaroes', 'Canadá': 'canada', 'CARICOM': 'varios_paises', 'Catar': 'catar', 'Cazaquistão': 'cazaquistao', 'CDB': 'cdb', 'CEEA': 'ceea', 'CFC/FCPB': 'cfcfcpb', 'Chade': 'chade', 'Chile': 'chile', 'China': 'china', 'Chipre': 'chipre', 'CIPV': 'cipv', 'Codex Alimentarius': 'codex_alimentarius', 'COI': 'coi', 'Colômbia': 'colombia', 'Comores': 'comores', 'Congo': 'congo', 'Coreia do Norte': 'coreia_do_norte', 'Coreia do Sul': 'coreia_do_sul', 'Costa Rica': 'costa_rica', "Côte D'Ivoire (Costa do Marfim)": 'cote_divoire_costa_do_marfim', 'cplp': 'cplp', 'Croácia': 'croacia', 'Cuba': 'cuba', 'CURAÇAO': 'curacao', 'Dinamarca': 'dinamarca', 'Djibouti (Djibuti)': 'djibouti_djibuti', 'Dominica': 'dominica', 'Egito': 'egito', 'El Salvador': 'el_salvador', 'Emirados Árabes': 'emirados_arabes', 'Equador': 'equador', 'Eritreia': 'eritreia', 'Eslováquia': 'eslovaquia', 'Eslovênia': 'eslovenia', 'Espanha': 'espanha', 'Estados Unidos': 'estados_unidos', 'Estônia': 'estonia', 'Etiópia': 'etiopia', 'FAO': 'fao', 'Fiji': 'fiji', 'Filipinas': 'filipinas', 'Finlândia': 'finlandia', 'França': 'franca', 'Gabão': 'gabao', 'Gâmbia': 'gambia', 'Gana': 'gana', 'Geórgia': 'georgia', 'GEÓRGIA DO SUL E AS ILHAS SANDWICH': 'georgia_do_sul_e_as_ilhas_sandwich', 'GIBRALTAR': 'gibraltar', 'Granada': 'granada', 'Grécia': 'grecia', 'GROENLÂNDIA': 'groenlandia', 'GUADALUPE': 'guadalupe', 'GUAM': 'guam', 'Guatemala': 'guatemala', 'GUERNSEY': 'guernsey', 'Guiana': 'guiana', 'GUIANA FRANCESA': 'guiana_francesa', 'Guiné Equatorial': 'guine_equatorial', 'Guiné-Bissau': 'guinebissau', 'Guiné-Conacri': 'guineconacri', 'Haiti': 'haiti', 'Honduras': 'honduras', 'Hong Kong': 'hong_kong', 'Hungria': 'hungria', 'ICAC': 'icac', 'Iêmen': 'iemen', 'ILHA DE BOUVET': 'ilha_de_bouvet', 'ILHA DE MAN': 'ilha_de_man', 'ILHA DO NATAL': 'ilha_do_natal', 'ILHA NORFOLK': 'ilha_norfolk', 'ILHAS CAYMAN': 'ilhas_cayman', 'ILHAS COCOS': 'ilhas_cocos', 'ILHAS COOK': 'ilhas_cook', 'ILHAS DE ÅLAND': 'ilhas_de_åland', 'ILHAS FAROE': 'ilhas_faroe', 'ILHAS HEARD E ILHAS McDONALD': 'ilhas_heard_e_ilhas_mcdonald', 'ILHAS MALVINAS (FALKLAND)': 'ilhas_malvinas_falkland', 'ILHAS MARIANAS DO NORTE': 'ilhas_marianas_do_norte', 'Ilhas Marshall': 'ilhas_marshall', 'ILHAS OUTONARES MENORES DOS ESTADOS UNIDOS': 'ilhas_outonares_menores_dos_estados_unidos', 'Ilhas Salomão': 'ilhas_salomao', 'ILHAS VIRGENS (BRITÂNICAS)': 'ilhas_virgens_britanicas', 'ILHAS VIRGENS (EUA)': 'ilhas_virgens_eua', 'Índia': 'india', 'Indonésia': 'indonesia', 'Irã': 'ira', 'Iraque': 'iraque', 'Irlanda': 'irlanda', 'Islândia': 'islandia', 'Israel': 'israel', 'Itália': 'italia', 'Jamaica': 'jamaica', 'Japão': 'japao', 'JERSEY': 'jersey', 'Jordânia': 'jordania', 'Kiribati': 'kiribati', 'Kuaite (Kuwait)': 'kuaite_kuwait', 'Laos': 'laos', 'Lesoto': 'lesoto', 'Letônia': 'letonia', 'Líbano': 'libano', 'Libéria': 'liberia', 'Líbia': 'libia', 'Liechtenstein': 'liechtenstein', 'Lituânia': 'lituania', 'Luxemburgo': 'luxemburgo', 'MACAU': 'macau', 'Macedônia do Norte': 'macedonia_do_norte', 'Madagascar': 'madagascar', 'Malásia': 'malasia', 'Malauí (Malawi)': 'malaui_malawi', 'Maldivas': 'maldivas', 'Mali': 'mali', 'Malta': 'malta', 'Marrocos': 'marrocos', 'MARTINICA': 'martinica', 'Maurício': 'mauricio', 'Mauritânia': 'mauritania', 'MAYOTTE': 'mayotte', 'MERCOSUL': 'mercosul', 'México': 'mexico', 'Mianmar (Myanmar)': 'mianmar_myanmar', 'Micronésia': 'micronesia', 'Moçambique': 'mocambique', 'Moldova (Moldávia)': 'moldova_moldavia', 'Mônaco': 'monaco', 'Mongólia': 'mongolia', 'Montenegro': 'montenegro', 'MONTSERRAT': 'montserrat', 'Namíbia': 'namibia', 'Nauru': 'nauru', 'Nepal': 'nepal', 'Nicarágua': 'nicaragua', 'Níger': 'niger', 'Nigéria': 'nigeria', 'NIUE': 'niue', 'Noruega': 'noruega', 'NOVA CALEDÔNIA': 'nova_caledonia', 'Nova Zelândia': 'nova_zelandia', 'OCDE': 'ocde', 'OEA': 'oea', 'OIAçúcar': 'oiacucar', 'OICacau/ICCO': 'oicacauicco', 'OICafé': 'oicafe', 'OIE': 'oie', 'OIV': 'oiv', 'Omã': 'oma', 'OMC': 'omc', 'OMS': 'oms', 'ONU': 'onu', 'Países Baixos': 'paises_baixos', 'Palau': 'palau', 'Palestina': 'palestina', 'Panamá': 'panama', 'Papua Nova Guiné': 'papua_nova_guine', 'Paquistão': 'paquistao', 'Paraguai': 'paraguai', 'Peru': 'peru', 'PITCAIRN': 'pitcairn', 'POLINÉSIA FRANCESA': 'polinesia_francesa', 'Polônia': 'polonia', 'PORTO RICO': 'porto_rico', 'Portugal': 'portugal', 'Quênia': 'quenia', 'Quirguistão': 'quirguistao', 'Reino Unido': 'reino_unido', 'Rep. Centro-Africana': 'rep_centroafricana', 'Rep. Dem. do Congo': 'rep_dem_do_congo', 'República Dominicana': 'republica_dominicana', 'República Tcheca': 'republica_tcheca', 'RÉUNION': 'reunion', 'Romênia': 'romenia', 'Ruanda': 'ruanda', 'Rússia': 'russia', 'SAARA OCIDENTAL': 'saara_ocidental', 'SAINT BARTHÉLEMY': 'saint_barthelemy', 'SAINT HELENA, ASCENSION E TRISTAN DA CUNHA': 'saint_helena_ascension_e_tristan_da_cunha', 'Samoa': 'samoa', 'SAMOA AMERICANA': 'samoa_americana', 'San Marino': 'san_marino', 'Santa Lúcia': 'santa_lucia', 'Santa Sé (Vaticano)': 'santa_se_vaticano', 'São Cristóvão e Névis': 'sao_cristovao_e_nevis', 'SÃO MARTINHO (PARTE FRANCESA)': 'sao_martinho_parte_francesa', 'SÃO PIERRE E MIQUELON': 'sao_pierre_e_miquelon', 'São Tomé e Principe': 'sao_tome_e_principe', 'São Vicente e Granadinas': 'sao_vicente_e_granadinas', 'Seichelles (Seychelles)': 'seichelles_seychelles', 'Senegal': 'senegal', 'Serra Leoa': 'serra_leoa', 'Sérvia': 'servia', 'Singapura (Cingapura)': 'singapura_cingapura', 'SINT MAARTEN (PARTE HOLANDESA)': 'sint_maarten_parte_holandesa', 'Síria': 'siria', 'Somália': 'somalia', 'Sri Lanka': 'sri_lanka', 'Suazilândia': 'suazilandia', 'Sudão': 'sudao', 'Sudão do Sul': 'sudao_do_sul', 'Suécia': 'suecia', 'Suíça': 'suica', 'Suriname': 'suriname', 'SVALBARD E JAN MAYEN': 'svalbard_e_jan_mayen', 'Tadjiquistão': 'tadjiquistao', 'Tailândia': 'tailandia', 'Taiwan': 'taiwan', 'Tanzânia': 'tanzania', 'TERRITÓRIO OCEANO BRITÂNICO (THE)': 'territorio_oceano_britanico_the', 'TERRITÓRIOS DO SUL FRANCÊS': 'territorios_do_sul_frances', 'Timor-Leste (Timor Leste)': 'timorleste_timor_leste', 'Togo': 'togo', 'TOKELAU': 'tokelau', 'Tonga': 'tonga', 'Trinidad e Tobago': 'trinidad_e_tobago', 'Tunísia': 'tunisia', 'Turcomenistão': 'turcomenistao', 'TURKS E CAICOS ISLANDS': 'turks_e_caicos_islands', 'Turquia': 'turquia', 'Tuvalu': 'tuvalu', 'Ucrânia': 'ucrania', 'Uganda': 'uganda', 'União Europeia': 'uniao_europeia', 'Uruguai': 'uruguai', 'Uzbequistão': 'uzbequistao', 'Vanuatu': 'vanuatu', 'VÁRIOS PAÍSES': 'varios_paises', 'Venezuela': 'venezuela', 'Vietnã': 'vietna', 'WALLIS E FUTUNA': 'wallis_e_futuna', 'Zâmbia': 'zambia', 'Zimbábue': 'zimbabue'}
    # ^ Dicionário com os países que aparecem no índice dos telegramas, mas com este tendo a Key como nome do país, e o Value a pasta em que o telegrama ficará armazenado.

    TELsValidos = [] #Lista com as páginas ordenadas dos telegramas de caráter Ostensivo e Reservado, desconsiderando os Secretos.

    #Parte inicial do programa, recebe o PDF e extrai o texto das páginas.
    try:
        with open(filepath, 'rb') as arquivoPdf:
            leitor = PyPDF2.PdfReader(arquivoPdf)
            texto_pdf = ""
            for pagina in leitor.pages:
                texto_pdf += pagina.extract_text()
                texto_pdf = texto_pdf.split("\n")
                if ("Expedido" in texto_pdf[1]) or ("CARAT=Secreto" in texto_pdf[2]) or (len(texto_pdf) <= 4): #Se o telegrama for expedido, de caráter secreto ou a página tiver apenas 3 linhas (defeito na página da coleção), será ignorado.
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
        numero_pagina = (((pagina.split("\n"))[0]).split(" "))[1].split("/") 
        # ^ Extrai o número da página separando a página por quebra de linha, pegando a primeira linha, 
        # separando por espaço ("Página x/y" -> ["Página", "x/y"]), pegando o elemento de índice 1 da lista e separando por "/".
        
        #Utiliza padrões de regex para extrair as informações das páginas.
        if int(numero_pagina[0])//int(numero_pagina[1]) != 1:
            TEL.append(pagina)
        elif int(numero_pagina[0])//int(numero_pagina[1]) == 1:
            TEL.append(pagina)
            if int(numero_pagina[0]) == 1 and int(numero_pagina[1]) == 1 and len(TEL[0].split("\n")) <= 7: #Se o número da página for 1/1 e a quantidade de linhas for menor ou igual a 7, a página veio com defeito e será ignorada.
                pass
            else:
                #Extrai informações do cabeçalho.
                match_numero_tel = re.search(r"De: .*? Recebido em: (\d{2}/\d{2}/\d{4}) (\d{2}:\d{2}:\d{2}) N.°: (\d{5})", '\n'.join(TEL))
                data_recebimento = match_numero_tel.group(1) 
                hora_entrada = match_numero_tel.group(2) 
                numero_tel = int(match_numero_tel.group(3))

                #Remove o cabeçalho de cada página para extrair o conteúdo.
                for i in range(len(TEL)):
                    TEL[i] = TEL[i].splitlines()
                    TEL[i] = TEL[i][3:-1]
                    TEL[i] = "\n".join(TEL[i])
                TEL = '\n'.join(TEL)
                match_distribuicao = re.search(r'DISTR=(.*)', TEL)
                match_redistribuicao = re.search(r"Redistribuído para\s*(.*?)\s* em \d{2}/\d{2}/\d{4}", TEL) 
                if match_redistribuicao:
                    Redistribuicao = match_redistribuicao.group(1) 
                    prim_redistribuicao = Redistribuicao.split("/")[0] 
                else:
                    Redistribuicao = "NA"
                    prim_redistribuicao = "NA"

                # O telegrama será adicionado à planilha somente se a distribuição principal for da DPAgro.
                if match_distribuicao.group(1).lower().startswith("dpagro") or Redistribuicao.lower().startswith("dpagro"): 
                    Distribuicao = match_distribuicao.group(1) 

                    match_remetente_e_data = re.search(r"De (.*?) para Exteriores em (\d{2}/\d{2}/\d{4})", TEL)
                    Remetente = match_remetente_e_data.group(1)
                    data_expedicao = match_remetente_e_data.group(2)

                    match_prioridade = re.search(r'PRIOR=([\wÀ-ÿ]+)', TEL)
                    match_carater = re.search(r'CARAT=([A-Za-z]+)', TEL)
                    Indice = (re.findall('//([\s\S]*?)//', TEL))[0].replace("\n", " ").lstrip() #Procura o índice onde houver duas barras seguidas no telegrama, delimitando o início e fim.
                    Carater = match_carater.group(1) 
                    Prioridade = match_prioridade.group(1) 
                    primeira_distribuicao = "DPAGRO"

                    data_e_hora = f"{data_recebimento[6:] + '-' + data_recebimento[3:5] + '-' + data_recebimento[:2]}T{hora_entrada}Z" #Coloca a data no formato aceito pelo fluxo do Power Automate.
                    Data = datetime.strptime(data_expedicao, "%d/%m/%Y") 
                    Ano = Data.year 
                    Documento = f"TEL {numero_tel}/{Ano}/{Remetente}"
                    match_resumo = re.search(r'RESUMO=\n(.*?)(\n\s*\n)', TEL, re.DOTALL)
                    
                    Pais = "NA"
                    pasta_pais = "NA"
                    IndicePais = Indice.split(".")
                    continua = True
                    #Procura país pelo índice, o país será o primeiro que for encontrado da esquerda para a direita.
                    for i in range(len(IndicePais)):
                        for pais, nome_país in dicioDetecta.items(): 
                            if continua:
                                padrao = rf'\b{re.escape(pais)}\b'
                                resultado = re.search(padrao, IndicePais[i])
                                if resultado:
                                    Pais = nome_país
                                    pasta_pais = dicioPasta[nome_país]
                                    continua = False;
                    
                    match_resumo = re.search(r'RESUMO=\n(.*?)(\n\s*\n)', TEL, re.DOTALL) #Procura o conteúdo que estiver depois de "RESUMO=" até achar uma quebra de linha vazia.
                    if match_resumo:
                        Resumo = match_resumo.group(1).split('\n')
                        Resumo = " ".join(Resumo)
                    else:
                        Resumo = "NA"

                    if Carater == 'Reservado':
                        Teor = "RESERVADO"
                        Corpo = "RESERVADO"
                    else:
                        Teor = TEL
                        Teor = re.sub(r'(\n\s*){2,}', '\n\n', Teor)
                        match_corpo = re.search(r'(Nr\.\s\d+\s)(.*)', Teor, re.DOTALL) #Extrai todo o texto após o número do telegrama 
                        Corpo = match_corpo.group(2)
                        Corpo = re.sub(r"Retransmissão automática para .*?\r?\n\r?\n|Retransmitido via clic para .*?\r?\n\r?\n|RESUMO=.*?\r?\n\r?\n", "", Corpo, flags=re.DOTALL)
                        # ^ Remove retransmissões e resumo do corpo.
                        Corpo = re.sub(r'\n\s*\n', '\n\n', Corpo) #Substitui quebras de linha dupla por quebras de linha simples.
                        Teor = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', Teor) #Remove caracteres invisíveis, que resultariam em problema na geração do .xlsx.
                        Corpo = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', Corpo) #Remove caracteres invisíveis, que resultariam em problema na geração do .xlsx.
                        Corpo = re.sub(r'(?<=[\w.,;!?])\n(?=[^\n])', ' ', Corpo)
                    
                    match_instrucoes = re.search(r"(cumpre |cumpri |cumpro )(instrução|instruções)", TEL, re.IGNORECASE)
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
                    
                    #Adiciona todos os dados obtidos à lista Dados, que posteriormente será utilizada para gerar os arquivos .xlsx e .csv.
                    Dados.append([data_e_hora, Data.date(), "TEL", numero_tel, Ano, Remetente, Documento, Indice, Prioridade, Carater, Distribuicao, primeira_distribuicao, Redistribuicao, prim_redistribuicao, refdoc, Processos, Teor, Corpo, Resumo, Pais, pasta_pais, Instrucoes])
            TEL = [] 
    if geraExcel == 1:
        gerarExcel(os.path.splitext(os.path.basename(filepath))[0], Dados)
    if geraCsv == 1:
        gerarCsv(os.path.splitext(os.path.basename(filepath))[0], Dados)
