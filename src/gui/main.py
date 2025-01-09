"""
Criação de uma GUI para o projeto
- Utilizará a biblioteca CustomTkinter para a criação de uma interface mais dinâmica
"""
import customtkinter as ctk
import os
from tkinter import filedialog, Menu, messagebox
from PIL import Image, ImageTk
import prototipo as funcoes

app = ctk.CTk()
current_dir = os.path.dirname(os.path.abspath(__file__))
ctk.set_appearance_mode("dark")

#Definição referente à base do aplicativo
app.title("Leitor de Telegramas")
app.geometry("925x649")
app.resizable(False, False)

filename = ''
filepath = ''

# Funções para o menu
# Ajuda
def show_help():
    help_window = ctk.CTkToplevel(app)
    help_window.title("Ajuda")
    help_window.geometry("400x200")
    help_window.resizable(False, False)  # Tamanho fixo

    help_label = ctk.CTkLabel(help_window, 
                              text="Este é um programa de tratamento de telegramas recebidos \n para a inserção no Portal DPAgro.\n \n Mais informações acessar:\n https://github.com/DPAgro-MRE/conversor-de-telegramas", 
                              font=("Lato", 14), 
                              justify="center")
    help_label.pack(pady=20)

    close_button = ctk.CTkButton(help_window, text="Fechar", command=help_window.destroy)
    close_button.pack(pady=10)

# Sobre
def show_about():
    help_window = ctk.CTkToplevel(app)
    help_window.title("Sobre")
    help_window.geometry("400x200")
    help_window.resizable(False, False)  # Tamanho fixo

    help_label = ctk.CTkLabel(help_window, 
                              text="Conversor de telegramas \n Desenvolvido pela Divisão de Política Agrícola (DPAgro)\n \n Link do repositório github:\nhttps://github.com/DPAgro-MRE/conversor-de-telegramas"
                              , 
                              font=("Lato", 14), 
                              justify="center")
    help_label.pack(pady=20)

    close_button = ctk.CTkButton(help_window, text="Fechar", command=help_window.destroy)
    close_button.pack(pady=10)



def message_success():
    messagebox.showinfo("Conversor de Telegramas", "Telegrama convertido com êxito!")

def message_error():
    messagebox.showerror("Conversor de Telegramas", "Erro na conversão \nPDF inserido não é um telegrama!")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.join(os.path.abspath("."), "src", "gui")
    return os.path.join(base_path, relative_path)

def select_pdf():
    global filename
    global filepath
    filepath = filedialog.askopenfilename(
        initialdir=os.path.expanduser("~/Downloads"),
        title="Selecione um arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    filename = filepath.split("/")[-1]
    if filepath:
        print(f"PDF selecionado {filename}")
        pdf_label = ctk.CTkLabel(app, text=filename, font=("Lato", 15))
        pdf_label.place(x=324, y=322)
    return filepath

def ensure_one_selected():
    if not checkbox_xlsx and not checkbox_csv:
        checkbox_xlsx.select()

#Definição e configurações menu na barra de tarefas
menu_bar = Menu(app)

#Menu Ajuda
help_menu = Menu(menu_bar, tearoff=0)
help_menu.add_command(label="Ajuda", command=show_help)
#help_menu.add_command(label="Sobre", command=show_about)
menu_bar.add_cascade(label="Ajuda", menu=help_menu)

#Menu Sobre
about_menu = Menu(menu_bar, tearoff=0)
about_menu.add_command(label="Sobre", command=show_about)
menu_bar.add_cascade(label="Sobre", menu=about_menu)

'''
# Criar um frame para centralizar os botões e informações
center_frame = ctk.CTkFrame(app, width=600, height=400, corner_radius=10)
center_frame.place(relx=0.5, rely=0.5, anchor="center")
'''

# Criação de um container para o pdf selecionado
'''
container_frame = ctk.CTkFrame(app, width=1516, height=488, corner_radius=10)
container_frame.place(relx=0.5, rely=0.5, anchor="center",)
container_frame.propagate(True)
'''
#Header
header_path = resource_path('header-dpagro.png')

try:
    if not os.path.isfile(header_path):
        raise FileNotFoundError(f"Header não encontrada no caminho: {header_path}")

    header_image = Image.open(header_path)
    header_resized = header_image.resize((925, 76))
    header_tk = ImageTk.PhotoImage(header_resized)
    header_label = ctk.CTkLabel(app, image=header_tk, text="")
    header_label.place(x=0, y=0)

except FileNotFoundError as e:
    print(e)

#Definição e configurações da logo

# Configurar caminho da logo
logo_path = resource_path('logo-portal-dpagro.png')

try:
    if not os.path.isfile(logo_path):
        raise FileNotFoundError(f"Logo não encontrada no caminho: {logo_path}")
    
    logo_image = Image.open(logo_path)
    logo_resized = logo_image.resize((242, 63))
    logo_tk = ImageTk.PhotoImage(logo_resized)
    
    logo_label = ctk.CTkLabel(app, image=logo_tk, text="", bg_color="#16214A")
    logo_label.place(x=38, y=6 )

except FileNotFoundError as e:
    print(e)

# Título central
title_label = ctk.CTkLabel(app, text="Conversor de Telegramas", font=("Lato", 40, "bold"))
#title_label.pack(side="top",pady=(10,0))
title_label.place(x=236, y=88)

# Subtitulo central
selection_label = ctk.CTkLabel(app, text="Selecione o arquivo pdf para a conversão", font=("Lato", 20, "bold"))
selection_label.place(x=273, y=190)

# Icon do pdf
pdf_icon = resource_path("pdf-icon.png")

try:
    if not os.path.isfile(pdf_icon):
        raise FileNotFoundError(f"Icone PDF não encontrado no caminho: {pdf_icon}")
    pdf_image = Image.open(pdf_icon)
    pdf_resized = pdf_image.resize((88, 88))
    pdf_tk = ImageTk.PhotoImage(pdf_resized)

    pdf_label = ctk.CTkLabel(app, image=pdf_tk, text="")
    pdf_label.place(x=309, y=233)

except FileNotFoundError as e:
    print(e)

button = ctk.CTkButton(app, text="Buscar", width=183, height=44, font=('Lato', 16, 'bold') ,command=select_pdf)
button.grid(row=0, column=0, padx=50, pady=50)
button.place(x=418, y=255)


# Nome do usuário
user_name = os.getlogin().split(".")
formated_name = f"{user_name[0].capitalize()} {user_name[len(user_name)-1].capitalize()}"
user_label = ctk.CTkLabel(app, text=formated_name, font=("Lato", 20, "bold"), bg_color="#16214A")
user_label.place(x=688, y=15)


#Escolha de formatos
format_title = ctk.CTkLabel(app, text="Selecione o formato de arquivo desejado", font=("Lato", 20, "normal"))
#title_label.pack(side="top",pady=(10,0))
format_title.place(x=273, y=397)

#XLSX
checkbox_xlsx = ctk.CTkCheckBox(app, text=".xlsx", command=ensure_one_selected, font=("Lato", 15, "normal"))
checkbox_xlsx.grid(pady=5)
checkbox_xlsx.select()
checkbox_xlsx.place(x=347, y=447)

#CSV
checkbox_csv = ctk.CTkCheckBox(app, text=".csv",command=ensure_one_selected, font=("Lato", 15, "normal"))
checkbox_csv.grid(pady=5)
checkbox_csv.select()
checkbox_csv.place(x=486, y=447)


def safe_extraction(filepath, csv_flag, xlsx_flag):
    try:
        funcoes.Extracao(filepath, csv_flag, xlsx_flag)
        return message_success() 
    except Exception as e:
        return message_error()  
    
button = ctk.CTkButton(app, text="Converter", width=183, height=44, font=('Lato', 24, "bold"), command=lambda: safe_extraction(filepath, checkbox_csv.get(), checkbox_xlsx.get()))
button.grid(row=0, column=0, padx=50, pady=50)
button.place(x=367, y=515)

icon = ImageTk.PhotoImage(file=resource_path("icone.ico"))

'''
#Definições e configurações de widgets
button = ctk.CTkButton(center_frame, text="Teste", width=183, height=44, font=('Lato', 80) ,command=button_callback)
button.grid(row=0, column=0, padx=50, pady=50)

btn_select_pdf = ctk.CTkButton(center_frame, text="Selecionar PDF", command=select_pdf)
btn_select_pdf.grid(pady=8)

checkbox_xlsx = ctk.CTkCheckBox(center_frame, text=".xlsx", command=ensure_one_selected)
checkbox_xlsx.grid(pady=5)
checkbox_xlsx.select()

checkbox_csv = ctk.CTkCheckBox(center_frame, text=".csv",command=ensure_one_selected)
checkbox_csv.grid(pady=5)
checkbox_csv.select()
'''

#Chamada do aplicativo
app.config(menu=menu_bar)
app.mainloop()

print(os.getcwd())