"""
Criação de uma GUI para o projeto
- Utilizará a biblioteca CustomTkinter para a criação de uma interface mais dinâmica
"""
import customtkinter as ctk
import os
from tkinter import filedialog, Menu
from PIL import Image, ImageTk

app = ctk.CTk()

#Definição referente à base do aplicativo
app.title("Leitor de Telegramas")
app.geometry("925x649")
app.resizable(False, False)

# Funções para o menu
# Ajuda
def show_help():
    help_window = ctk.CTkToplevel(app)
    help_window.title("Ajuda")
    help_window.geometry("400x200")
    help_window.resizable(False, False)  # Tamanho fixo

    help_label = ctk.CTkLabel(help_window, 
                              text="Esta é a janela de ajuda.\nAqui você pode adicionar informações relevantes.", 
                              font=("Arial", 14), 
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
                              text="Esta é a janela de ajuda.\nAqui você pode adicionar informações relevantes.", 
                              font=("Arial", 14), 
                              justify="center")
    help_label.pack(pady=20)

    close_button = ctk.CTkButton(help_window, text="Fechar", command=help_window.destroy)
    close_button.pack(pady=10)

def select_pdf():
    filepath = filedialog.askopenfilename(
        initialdir=os.path.expanduser("~/Downloads"),
        title="Selecione um arquivo PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    if filepath:
        print(f"PDF selecionado {filepath}")

def button_callback():
    print("Testando")

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

# Título central
title_label = ctk.CTkLabel(app, text="Conversor de Telegramas", font=("Lato", 32, "bold"))
#title_label.pack(side="top",pady=(10,0))
title_label.place(x=284, y=97)

# Subtitulo central
selection_label = ctk.CTkLabel(app, text="Selecione o arquivo pdf para a conversão", font=("Lato", 15, "bold"))
selection_label.place(x=277, y=198)

#Definição e configurações da logo

# Configurar caminho da logo
current_dir = os.path.dirname(os.path.abspath(__file__))
logo_path = os.path.join(current_dir,"logo-portal-dpagro.png")

try:
    if not os.path.isfile(logo_path):
        raise FileNotFoundError(f"Logo não encontrada no caminho: {logo_path}")
    
    logo_image = Image.open(logo_path)
    logo_resized = logo_image.resize((242, 63))
    logo_tk = ImageTk.PhotoImage(logo_resized)
    
    logo_label = ctk.CTkLabel(app, image=logo_tk, text="")
    logo_label.place(x=35, y=25 )

except FileNotFoundError as e:
    print(e)
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