"""
Criação de uma GUI para o projeto
- Utilizará a biblioteca CustomTkinter para a criação de uma interface mais dinâmica
"""
import customtkinter as ctk
import os
from tkinter import filedialog, Menu
from PIL import Image,ImageTk

app = ctk.CTk()

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

#Definição referente à base do aplicativo
app.title("Leitor de Telegramas")
app.geometry("925x649")
app.resizable(False, False)

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

# Criar um frame para centralizar os botões e informações
center_frame = ctk.CTkFrame(app, width=600, height=400, corner_radius=10)
center_frame.place(relx=0.5, rely=0.5, anchor="center")  # Centraliza o frame

#Definição e configurações da logo
try:
    image = Image.open("dpagro-portal.png")
    image_resized = image.resize((300,200))
    image_tk = ImageTk.PhotoImage(image_resized)
    image_label = ctk.CTkLabel(app, image=image_tk, text="")
    image_label.pack(pady=20)
except:
    print("Imagem não encontrada")


#Definições e configurações de widgets
button = ctk.CTkButton(center_frame, text="Teste", command=button_callback)
button.grid(row=0, column=0, padx=50, pady=50)

btn_select_pdf = ctk.CTkButton(center_frame, text="Selecionar PDF", command=select_pdf)
btn_select_pdf.grid(pady=8)

checkbox_xlsx = ctk.CTkCheckBox(center_frame, text=".xlsx", command=ensure_one_selected)
checkbox_xlsx.grid(pady=5)
checkbox_xlsx.select()

checkbox_csv = ctk.CTkCheckBox(center_frame, text=".csv", command=ensure_one_selected)
checkbox_csv.grid(pady=5)
checkbox_csv.select()

#Chamada do aplicativo
app.config(menu=menu_bar)
app.mainloop()
