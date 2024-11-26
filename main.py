"""
Criação de uma GUI para o projeto
- Utilizará a biblioteca CustomTkinter para a criação de uma interface mais dinâmica
"""
import customtkinter
import os
from tkinter import filedialog

app = customtkinter.CTk()

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

app.title("Leitor de Telegramas")
app.geometry("925x649")

button = customtkinter.CTkButton(app, text="Teste", command=button_callback)
button.grid(row=0, column=0, padx=20, pady=20)

btn_select_pdf = customtkinter.CTkButton(app, text="Selecionar PDF", command=select_pdf)
btn_select_pdf.grid(pady=5)

checkbox_xlsx = customtkinter.CTkCheckBox(app, text=".xlsx", command=ensure_one_selected)
checkbox_xlsx.grid(pady=5)
checkbox_xlsx.select()

checkbox_csv = customtkinter.CTkCheckBox(app, text=".csv", command=ensure_one_selected)
checkbox_csv.grid(pady=5)
checkbox_csv.select()

app.mainloop()
