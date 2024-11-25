"""
Criação de uma GUI para o projeto
- Utilizará a biblioteca CustomTkinter para a criação de uma interface mais dinâmica
"""
import customtkinter

app = customtkinter.CTk()

def button_callback():
    print("Testando")

app.title("Leitor de Telegramas")
app.geometry("400x150")

button = customtkinter.CTkButton(app, text="Teste", command=button_callback)
button.grid(row=0, column=0, padx=20, pady=20)

app.mainloop()
