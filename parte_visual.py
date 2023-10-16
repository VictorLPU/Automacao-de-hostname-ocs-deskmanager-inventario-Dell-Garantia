import customtkinter as ctk
from customtkinter import filedialog
from excel_2 import dados_excel
from tkinter import messagebox

caminho_arquivo = ""

def buscar_arquivo_excel():
    global caminho_arquivo, label_arquivo

    arquivo = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo:
        caminho_arquivo = arquivo
        label_arquivo.configure(text=f"Arquivo Excel selecionado: {caminho_arquivo}")
        verificar_campos()

def verificar_campos():
    global caminho_arquivo, botao_executar

    if caminho_arquivo:
        botao_executar.configure(state=ctk.NORMAL)
    else:
        botao_executar.configure(state=ctk.DISABLED)
      
def executar_acao():
    global caminho_arquivo

    try:
        dados_excel(caminho_arquivo=caminho_arquivo)
        messagebox.showinfo("Sucesso", "A execução foi concluída com sucesso.")
    except IndexError:
        messagebox.showerror("Erro", f"Nenhum dado encontrado!")
    except Exception as e:
        messagebox.showerror("Erro", f"A execução falhou: {str(e)}")

root = ctk.CTk()
root.title("Coleta de dados")
root.geometry('300x300')
root.iconbitmap('assets/icon.ico')

botao_buscar_arquivo = ctk.CTkButton(root, text="Buscar Excel", command=buscar_arquivo_excel)
botao_buscar_arquivo.pack(pady=10)

label_arquivo = ctk.CTkLabel(root, text="")
label_arquivo.pack()

botao_executar = ctk.CTkButton(root, text="Executar", command=executar_acao, state=ctk.DISABLED)
botao_executar.pack(pady=10)

botao_buscar_arquivo.bind("<KeyRelease>", lambda event: verificar_campos())

root.mainloop()
