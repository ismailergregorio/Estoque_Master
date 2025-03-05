import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from funcoes import*
# from Saida_pack.Janela_saida import*
import tkinter as tk
from datetime import date, timedelta
from tkcalendar import DateEntry
import subprocess

from funcoes_saida import*

data_atual = date.today()

# Subtrair 360 dias
data_menos_360_dias = data_atual - timedelta(days=360)

base_de_dados = "Base de dados\saida.xlsx"

def atualizar_filtros(event):
    filtro(tabela_frame,base_de_dados,entry_data_inicial.get_date(),entry_data_final.get_date(),entry_buscar_produto.get(),3,1)

def converter_para_maiusculo_1(event):
    # Obtém o texto atual do Entry
    texto = entry_buscar_produto.get()
    # Converte o texto para maiúsculas
    entry_buscar_produto.delete(0, tk.END)  # Limpa o conteúdo atual
    entry_buscar_produto.insert(0, texto.upper())  # Insere o texto em maiúsculas

# Função de inicialização da interface
def abrir_adcionar_produto():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "Saida_pack\Janela_saida.py"])  # Abre a Página 2

def abrir_main():
    root.quit()  # Fecha a janela atual

def abri_admissao():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "Admissao_pack\Admissao.py"])  # Abre a Página 2

def deletar():
    deletar_item_saida(tabela_frame,base_de_dados)
    carregar_dados(base_de_dados,tabela_frame)

root = tk.Tk()
root.title("Sistema de Cadastro de Produtos")

# # Definindo o layout
# root.geometry("1000x500")
# root.configure(bg="white")

# Labels e entradas para o cadastro de produto
tk.Label(root, text="Saida", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, columnspan=2)

tk.Button(root, text="Adicionar Produto", bg="blue", fg="white",command=abrir_adcionar_produto).grid(row=1, column=0, padx=5, pady=5, sticky="ew")


tk.Button(root, text="Deletar Produto", bg="blue", fg="white",command=lambda:deletar()).grid(row=1, column=1, padx=5, pady=5, sticky="ew")

tk.Button(root, text="Abri excel", bg="blue", fg="white").grid(row=1, column=2, padx=5, pady=5, sticky="ew")

tk.Button(root, text="Saida de Funcionaio Novos", bg="blue", fg="white",command=abri_admissao).grid(row=1, column=11, padx=5, pady=5, sticky="ew")

def atualizar_e_converter(event):
    converter_para_maiusculo_1(event)
    atualizar_filtros(event)

tk.Label(root, text="Buscar produto", bg="blue", fg="white").grid(row=5, column=0, padx=5, pady=5, sticky="ew")
entry_buscar_produto = tk.Entry(root)
entry_buscar_produto.grid(row=6, column=0, padx=5, pady=5)
entry_buscar_produto.bind("<KeyRelease>", atualizar_e_converter)

# Filtros de data
tk.Label(root, text="Data inicial", bg="blue", fg="white").grid(row=7, column=0, padx=5, pady=5)
entry_data_inicial = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2,)
entry_data_inicial.grid(row=8, column=0, padx=5, pady=5)
entry_data_inicial.set_date(data_menos_360_dias)
entry_data_inicial.bind("<<DateEntrySelected>>", atualizar_filtros)

tk.Label(root, text="Data final", bg="blue", fg="white").grid(row=9, column=0, padx=5, pady=5)
entry_data_final = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2,)
entry_data_final.grid(row=10, column=0, padx=5, pady=5)
entry_data_final.bind("<<DateEntrySelected>>", atualizar_filtros)

tabela_frame = tabela(base_de_dados,root,0,)
tabela_frame.grid(row=3, column=1, columnspan=11, rowspan=20, padx=5, pady=5, sticky="nsew")

carregar_dados(base_de_dados,tabela_frame)

if __name__ == "__main__":
    root.mainloop()
