import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from funcoes import*
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import subprocess
from funcoes_NF import*
from datetime import date, timedelta

base_de_dados = "Base de dados\Cadastro NF.xlsx"

data_atual = date.today()

# Subtrair 360 dias
data_menos_360_dias = data_atual - timedelta(days=360)

def atualizar_filtros(event):
    filtro(tree,base_de_dados,entry_data_inicial.get_date(),entry_data_final.get_date(),entry_buscar_produto.get(),3,1)

def deletar_nota(tree_descrição):
    deletar_item_entrada(tree_descrição,"Base de dados\entrada.xlsx")


# Função de inicialização da interface

def abrir_main():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "main.py"])
def abrir_janel_entrada():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "NF\Janela_entrada_nf.py"])

root = tk.Tk()

root.title("Sistema de Cadastro de Produtos")

# # Definindo o layout
# root.geometry("1000x500")
# root.configure(bg="white")

# Labels e entradas para o cadastro de produto
tk.Label(root, text="Cadastro de NF", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, columnspan=2)
tk.Button(root, text="Sair", bg="blue", fg="white",
            command=abrir_main).grid(row=0, column=6, padx=5, pady=5, sticky="ew")

tk.Button(root, text="Adicionar NF", bg="blue", fg="white",command=abrir_janel_entrada).grid(row=1, column=0, padx=5, pady=5, sticky="ew")


tk.Button(root, text="Deletar NF", bg="blue", fg="white",command=lambda:deletar_nota(tree)).grid(row=1, column=1, padx=5, pady=5, sticky="ew")


tk.Button(root, text="Abri excel", bg="blue", fg="white").grid(row=1, column=2, padx=5, pady=5, sticky="ew")

tk.Button(root, text="Atulizar dados", bg="blue", fg="white",command=lambda:carregar_dados(base_de_dados,tree)).grid(row=1, column=12, padx=5, pady=5, sticky="ew")

tk.Label(root, text="Buscar produto", bg="blue", fg="white").grid(row=5, column=0, padx=5, pady=5, sticky="ew")
entry_buscar_produto = tk.Entry(root)
entry_buscar_produto.grid(row=6, column=0, padx=5, pady=5)
entry_buscar_produto.bind("<KeyRelease>" )

# Filtros de data
tk.Label(root, text="Data inicial", bg="blue", fg="white").grid(row=7, column=0, padx=5, pady=5)
entry_data_inicial = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
entry_data_inicial.set_date(data_menos_360_dias)
entry_data_inicial.grid(row=8, column=0, padx=5, pady=5)
entry_data_inicial.bind("<<DateEntrySelected>>", atualizar_filtros)

tk.Label(root, text="Data final", bg="blue", fg="white").grid(row=9, column=0, padx=5, pady=5)
entry_data_final = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2,)

entry_data_final.grid(row=10, column=0, padx=5, pady=5)
entry_data_final.bind("<<DateEntrySelected>>", atualizar_filtros)

tk.Button(root, text="Filtrar", bg="blue", fg="white",command=lambda:filtro(tree,entry_data_inicial,entry_data_final,base_de_dados)).grid(row=11, column=0, padx=5, pady=5, sticky="ew")
tk.Button(root, text="Deletar filtro", bg="blue", fg="white",command=lambda:carregar_dados(base_de_dados,tree)).grid(row=12, column=0, padx=5, pady=5, sticky="ew")

# Tabela (Treeview) para exibição dos produtos cadastrados
tree = ttk.Treeview(root, columns=["CNPJ","Nome do Fornecedor", "Seguimento","Data de Entrada","N° NF","VALOR NF","Motivo","Obs"], show="headings")

tree.heading("CNPJ", text="CNPJ")
tree.heading("Nome do Fornecedor", text="Nome do Fornecedor")
tree.heading("Seguimento", text="Seguimento")
tree.heading("Data de Entrada", text="Data de Entrada")
tree.heading("N° NF", text="N° NF")
tree.heading("VALOR NF", text="VALOR NF")
tree.heading("Motivo", text="Motivo")
tree.heading("Obs", text="Obs")

tree.column("CNPJ", width=100)
tree.column("Nome do Fornecedor", width=100)
tree.column("Seguimento", width=100)
tree.column("Data de Entrada", width=100)
tree.column("N° NF", width=100)
tree.column("VALOR NF", width=100)
tree.column("Motivo", width=100)
tree.column("Obs", width=100)

tree.grid(row=3, column=1, columnspan=12, rowspan=21, padx=5, pady=5, sticky="nsew")

tree.bind("<Double-1>", lambda event: mostrar_detalhes_nf_gerenciamento(event,tree))

# lista_completa = abrir_arquivo("Base de dados\Cadastro NF.xlsx")
workbook = openpyxl.load_workbook("Base de dados\Cadastro NF.xlsx")
sheet = workbook.active

for item in sheet.iter_rows(min_row=2,values_only=True):
    print(item)
    # print(item)
    tree.insert("", "end", values=item)

if __name__ == "__main__":
    root.mainloop()
