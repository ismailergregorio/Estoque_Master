import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))


from funcoes import*
import tkinter as tk
from tkinter import ttk
import subprocess

# Função de inicialização da interface


def abrir_main():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "main.py"])

base_de_dados = "Base de dados\Estoque.xlsx"

root = tk.Tk()
root.title("Sistema de Cadastro de Produtos")



# # Definindo o layout
# root.geometry("1000x500")
# root.configure(bg="white")

# Labels e entradas para o cadastro de produto
tk.Label(root, text="Estoque", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, columnspan=2)
tk.Button(root, text="Sair", bg="blue", fg="white",
            command=abrir_main).grid(row=0, column=6, padx=5, pady=5, sticky="ew")

tk.Button(root, text="Adicionar Entrada", bg="blue", fg="white").grid(row=1, column=0, padx=5, pady=5, sticky="ew")


tk.Button(root, text="Adicionar Saida", bg="blue", fg="white").grid(row=1, column=1, padx=5, pady=5, sticky="ew")


tk.Button(root, text="Abri excel", bg="blue", fg="white").grid(row=1, column=2, padx=5, pady=5, sticky="ew")

tk.Label(root, text="Buscar produto", bg="blue", fg="white").grid(row=5, column=0, padx=5, pady=5, sticky="ew")
entry_buscar_produto = tk.Entry(root)
entry_buscar_produto.grid(row=6, column=0, padx=5, pady=5)

# Tabela (Treeview) para exibição dos produtos cadastradoS

tree = ttk.Treeview(root, columns=("Codigo","Nome", "Setor","Estoque Minimo", "Entrada", "Saida","Saldo", "Valor Total","Valor Minimo"), show='headings')
tree.heading("Codigo", text="Codigo",)
tree.heading("Nome", text="Nome do Produto",)
tree.heading("Setor", text="Setor")
tree.heading("Estoque Minimo", text="Estoque Minimo")
tree.heading("Saida", text="Saida")
tree.heading("Entrada", text="Entrada")
tree.heading("Saldo", text="Saldo")
tree.heading("Valor Total", text="Valor Total")
tree.heading("Valor Minimo", text="Valor Minimo")

tree.column("Codigo", width=50, stretch=True)
tree.column("Nome", width=200, stretch=True)
tree.column("Setor", width=200, anchor="center", stretch=True)
tree.column("Estoque Minimo", width=100, anchor="center", stretch=True)
tree.column("Entrada", width=100, anchor="center", stretch=True)
tree.column("Saida", width=100, anchor="center", stretch=True)
tree.column("Saldo", width=100, anchor="center", stretch=True)
tree.column("Valor Total", width=100, anchor="center", stretch=True)
tree.column("Valor Minimo", width=100, anchor="center", stretch=True)

tree.grid(row=3, column=1, columnspan=11, rowspan=20, padx=5, pady=5, sticky="nsew")

carregar_dados(base_de_dados,tree)

root.mainloop()