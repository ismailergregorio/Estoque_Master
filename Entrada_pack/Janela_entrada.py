import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from funcoes import*
from Entrada_pack.funcoes_entrada import*

import subprocess

# Variáveis globais
contador = 0  # Contador para gerar IDs incrementais
pre_lista_entrada = []  # Lista para armazenar os itens antes de serem finalizados


base_de_dados_cadastro = 'Base de dados\CADASTRO.xlsx'
base_de_dados = 'Base de dados\entrada.xlsx'

def calcular_total(event):
    """
    Calcula o valor total do produto com base na quantidade e no preço unitário.
    """
    try:
        quantidade = int(entry_quantidade.get())  # Obtém a quantidade do entry
        preco_unitario = float(entry_valor_unitario.get())  # Obtém o preço unitário do entry
        total = quantidade * preco_unitario  # Calcula o valor total
        valor = float(f"{total:.2f}")
        entry_valor_total.delete(0, tk.END)  # Limpa o campo antes de atualizar

        entry_valor_total.insert(0, valor)  # Insere o valor calculado
    except ValueError:
        entry_valor_total.delete(0, tk.END)  # Limpa o campo se houver erro
        entry_valor_total.insert(0, "0.00")


def verificar_campos_vazios(
        codigo,nome_produto, setor,data,quantidade,tipo_entrada,valor_unitario,valor_total):
    campos = {
        "Codigo":codigo.get(),
        "Nome do Produto": nome_produto.get(),
        "Setor": setor.get(),
        "Data": data.get_date(),
        "Quantidade":quantidade.get(),
        "Tipo de Entrada":tipo_entrada.get(),
        "Valor Unidade":valor_unitario.get(),
        "Valor Total":valor_total.get()

    }

    # Verifique os campos vazios
    campos_vazios = [campo for campo, valor in campos.items() if not str(valor).strip()]

    if campos_vazios:
        mensagem = "Os seguintes campos estão vazios:\n" + "\n".join(campos_vazios)
        print(mensagem)  # Ou exiba uma mensagem usando messagebox
        return False
    return True

pre_lista_de_itens = []


def adicionar_lista():
    global pre_lista_de_itens
    decisao1 = verificar_campos_vazios(
        entry_codigo,entry_produto,entry_setor,entry_data,entry_quantidade,entry_tipo_entrada,entry_valor_unitario,entry_valor_total)
    if decisao1:
        itens = adicionar_itens_na_lista(tree,entry_codigo,entry_produto,entry_setor,entry_data,entry_quantidade,entry_tipo_entrada,entry_valor_unitario,entry_valor_total,entry_obs)
        pre_lista_de_itens.append(itens)

    # validar_numeros_decimais(entry_quantidade.get())
    # salvar_itens_entrada(base_de_dados,entry_codigo.get(),entry_produto.get(),entry_setor.get(),entry_data.get_date(),entry_quantidade.get(),entry_tipo_entrada.get(),entry_valor_unitario.get(),entry_valor_total.get(),entry_obs.get())

def deletar_iten_lista():
    global pre_lista_de_itens

    pre_lista_de_itens = deletar_pre_lista(tree,pre_lista_de_itens)

    for_lista_print(pre_lista_de_itens)

def finalizar():
    if (finalizar_itens_entrada(pre_lista_de_itens,base_de_dados)):
        for item in tree.get_children():
            tree.delete(item)


def cancela_ferchar():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "Entrada_pack\Entradai.py"])

# Criação da janela principal
root = tk.Tk()
root.title("Sistema de Saída de Produtos")

# Configurações gerais da interface
root.configure(bg="blue")

# Labels e entradas para o cadastro de produtos
tk.Label(root, text="Entrada", font=("Arial", 14), bg="blue", fg="black").grid(row=0, column=0, padx=1, pady=1, sticky="w")

tk.Label(root, text="Buscar Iten", bg="blue", fg="black").grid(row=1, column=0, padx=1, pady=1)
entry_buscar = ttk.Combobox(root, width=50)
entry_buscar.grid(row=1, column=1, padx=1, pady=1)
valores_produto = carregar_dados_entry(base_de_dados_cadastro, 1)  # Carrega os produtos
entry_buscar['values'] = valores_produto  # Define os valores no Combobox
configurar_busca_combobox(entry_buscar, valores_produto)

tk.Button(root, text="Adicionar", bg="blue", fg="black",command=lambda:atualizar_campos1(base_de_dados_cadastro,entry_buscar,entry_codigo,entry_produto,entry_setor)).grid(row=1, column=2, padx=1, pady=1, sticky="ew")


tk.Label(root, text="Código", bg="blue", fg="black", width=10).grid(row=2, column=0, padx=1, pady=1)
entry_codigo = ttk.Combobox(root, width=10,state="readonly")
entry_codigo.grid(row=3, column=0, padx=1, pady=1)

# Produto
tk.Label(root, text="Produto", bg="blue", fg="black").grid(row=2, column=1, padx=1, pady=1)
entry_produto = ttk.Combobox(root, width=50,state="readonly")
entry_produto.grid(row=3, column=1, padx=1, pady=1)


tk.Label(root, text="Setor", bg="blue", fg="black",width=12).grid(row=2, column=2, padx=1, pady=1)
entry_setor = ttk.Combobox(root,width=12,state="readonly")
entry_setor.grid(row=3, column=2, padx=1, pady=1)

# Data de entrada do produto
tk.Label(root, text="Data", bg="blue", fg="black").grid(row=2, column=3, padx=1, pady=1)
entry_data = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2,)
entry_data.grid(row=3, column=3, padx=1, pady=1)

# Quantidade de produtos
tk.Label(root, text="Quantidade", bg="blue", fg="black").grid(row=2, column=4, padx=1, pady=1)
entry_quantidade = tk.Entry(root)
entry_quantidade.grid(row=3, column=4, padx=1, pady=1)
entry_quantidade.bind("<KeyRelease>", calcular_total)


# Motivo da entrada (e.g., compra, devolução)
tk.Label(root, text="Motivo", bg="blue", fg="black").grid(row=2, column=5, padx=1, pady=1)
lista_motivo_entrada = ["COMPRA", "DEVOLUÇÃO", "TROCA", "OUTROS"]
entry_tipo_entrada = ttk.Combobox(root, values=lista_motivo_entrada)
entry_tipo_entrada.grid(row=3, column=5, padx=1, pady=1)

# Valor unitário do produto
tk.Label(root, text="Valor Unitario", bg="blue", fg="black").grid(row=2, column=6, padx=1, pady=1)
entry_valor_unitario = tk.Entry(root)
entry_valor_unitario.grid(row=3, column=6, padx=1, pady=1)
entry_valor_unitario.bind("<KeyRelease>", calcular_total)


# Valor total do produto
tk.Label(root, text="Valor total", bg="blue", fg="black").grid(row=2, column=7, padx=1, pady=1)
entry_valor_total = tk.Entry(root)
entry_valor_total.grid(row=3, column=7, padx=1, pady=1)

# Campo de observações
tk.Label(root, text="Obs", bg="blue", fg="black").grid(row=2, column=8, padx=1, pady=1)
entry_obs = tk.Entry(root)
entry_obs.grid(row=3, column=8, padx=1, pady=1)

# Botões de ação (Salvar, Deletar, Cancelar, Finalizar)
tk.Button(root, text="Salvar", bg="blue", fg="black",command=adicionar_lista).grid(row=4, column=7, padx=1, pady=1, sticky="ew")
tk.Button(root, text="Deletar", bg="blue", fg="black",command=deletar_iten_lista).grid(row=4, column=8, padx=1, pady=1, sticky="ew")

# Tabela (Treeview) para exibição dos produtos cadastrados
colunas = ("Nº","Codigo", "Produto", "Setor", "Data de Entrada", "Quantidade", "Motivo", "Valor Unitario", "Valor Total", "Obs")
tree = ttk.Treeview(root, columns=colunas, show="headings")


# Definição das colunas e seus respectivos cabeçalhos
tree.heading("Nº", text="Nº")
tree.heading("Codigo", text="Codigo")
tree.heading("Produto", text="Nome do produto")
tree.heading("Setor", text="Setor")
tree.heading("Data de Entrada", text="Data Entrada")
tree.heading("Quantidade", text="Quantidade")
tree.heading("Motivo", text="Motivo")
tree.heading("Valor Unitario", text="Valor Unitario")
tree.heading("Valor Total", text="Valor Total")
tree.heading("Obs", text="Obs")

# Definição da largura de cada coluna
tree.column("Nº", width=30)
tree.column("Codigo", width=30)
tree.column("Produto", width=100)
tree.column("Setor", width=100)
tree.column("Data de Entrada", width=100)
tree.column("Quantidade", width=100)
tree.column("Motivo", width=100)
tree.column("Valor Unitario", width=100)
tree.column("Valor Total", width=100)
tree.column("Obs", width=100)

# Configurando o Treeview para preencher o espaço restante e ser responsivo
tree.grid(row=5, column=0, columnspan=17, rowspan=1, padx=1, pady=1, sticky="nsew")

tk.Button(root, text="Finalizar", bg="blue", fg="black",command=finalizar).grid(row=6, column=7, padx=1, pady=1, sticky="ew")
tk.Button(root, text="Cancelar", bg="blue", fg="black",command=cancela_ferchar).grid(row=6, column=8, padx=1, pady=1, sticky="ew")

# Executa a criação da interface se o script for rodado diretamente
if __name__ == "__main__":
    root.mainloop()
