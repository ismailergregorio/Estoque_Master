import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from funcoes import*
from Saida_pack.funcoes_saida import*

import subprocess

contador = 0  # Contador para gerar IDs incrementais
pre_lista_saida = []  # Lista para armazenar os itens antes de serem finalizados



base_de_dados = "Base de dados\saida.xlsx"
base_de_dados_entrada = "Base de dados\entrada.xlsx"
base_de_dados_casdastro = "Base de dados\CADASTRO.xlsx"

def fechar_pagina():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "Saida_pack\Saidai.py"])


def calcular_total(event):
    """
    Calcula o valor total do produto com base na quantidade e no preço unitário.
    """
    dados = abrir_arquivo(dados_de_estoque)

    try:
        for i in dados.iter_rows(min_row=2, values_only=True):
            if int(i[0]) == int(entry_codigo.get()):
                quantidade = int(entry_quantidade.get())  # Obtém a quantidade do entry
                valor_medio = float(i[8])  # Obtém o preço unitário do entry
                total = quantidade * valor_medio  # Calcula o valor total
                valor = float(f"{total:.2f}")
                entry_valor_total.delete(0, tk.END)  # Limpa o campo antes de atualizar

        entry_valor_total.insert(0, valor)  # Insere o valor calculado
    except ValueError:
        entry_valor_total.delete(0, tk.END)  # Limpa o campo se houver erro
        entry_valor_total.insert(0, "0.00")

def verificar_campos_vazios(
        codigo,nome_produto, setor,data,quantidade,motivo_saida,funcionario,setor_funcionario,valor_total):
    campos = {
        "Codigo":codigo.get(),
        "Nome do Produto": nome_produto.get(),
        "Setor": setor.get(),
        "Data": data.get_date(),
        "Quantidade":quantidade.get(),
        "Motivo da Saida":motivo_saida.get(),
        "Funcionario":funcionario.get(),
        "Setor do Funcionario":setor_funcionario.get(),
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
        entry_codigo,entry_produto,entry_setor,entry_data,entry_quantidade,entry_tipo_sai,entry_funcionario,entry_setor_fun,entry_valor_total)
    if decisao1:
        itens = adicionar_itens_na_lista(tree,entry_codigo,entry_produto,entry_setor,entry_data,entry_quantidade,entry_tipo_sai,entry_funcionario,entry_setor_fun,entry_valor_total,entry_obs)
        pre_lista_de_itens.append(itens)
    
def deletar_iten_lista():
    global pre_lista_de_itens

    pre_lista_de_itens = deletar_pre_lista(tree,pre_lista_de_itens)
    pre_lista_de_itens = []

def finalizar():
    global pre_lista_de_itens
    for i in pre_lista_de_itens:
        print(i)
    if(finalizar_itens_saida(pre_lista_de_itens,base_de_dados)):
        for item in tree.get_children():
            tree.delete(item)
        pre_lista_de_itens = []
        fechar_pagina()

def busca_setor_funcionario(event):
    dados = abrir_arquivo(dados_de_saida)

    for x in dados.iter_rows(min_row=2,values_only=True):
        if x[6] == entry_funcionario.get():
            entry_setor_fun.delete(0, tk.END)
            entry_setor_fun.insert(0,x[7])
            break


root = tk.Tk()
root.title("Sistema de Saída de Produtos")

root.configure(bg="blue")

# Labels e entradas para o cadastro de saída de produto
tk.Label(root, text="Saída", font=("Arial", 14), bg="blue", fg="black").grid(row=0, column=0, padx=1, pady=1, sticky="w")

tk.Label(root, text="Buscar Iten", bg="blue", fg="black").grid(row=1, column=0, padx=1, pady=1)
entry_buscar = ttk.Combobox(root, width=50)
entry_buscar.grid(row=1, column=1, padx=1, pady=1)
valores_produto = carregar_dados_entry(base_de_dados_cadastro, 1)  # Carrega os produtos
entry_buscar['values'] = valores_produto  # Define os valores no Combobox
configurar_busca_combobox(entry_buscar, valores_produto)

# entry_buscar.bind("<KeyRelease>", converter_para_maiusculo)

tk.Button(root, text="Adicionar", bg="blue", fg="black",command=lambda:atualizar_campos1(base_de_dados_cadastro,entry_buscar,entry_codigo,entry_produto,entry_setor)).grid(row=1, column=2, padx=1, pady=1, sticky="ew")

tk.Label(root, text="Código", bg="blue", fg="black").grid(row=2, column=0, padx=1, pady=1)
entry_codigo = ttk.Combobox(root,state="readonly")
entry_codigo.grid(row=3, column=0, padx=1, pady=1)
entry_codigo.bind("<KeyRelease>", converter_para_maiusculo)

tk.Label(root, text="Produto", bg="blue", fg="black").grid(row=2, column=1, padx=1, pady=1)
entry_produto = ttk.Combobox(root,width=47,state="readonly")
entry_produto.grid(row=3, column=1, padx=1, pady=1)
entry_produto.bind("<KeyRelease>", converter_para_maiusculo)

tk.Label(root, text="Setor", bg="blue", fg="black",width=12).grid(row=2, column=2, padx=1, pady=1)
entry_setor = ttk.Combobox(root,state="readonly")
entry_setor.grid(row=3, column=2, padx=1, pady=1)
entry_setor.bind("<KeyRelease>", converter_para_maiusculo)

tk.Label(root, text="Data", bg="blue", fg="black").grid(row=2, column=3, padx=1, pady=1)
entry_data = DateEntry(root, background='darkblue', foreground='white', borderwidth=2, year=2024)
entry_data.grid(row=3, column=3, padx=1, pady=1)

tk.Label(root, text="Quantidade", bg="blue", fg="black").grid(row=2, column=4, padx=1, pady=1)
entry_quantidade = tk.Entry(root)
entry_quantidade.grid(row=3, column=4, padx=1, pady=1)
entry_quantidade.bind("<KeyRelease>", calcular_total)

lista_motivo_saida = ["ADMISSÃO","TROCA DE FUNÇÃO","USO","TROCA","OUTROS"]

tk.Label(root, text="Motivo", bg="blue", fg="black",width=10).grid(row=4, column=0, padx=1, pady=1)
entry_tipo_sai = ttk.Combobox(root,values=lista_motivo_saida)
entry_tipo_sai.grid(row=5, column=0, padx=1, pady=1)
entry_tipo_sai.bind("<KeyRelease>", converter_para_maiusculo)

nome_funcionario = carregar_dados_entry(base_de_dados, 6)  # Carrega os produtos
tk.Label(root, text="Funcionário", bg="blue", fg="black").grid(row=4, column=1, padx=1, pady=1)
entry_funcionario = ttk.Combobox(root,width=50,values=nome_funcionario)
entry_funcionario.grid(row=5, column=1, padx=1, pady=1)
configurar_busca_combobox(entry_funcionario, nome_funcionario)
entry_funcionario.bind("<<ComboboxSelected>>", busca_setor_funcionario)

lista_setores_funcionarios = ["APOIO AO CONSUMIDOR","COMERCIAL","CONTINUO","CULINARISTA","DESCONHECIDO",
                                "FATURAMENTO","FINANCEIRO","FISCAL","GERENTE","LIMPEZA","OPERAÇÃO",
                                "PROMOTOR","RH","STAR B","SUPERVISOR","SUPLIMENTOS","TECNICA SEGURANÇA",
                                "TELEVENDAS","TI","TRADE","TRANSPORTE","VENDEDOR"
                                ]

tk.Label(root, text="Setor funcionaria", bg="blue", fg="black").grid(row=4, column=2, padx=1, pady=1)
entry_setor_fun = ttk.Combobox(root,values=lista_setores_funcionarios)
entry_setor_fun.grid(row=5, column=2, padx=1, pady=1)
entry_setor_fun.bind("<KeyRelease>", converter_para_maiusculo)

tk.Label(root, text="Valor total", bg="blue", fg="black").grid(row=4, column=3, padx=1, pady=1)
entry_valor_total = tk.Entry(root,)
entry_valor_total.grid(row=5, column=3, padx=1, pady=1)


tk.Label(root, text="Obs", bg="blue", fg="black").grid(row=4, column=4, padx=1, pady=1)
entry_obs = tk.Entry(root)
entry_obs.grid(row=5, column=4, padx=1, pady=1)
entry_obs.bind("<KeyRelease>", converter_para_maiusculo)

# Botões de ação
tk.Button(root, text="Salvar", bg="blue", fg="black",command=adicionar_lista).grid(row=3, column=5,columnspan=2, padx=1, pady=1, sticky="ew")
tk.Button(root, text="Deletar", bg="blue", fg="black",command=deletar_iten_lista).grid(row=5, column=5,columnspan=2, padx=1, pady=1, sticky="ew")

colunas = ("Nº","Codigo","nome_produto", "setor", "data_de_entrada", "quantidade","motivo", "valor","valor_total","observação")
tree = ttk.Treeview(root, columns=colunas, show="headings")

tree = ttk.Treeview(root, columns=("Nº","Codigo","Nome", "Setor", "Data saida", "Quantidade","Motivo", "Funcionario","Setor Funcionario","Valor Total","Observação"), show='headings')

tree.heading("Nº", text="Nº")
tree.heading("Codigo", text="Codigo")
tree.heading("Nome", text="Nome do Produto",)
tree.heading("Setor", text="Setor")
tree.heading("Data saida", text="Data saida")
tree.heading("Quantidade", text="Quantidade")
tree.heading("Motivo", text="Motivo")
tree.heading("Funcionario", text="Funcionario")
tree.heading("Setor Funcionario", text="Setor Funcionario")
tree.heading("Valor Total", text="Valor Total")
tree.heading("Observação", text="Observação")

tree.column("Nº", width=30)
tree.column("Codigo", width=30)
tree.column("Nome", width=100, stretch=True)
tree.column("Setor", width=100, anchor="center", stretch=True)
tree.column("Data saida", width=100, anchor="center", stretch=True)
tree.column("Quantidade", width=80, anchor="center", stretch=True)
tree.column("Motivo", width=100, anchor="center", stretch=True)
tree.column("Funcionario", width=100, anchor="center", stretch=True)
tree.column("Setor Funcionario", width=100, anchor="center", stretch=True)
tree.column("Valor Total", width=100, anchor="center", stretch=True)
tree.column("Observação", width=100, anchor="center", stretch=True)

# Configurando o Treeview para preencher o espaço restante e ser responsivo
tree.grid(row=7, column=0, columnspan=17, rowspan=1, padx=1, pady=1, sticky="nsew")

# Botões de cancelamento e finalização
tk.Button(root, text="cancelar", bg="blue", fg="black",command=lambda:fechar_pagina()).grid(row=20, column=7, padx=1, pady=1, sticky="ew")
tk.Button(root, text="finalizar", bg="blue", fg="black",command=finalizar).grid(row=20, column=8, padx=1, pady=1, sticky="ew")

if __name__ == "__main__":
    root.mainloop()
