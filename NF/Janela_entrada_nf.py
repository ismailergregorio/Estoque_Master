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

    
base_de_dados_cadastro_fornecedor = 'Base de dados\Cadastro Fornecedor.xlsx'
base_de_dados_entrada = 'Base de dados\entrada.xlsx'
base_de_dados_entrada_itens = 'Base de dados\CADASTRO.xlsx'
base_de_dados_nf = "Base de dados\Cadastro NF.xlsx"

pre_lista =[]
cotagem = 1
# base_de_dados = 'Base de dados\entrada.xlsx'
def fechar_janela(funcao=None):
    if funcao == None:
        root.quit()  # Fecha a janela atual
        subprocess.Popen(["python", "NF\Cadastro_NF.py"])  # Abre a Página 2
    else:
        root.quit()  # Fecha a janela atual

def carregar_fornecedor():
    dados = abrir_arquivo(base_de_dados_cadastro_fornecedor)

    lista_alfaetica = []
    for i in dados.iter_rows(min_row=2,values_only=True):
        lista_alfaetica.append(i)
    
    lista_alfaetica.sort(key=lambda x: x[1])

    lista = []
    for x in lista_alfaetica:
        i = str(x[1])
        lista.append(i)
    
    return lista

def adicionar_fornecedor():
    fornecedor_selecionado = entry_busca_fornecedor.get()

    dados = abrir_arquivo(base_de_dados_cadastro_fornecedor)

    for x in dados.iter_rows(min_row=2,values_only=True):
        i = str(x[1])
        if i == fornecedor_selecionado:
            entry_cnpj.set(x[0])
            entry_nome_fornecedor.set(x[1])
            entry_seguimento.set(x[2])

            entry_busca_fornecedor.delete(0, tk.END)

def calcula_preco_total(event):
    try:
        quantidade = int(entry_quantidade.get())
        valor_unitario = float(entry_valor_unitario.get())
        valor_total = quantidade * valor_unitario
        entry_valor_total.delete(tk.END)
        entry_valor_total.set(f"{valor_total:.2f}")  # Correção na formatação
    except ValueError:  # Exceção específica para evitar capturar outros erros inesperados
        entry_valor_total.delete(0, tk.END)
        entry_valor_total.set("0.00")

def verifica_campo_vasio():
    campos = {"CNPJ": entry_cnpj.get(),
            "NOME FORNECEDOR": entry_nome_fornecedor.get(),
            "SEGUINMENTO": entry_seguimento.get(),
            "DATA DA ENTRADA": entry_date.get_date(),
            "Nº NF": numero_nf.get(),
            "TIPO DE ENTRADA": entry_tipo_entrada.get(),
            "QUANTIDADE": entry_quantidade.get(),
            "PRODUTO": entry_produto.get(),
            "VALOR UNITARIO": entry_valor_unitario.get(),
            "VALOR TOTAL": entry_valor_total.get(),
            "CODIGO": entry_codigo.get(),
            "SETOR": entry_setor.get(),
    }

    campos_vazios = [campo for campo, valor in campos.items() if not str(valor).strip()]

    if campos_vazios:
        mensagem = "Os seguintes campos estão vazios:\n" + "\n".join(campos_vazios)
        messagebox.showwarning("Erro", f"Iten não preenchidos {mensagem}")
        return False
    return True
valor_nota = 0

def validacao_entrada_numero(texto):
    return texto.isdigit() or texto == ""



def ajusta_valor_total_soma():
    global valor_nota

    valor_nota = float(entry_valor_total.get()) + valor_nota
    texto_var.set(f"R$ {valor_nota}")

def ajusta_valor_total_subtracao():
    global valor_nota

    item_selecionado = tree.selection()
    valores = tree.item(item_selecionado, "values")

    valor_nota = float(valor_nota) - float(valores[8]) 

    texto_var.set(f"R$ {valor_nota}")


def pre_salvamento():
    global pre_lista 
    global cotagem

    feito = False

    maior_valor = 0
    for x in pre_lista:
        if int(x[0]) > maior_valor:
            maior_valor = int(x[0])

    cotagem = maior_valor + 1
    if not verifica_nota(event=0):
        if verifica_campo_vasio():
            if entry_valor_total.get() != "0.00":
                tree.insert("", "end",values=(cotagem,entry_codigo.get(),entry_produto.get(),entry_setor.get(),entry_date.get_date(),entry_quantidade.get(),entry_tipo_entrada.get(),entry_valor_unitario.get(),entry_valor_total.get(),f"NF° {numero_nf.get()},For:{entry_nome_fornecedor.get()}"))
                itens = [cotagem,entry_codigo.get(),entry_produto.get(),entry_setor.get(),entry_date.get_date(),entry_quantidade.get(),entry_tipo_entrada.get(),entry_valor_unitario.get(),entry_valor_total.get(),f"NF° {numero_nf.get()},For:{entry_nome_fornecedor.get()}"]
                pre_lista.append(itens)
                ajusta_valor_total_soma()
                feito = True
            else:
                print("ERRO AO SALVAR")
                feito = False
        else:
            print("ERRO AO SALVAR")
            feito = False
    else:
        mostrar_alerta("NF ja existe verifique a DANF")

    if feito:
        entry_codigo.delete(0, tk.END)
        entry_produto.delete(0, tk.END)
        entry_setor.delete(0, tk.END)
        entry_quantidade.delete(0, tk.END)
        entry_valor_unitario.delete(0, tk.END)
        entry_valor_total.delete(0, tk.END)


def deletar_iten_lista():
    global pre_lista

    if tree.selection():
        ajusta_valor_total_subtracao()

        pre_lista = deletar_pre_lista(tree,pre_lista)

        for_lista_print(pre_lista)

def busca_codigo(event):
    dados = abrir_arquivo(base_de_dados_entrada_itens)

    for x in dados.iter_rows(min_row=2,values_only=True):
        if x[1] == entry_produto.get():
            entry_codigo.delete(0, tk.END)
            entry_setor.delete(0, tk.END)
            entry_codigo.insert(0,x[0])
            entry_setor.insert(0,x[2])

def busca_produto(event):
    dados = abrir_arquivo(base_de_dados_entrada_itens)

    for x in dados.iter_rows(min_row=2,values_only=True):
        if str(x[0]) == entry_codigo.get():
            entry_produto.delete(0, tk.END)
            entry_setor.delete(0, tk.END)
            entry_produto.insert(0,x[1])
            entry_setor.insert(0,x[2])

def verifica_nota(event):
    numero_digitado = numero_nf.get()
    cnpj_digitado = entry_nome_fornecedor.get()
    nota_existe = False

    if numero_digitado.isdigit():  # Verifica se o valor é numérico
        numero_digitado = int(numero_digitado)
        dados_nf = abrir_arquivo(dados_de_cadastro_nf)

        nota_existe = False

        if dados_nf:
            for numeroNF in dados_nf.iter_rows(min_row=2, values_only=True):
                try:
                    if int(numeroNF[3]) == numero_digitado:
                        if numeroNF[0] == cnpj_digitado:
                            print(f"Achei: {numeroNF}")
                            nota_existe = True
                            break
                except (ValueError, TypeError):
                    continue

        # Altera a cor da fonte conforme a existência da nota
        if nota_existe:
            numero_nf.config(fg="red")
        else:
            numero_nf.config(fg="black")
    else:
        numero_nf.config(fg="black")  # Mantém a cor preta se for inválido
    
    return nota_existe

def adicionar_nota():
    try:
        workbook_dados = load_workbook(base_de_dados_nf)
        sheet_dados = workbook_dados.active  # Seleciona a planilha ativa

        fornecedor = [entry_cnpj.get(),entry_nome_fornecedor.get(),entry_seguimento.get(),entry_date.get_date(),numero_nf.get(),valor_nota,entry_tipo_entrada.get(),obs.get()]

        sheet_dados.append(fornecedor)
        workbook_dados.save(base_de_dados_nf)
        return True
    except:
        print("Erro ao Salvar")
        return False



def finalizar():
    global pre_lista
    global valor_nota

    finalizado = False
    if pre_lista and valor_nota:
        if (finalizar_itens_entrada(pre_lista,base_de_dados_entrada)):
            for item in tree.get_children():
                tree.delete(item)
            finalizado = True
        
        if finalizado:
            pre_lista = []
        
        if adicionar_nota():
            entry_cnpj.delete(0, tk.END)
            entry_nome_fornecedor.delete(0, tk.END)
            entry_seguimento.delete(0, tk.END)
            entry_date.delete(0, tk.END)
            numero_nf.delete(0, tk.END)
            texto_var.set("R$ 0000.00")
            entry_tipo_entrada.delete(0, tk.END)
            obs.delete(0, tk.END)
        
        fechar_janela()
    else:
        messagebox.showwarning("Erro", f"Nenhun item foi adicionado nesta nota")

lista_fornecedor= []

def atulizar():
    global lista_fornecedor
    valores_produto = carregar_dados_entry(base_de_dados_entrada_itens, 1)  # Carrega os produtos
    entry_produto['values'] = valores_produto  # Define os valores no Combobox

    lista_fornecedor = carregar_fornecedor()
    entry_busca_fornecedor['values'] = lista_fornecedor

    root.update()


lista_fornecedor = carregar_fornecedor()
# Criação da janela principal
root = tk.Tk()
root.title("Sistema de Saída de Produtos")
texto_var = tk.StringVar()
# Configurações gerais da interface
root.configure(bg="blue")

vcmd = (root.register(validacao_entrada_numero),"%P")

# Labels e entradas para o cadastro de produtos
tk.Label(root, text="Entrada NF", font=("Arial", 14), bg="blue", fg="black").grid(row=0, column=0, padx=1, pady=1, sticky="w")

tk.Label(root, text="Fornecedor", bg="blue", fg="black",).grid(row=1, column=0, padx=1, pady=1)   
entry_busca_fornecedor = ttk.Combobox(root,width=50, values=lista_fornecedor)
entry_busca_fornecedor.grid(row=1, column=1, padx=1, pady=1)
configurar_busca_combobox(entry_busca_fornecedor, lista_fornecedor)
# entry_busca_fornecedor.bind("<KeyRelease>", converter_para_maiusculo)

tk.Button(root, text="Adicionar", bg="blue", fg="black",command=adicionar_fornecedor).grid(row=1, column=2, padx=1, pady=1, sticky="ew")


tk.Label(root, text="Nome Fornecedor", bg="blue", fg="black",width=20).grid(row=2, column=1, padx=1, pady=1)
entry_nome_fornecedor = ttk.Combobox(root,width=50,state="readonly")
entry_nome_fornecedor.grid(row=3, column=1, padx=1, pady=1)
entry_nome_fornecedor.bind("<KeyRelease>", converter_para_maiusculo)

tk.Label(root, text="CNPJ", bg="blue", fg="black",width=20).grid(row=2, column=0, padx=1, pady=1)
entry_cnpj = ttk.Combobox(root,width=20,state="readonly")
entry_cnpj.grid(row=3, column=0, padx=1, pady=1)
entry_cnpj.bind("<KeyRelease>", converter_para_maiusculo)


# Data de entrada do produto
tk.Label(root, text="Seguimento", bg="blue", fg="black",width=20).grid(row=2, column=2, padx=1, pady=1)
entry_seguimento = ttk.Combobox(root,width=20,state="readonly")
entry_seguimento.grid(row=3, column=2, padx=1, pady=1)
entry_seguimento.bind("<KeyRelease>", converter_para_maiusculo)

# Quantidade de produtos
tk.Label(root, text="Data", bg="blue", fg="black").grid(row=2, column=3, padx=1, pady=1)
entry_date = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
entry_date.grid(row=3, column=3, padx=1, pady=1)


# Motivo da entrada (e.g., compra, devolução)
tk.Label(root, text="N°-NF", bg="blue", fg="black").grid(row=2, column=4, padx=1, pady=1)
numero_nf = tk.Entry(root,validate="key",validatecommand=vcmd)
numero_nf.grid(row=3, column=4, padx=1, pady=1)

numero_nf.bind("<KeyRelease>",verifica_nota)
numero_nf.bind("<KeyRelease>", converter_para_maiusculo)

tk.Label(root, text="Motivo", bg="blue", fg="black").grid(row=2, column=5, padx=1, pady=1)
lista_motivo_entrada = ["COMPRA", "DEVOLUÇÃO", "TROCA", "OUTROS"]
entry_tipo_entrada = ttk.Combobox(root, values=lista_motivo_entrada)
entry_tipo_entrada.grid(row=3, column=5, padx=1, pady=1)
entry_tipo_entrada.bind("<KeyRelease>", converter_para_maiusculo)

# Valor unitário do produto
tk.Label(root, text="Obs", bg="blue", fg="black").grid(row=2, column=6, padx=1, pady=1)
obs = tk.Entry(root)
obs.grid(row=3, column=6, padx=1, pady=1)
obs.bind("<KeyRelease>", converter_para_maiusculo)

texto_var.set("R$ 0.000,00")

tk.Label(root, textvariable=texto_var, bg="blue", fg="black",font=("Arial", 20)).grid(row=1, column=5,columnspan=2, padx=1, pady=1)


tk.Label(root, text="Produto", bg="blue", fg="black",width=50).grid(row=4, column=1, padx=1, pady=1)
entry_produto = ttk.Combobox(root,width=50)
entry_produto.grid(row=5, column=1, padx=1, pady=1)
valores_produto = carregar_dados_entry(base_de_dados_entrada_itens, 1)  # Carrega os produtos
entry_produto['values'] = valores_produto  # Define os valores no Combobox
configurar_busca_combobox(entry_produto, valores_produto)
entry_produto.bind("<<ComboboxSelected>>",busca_codigo)


tk.Label(root, text="Codigo", bg="blue", fg="black").grid(row=4, column=0, padx=1, pady=1)
entry_codigo = ttk.Combobox(root)
entry_codigo.grid(row=5, column=0, padx=1, pady=1)
valor_codigo = carregar_dados_entry(base_de_dados_entrada_itens, 0)  # Carrega os produtos
entry_codigo['values'] = valor_codigo  # Define os valores no Combobox
configurar_busca_combobox(entry_codigo, valores_produto)
entry_codigo.bind("<<ComboboxSelected>>",busca_produto)
entry_codigo.bind("<KeyRelease>", converter_para_maiusculo)


tk.Label(root, text="Setor", bg="blue", fg="black").grid(row=4, column=2, padx=1, pady=1)
entry_setor = ttk.Combobox(root)
entry_setor.grid(row=5, column=2, padx=1, pady=1)
entry_setor.bind("<KeyRelease>", converter_para_maiusculo)

# Quantidade de produtos
tk.Label(root, text="Quantidade", bg="blue", fg="black").grid(row=4, column=3, padx=1, pady=1)
entry_quantidade = tk.Entry(root,validate="key",validatecommand=vcmd)
entry_quantidade.grid(row=5, column=3, padx=1, pady=1)
entry_quantidade.bind("<KeyRelease>",calcula_preco_total)

entry_var = tk.StringVar()
# Valor unitário do produto
tk.Label(root, text="Valor Unitario", bg="blue", fg="black").grid(row=4, column=4, padx=1, pady=1)
entry_valor_unitario = tk.Entry(root,textvariable=entry_var, validate="key",)
entry_valor_unitario.grid(row=5, column=4, padx=1, pady=1)
entry_valor_unitario.bind("<KeyRelease>",calcula_preco_total)
# Valor total do produto
tk.Label(root, text="Valor total", bg="blue", fg="black").grid(row=4, column=5, padx=1, pady=1)
entry_valor_total = ttk.Combobox(root,state="readonly")
entry_valor_total.grid(row=5, column=5, padx=1, pady=1)

# Botões de ação (Salvar, Deletar, Cancelar, Finalizar)
tk.Button(root, text="Salvar", bg="blue", fg="black",command=pre_salvamento).grid(row=4, column=6, padx=1, pady=1, sticky="ew")
tk.Button(root, text="Deletar", bg="blue", fg="black",command=deletar_iten_lista).grid(row=5, column=6, padx=1, pady=1, sticky="ew")

def aba_criar_item():
    subprocess.Popen(["python", "Cadastro_pack\Cadastro.py"])

def aba_criar_for():
    subprocess.Popen(["python", "Fornecedor\Janela_entrada_fornecedor.py"])

tk.Button(root, text="Adicionar Produto Inexistente", bg="blue", fg="black",command=aba_criar_item).grid(row=1, column=7, padx=1, pady=1, sticky="ew")
tk.Button(root, text="Adicionar Produto Fornecedor", bg="blue", fg="black",command=aba_criar_for).grid(row=2, column=7, padx=1, pady=1, sticky="ew")
tk.Button(root, text="Atualizar Dados", bg="blue", fg="black",command=atulizar).grid(row=3, column=7, padx=1, pady=1, sticky="ew")


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
tree.grid(row=7, column=0, columnspan=17, rowspan=1, padx=1, pady=1, sticky="nsew")

# Botões de cancelamento e finalização
tk.Button(root, text="cancelar", bg="blue", fg="black",command=lambda:fechar_janela(1)).grid(row=22, column=6, padx=1, pady=1, sticky="ew")
tk.Button(root, text="finalizar", bg="blue", fg="black",command=finalizar).grid(row=22, column=7, padx=1, pady=1, sticky="ew")

# Executa a criação da interface se o script for rodado diretamente
if __name__ == "__main__":
    root.mainloop()
