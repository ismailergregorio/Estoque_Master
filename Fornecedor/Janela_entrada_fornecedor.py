import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import tkinter as tk
from tkinter import ttk
from funcoes import*
import subprocess


# Variáveis globais
contador = 0  # Contador para gerar IDs incrementais

lista_seguimento = ["seguimento 1", "seguimento 2"]

pre_lista_entrada = []  # Lista para armazenar os itens antes de serem finalizados


base_de_dados = 'Base de dados\Cadastro Fornecedor.xlsx'

def fechar_janela():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "Fornecedor\Cadastro Fornecedor.py"])  # Abre a Página 2

atraso_formatacao = None
def verificar_campos_vazios(
        cnpj,nome_fonecedor, seguimento,numero,email):
    campos = {
        "CNPJ":cnpj.get(),
        "NOME FORNECEDOR": nome_fonecedor.get(),
        "SEGUIMENTO": seguimento.get(),
        "NUMERO": numero.get(),
        "E-MAIL":email.get(),
    }

    # Verifique os campos vazios
    campos_vazios = [campo for campo, valor in campos.items() if not str(valor).strip()]

    if len(cnpj.get()) < 14:
        mensagem = "CNPJ não esta correto"
        messagebox.showwarning("Campo Vazio", f"Verifique os capos de preenchimento {mensagem}.")
        return False

    if len(numero.get()) < 14:
        mensagem = "Numero não esta correto"
        print(mensagem)  # Ou exiba uma mensagem usando messagebox
        messagebox.showwarning("Campo Vazio", f"Verifique os capos de preenchimento {mensagem}.")
        return False
    
    if campos_vazios:
        mensagem = "Os seguintes campos estão vazios:\n" + "\n".join(campos_vazios)
        print(mensagem)  # Ou exiba uma mensagem usando messagebox
        messagebox.showwarning("Campo Vazio", f"Verifique os capos de preenchimento {mensagem}.")
        return False
    
    return True

def converter_para_maiusculo_1(event):
    # Obtém o texto atual do Entry
    texto = entry_nome_fornecedor.get()
    # Converte o texto para maiúsculas
    entry_nome_fornecedor.delete(0, tk.END)  # Limpa o conteúdo atual
    entry_nome_fornecedor.insert(0, texto.upper())  # Insere o texto em maiúsculas

def converter_para_maiusculo_2(event):
    # Obtém o texto atual do Entry
    texto = entry_obs.get()
    # Converte o texto para maiúsculas
    entry_obs.delete(0, tk.END)  # Limpa o conteúdo atual
    entry_obs.insert(0, texto.upper())  # Insere o texto em maiúsculas

def formatar_cnpj():
    cnpj = entry_cnpj_fornecedor.get()
    cnpj = ''.join(filter(str.isdigit, cnpj))  # Remove tudo que não for dígito
    cnpj = cnpj[:14]  # Limita o CNPJ a 14 dígitos

    # Aplica a formatação do CNPJ
    formatado = ''
    if len(cnpj) > 2:
        formatado += cnpj[:2] + '.'
    if len(cnpj) > 5:
        formatado += cnpj[2:5] + '.'
    if len(cnpj) > 8:
        formatado += cnpj[5:8] + '/'
    if len(cnpj) > 12:
        formatado += cnpj[8:12] + '-'
    
    formatado += cnpj[12:]

    # Atualiza o Entry sem disparar o evento novamente
    entry_cnpj_fornecedor.delete(0, tk.END)
    entry_cnpj_fornecedor.insert(0, formatado)

def ao_digitar(event):
    global atraso_formatacao

    # Cancela a formatação anterior se o usuário ainda estiver digitando
    if atraso_formatacao:
        root.after_cancel(atraso_formatacao)
    
    # Aguarda 500ms após o último dígito antes de formatar
    atraso_formatacao = root.after(500, formatar_cnpj)


atraso_formatacao_telefone = None

def formatar_telefone():
    numero = entry_contato.get()
    numero = ''.join(filter(str.isdigit, numero))  # Remove tudo que não for dígito

    # Limita o número a 11 dígitos (máximo para celular)
    numero = numero[:11]

    # Formatação automática com base no comprimento do número
    if len(numero) <= 10:  # Número Fixo
        if len(numero) >= 2:
            formatado = f"({numero[:2]})"
            if len(numero) >= 6:
                formatado += f" {numero[2:6]}-{numero[6:]}"
            else:
                formatado += f" {numero[2:]}"
        else:
            formatado = numero

    else:  # Número de Celular
        formatado = f"({numero[:2]}) {numero[2:7]}-{numero[7:]}"

    # Atualiza o Entry sem disparar o evento novamente
    entry_contato.delete(0, tk.END)
    entry_contato.insert(0, formatado)

def ao_digitar_telefone(event):
    global atraso_formatacao_telefone

    # Cancela a formatação anterior se o usuário ainda estiver digitando
    if atraso_formatacao_telefone:
        root.after_cancel(atraso_formatacao_telefone)

    # Aguarda 500ms após o último dígito antes de formatar
    atraso_formatacao_telefone = root.after(500, formatar_telefone)

# Função para incrementar o contador global e retornar o novo valor
def incrementar_contador():
    global contador
    contador += 1
    return contador

def pre_lista():
    global pre_lista_entrada

    campos_vazios = verificar_campos_vazios(entry_cnpj_fornecedor,entry_nome_fornecedor,entry_seguimeto,entry_contato,entry_email,)

    if campos_vazios:

        lista_entry = [entry_cnpj_fornecedor,entry_nome_fornecedor,entry_seguimeto,entry_contato,entry_email,entry_obs]
        
        iten = []
        for i in lista_entry:
            iten.append(i.get())
        pre_lista_entrada.append(iten)
        
        tree.insert("", "end", values=iten)
        iten = []
        print(pre_lista_entrada)
    else:
        messagebox.showwarning("Campo Vazio", f"Verifique os capos de preenchimento.")

def deletar_pre_lista():
    global pre_lista_entrada

    selecionado = tree.selection()
    item_selecionado = selecionado[0]
    item_values = list(tree.item(item_selecionado, 'values')) 
    # print(item_values[1]) 

    lista_atualizada = []
    for i in pre_lista_entrada:
        print(i[1] != item_values[1])
        if item_values[1] != i[1]:

            lista_atualizada.append(i)
    
    pre_lista_entrada = []
    pre_lista_entrada = lista_atualizada

    tree.delete(selecionado)
    for i in pre_lista_entrada:
        print(i)

def finalizar_cadastro_fornecedor(lista,dados):
    workbook_estoque = load_workbook(dados)
    sheet_estoque = workbook_estoque.active  # Seleciona a planilha ativa

    for i in lista:
        try:
            sheet_estoque.append(i)
            # Salvar o arquivo
            workbook_estoque.save(dados)
            print(i)
        except:
            print("Erro ao Salvar")
    
    for item in tree.get_children():
        tree.delete(item)
    


def finalizar():
    finalizar_cadastro_fornecedor(pre_lista_entrada,base_de_dados)

# Criação da janela principal
root = tk.Tk()
root.title("Sistema de Saída de Produtos")

# Configurações gerais da interface
root.configure(bg="blue")

# Labels e entradas para o cadastro de produtos
tk.Label(root, text="Cadastro de fornecerdor", font=("Arial", 14), bg="blue", fg="black").grid(row=0, column=0, padx=1, pady=1, sticky="w")


tk.Label(root, text="CNPJ", bg="blue", fg="black").grid(row=1, column=0, padx=1, pady=1)
entry_cnpj_fornecedor = tk.Entry(root)
entry_cnpj_fornecedor.grid(row=2, column=0, padx=1, pady=1)
entry_cnpj_fornecedor.bind('<KeyRelease>', ao_digitar)
# entry_cnpj_fornecedor.bind("<KeyRelease>", converter_para_maiusculo)

tk.Label(root, text="Nome Fornecedor", bg="blue", fg="black").grid(row=1, column=1, padx=1, pady=1)
entry_nome_fornecedor = tk.Entry(root)
entry_nome_fornecedor.grid(row=2, column=1, padx=1, pady=1)
entry_nome_fornecedor.bind('<KeyRelease>', converter_para_maiusculo_1)
# entry_nome_fornecedor.bind("<KeyRelease>", converter_para_maiusculo)
# Data de entrada do produto
tk.Label(root, text="Seguimento", bg="blue", fg="black").grid(row=1, column=2, padx=1, pady=1)
entry_seguimeto = ttk.Combobox(root,values=lista_seguimento)
entry_seguimeto.grid(row=2, column=2, padx=1, pady=1)
entry_seguimeto.bind("<KeyRelease>", converter_para_maiusculo)

# Quantidade de produtos
tk.Label(root, text="Numero", bg="blue", fg="black").grid(row=1, column=3, padx=1, pady=1)
entry_contato = tk.Entry(root)
entry_contato.grid(row=2, column=3, padx=1, pady=1)
entry_contato.bind('<KeyRelease>', ao_digitar_telefone)
# entry_contato.bind("<KeyRelease>", converter_para_maiusculo)

# Motivo da entrada (e.g., compra, devolução)
tk.Label(root, text="E-mail", bg="blue", fg="black").grid(row=1, column=4, padx=1, pady=1)
entry_email = tk.Entry(root)
entry_email.grid(row=2, column=4, padx=1, pady=1)

# Valor unitário do produto
tk.Label(root, text="Obs", bg="blue", fg="black").grid(row=1, column=5, padx=1, pady=1)
entry_obs = tk.Entry(root)
entry_obs.grid(row=2, column=5, padx=1, pady=1)
entry_obs.bind('<KeyRelease>', converter_para_maiusculo_2)
entry_obs.bind("<KeyRelease>", converter_para_maiusculo)


# Botões de ação (Salvar, Deletar, Cancelar, Finalizar)
tk.Button(root, text="Salvar", bg="blue", fg="black",command=pre_lista).grid(row=1, column=8, padx=1, pady=1, sticky="ew")
tk.Button(root, text="Deletar", bg="blue", fg="black",command=deletar_pre_lista).grid(row=2, column=8, padx=1, pady=1, sticky="ew")

# Tabela (Treeview) para exibição dos produtos cadastrados
colunas = ("CNPJ","nome_fornecedor", "Seguimento", "telefone", "e-mail","OBS")
tree = ttk.Treeview(root, columns=colunas, show="headings")

tree.heading("CNPJ", text="CNPJ")
tree.heading("nome_fornecedor", text="Nome do Fornecedor")
tree.heading("Seguimento", text="Seguimento")
tree.heading("telefone", text="Telefone")
tree.heading("e-mail", text="E-mail")
tree.heading("OBS", text="Obs")

tree.column("CNPJ", width=10)
tree.column("nome_fornecedor", width=100)
tree.column("Seguimento", width=100)
tree.column("telefone", width=100)
tree.column("e-mail", width=100)
tree.column("OBS", width=100)

# Configurando o Treeview para preencher o espaço restante e ser responsivo
tree.grid(row=3, column=0, columnspan=17, rowspan=1, padx=1, pady=1, sticky="nsew")

# Botões de cancelamento e finalização
tk.Button(root, text="cancelar", bg="blue", fg="black",command=fechar_janela).grid(row=20, column=7, padx=1, pady=1, sticky="ew")
tk.Button(root, text="Finalizar", bg="blue", fg="black",command=finalizar).grid(row=20, column=8, padx=1, pady=1, sticky="ew")


# Executa a criação da interface se o script for rodado diretamente
if __name__ == "__main__":
    root.mainloop()

