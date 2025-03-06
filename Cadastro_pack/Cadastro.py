import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import funcoes
from Cadastro_pack.funcoes_cadastro_produtos import*
from tkinter import *
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import datetime  # Trabalhar com datas
import datetime  # Biblioteca datetime para manipulação de datas
from datetime import date, timedelta


cadastro = "Base de dados\CADASTRO.xlsx"
dados_de_saida = "Base de dados\saida.xlsx"
dados_de_entrada = "Base de dados\entrada.xlsx"
dados_de_estoque = "Base de dados\Estoque.xlsx"
dados_de_cadastro_nf = "Base de dados\Cadastro NF.xlsx"
dados_de_cadastro_fornecedor = "Base de dados\Cadastro Fornecedor.xlsx"

base_de_dados = cadastro

data_atual = date.today()

# Subtrair 360 dias
data_menos_360_dias = data_atual - timedelta(days=360)

def atualizar_filtros(event):
    filtro(tabela_frame,base_de_dados,entry_data_inicial.get_date(),entry_data_final.get_date(),entry_buscar_produto.get(),3,1)

# def converter_para_maiusculo(event):
#     # Obtém o texto atual do Entry
#     texto = entry_nome_produto.get()
#     # Converte o texto para maiúsculas
#     entry_nome_produto.delete(0, tk.END)  # Limpa o conteúdo atual
#     entry_nome_produto.insert(0, texto.upper())  # Insere o texto em maiúsculas

def converter_para_maiusculo_1(event):
    # Obtém o texto atual do Entry
    texto = entry_buscar_produto.get()
    # Converte o texto para maiúsculas
    entry_buscar_produto.delete(0, tk.END)  # Limpa o conteúdo atual
    entry_buscar_produto.insert(0, texto.upper())  # Insere o texto em maiúsculas

def verificar_campos_vazios(entry_nome_produto, entry_setor, entry_data, entry_estoque_m, entry_estoque_d):
    campos = {
        "Nome do Produto": entry_nome_produto.get(),
        "Setor": entry_setor.get(),
        "Data": entry_data.get_date(),
        "Estoque Mínimo": entry_estoque_m.get(),
        "Estoque Desejável": entry_estoque_d.get()
    }

    # Verifique os campos vazios
    campos_vazios = [campo for campo, valor in campos.items() if not str(valor).strip()]

    if campos_vazios:
        mensagem = "Os seguintes campos estão vazios:\n" + "\n".join(campos_vazios)
        print(mensagem)  # Ou exiba uma mensagem usando messagebox
        return False
    return True

def salvar():
    if verificar_campos_vazios(entry_nome_produto, entry_setor, entry_data, entry_estoque_m, entry_estoque_d):
        cadastrar_produto(base_de_dados,entry_nome_produto.get(),entry_setor.get(),entry_data.get_date(),entry_estoque_m.get(),entry_estoque_d.get(),entry_obs.get())
        carregar_dados(base_de_dados,tabela_frame)

        entry_nome_produto.delete(0, tk.END)
        entry_setor.delete(0, tk.END)
        # entry_data.set_date("")
        entry_estoque_m.delete(0, tk.END)
        entry_estoque_d.delete(0, tk.END)
        entry_obs.delete(0, tk.END)

    else:
        print("Por favor, preencha todos os campos.")
        
def deletar():
    deletar_item_cadastro(tabela_frame,base_de_dados)
    carregar_dados(base_de_dados,tabela_frame)

    
def abrir_main():
    root.quit()  # Fecha a janela atual

root = Tk()
root.title("Sistema de Cadastro de Produtos")

# Labels e entradas para o cadastro de produto
Label(root, text="Cadastro", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, columnspan=2)

Button(root, text="Sair", bg="blue", fg="white",
            command=abrir_main).grid(row=0, column=6, padx=5, pady=5, sticky="ew")

Label(root, text="Nome produto", bg="blue", fg="white").grid(row=1, column=0, padx=5, pady=5)
entry_nome_produto = Entry(root)
entry_nome_produto.grid(row=2, column=0, padx=5, pady=5)
entry_nome_produto.bind("<KeyRelease>", converter_para_maiusculo)

setores = ["RH", "ESCRITORIO", "SEGURANÇA", "COSINHA LIMPESA", "COSINHA COMIDA", "TRADE", "OPERACIONAL"]
Label(root, text="Setor", bg="blue", fg="white").grid(row=1, column=1, padx=5, pady=5)
entry_setor = ttk.Combobox(root, values=setores)
entry_setor.grid(row=2, column=1, padx=5, pady=5)

Label(root, text="Data", bg="blue", fg="white").grid(row=1, column=2, padx=5, pady=5)
entry_data = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
entry_data.grid(row=2, column=2, padx=5, pady=5)

Label(root, text="Estoque m", bg="blue", fg="white").grid(row=1, column=3, padx=5, pady=5)
entry_estoque_m = Entry(root)
entry_estoque_m.grid(row=2, column=3, padx=5, pady=5)
entry_estoque_m.bind("<KeyRelease>",lambda event: calcular_estoque_desejavel(event,entry_estoque_m,entry_estoque_d))

Label(root, text="Estoque d", bg="blue", fg="white").grid(row=1, column=4, padx=5, pady=5)
entry_estoque_d = Entry(root)
entry_estoque_d.grid(row=2, column=4, padx=5, pady=5)

Label(root, text="Obs.", bg="blue", fg="white").grid(row=1, column=5, padx=5, pady=5)
entry_obs = Entry(root)
entry_obs.grid(row=2, column=5, padx=5, pady=5)

# Botões de ação
Button(root, text="Salvar", bg="blue", fg="white",command=lambda:salvar()
            ).grid(row=3, column=0, padx=5, pady=5, sticky="ew")

Button(root, text="Deletar seleção", bg="blue", fg="white",
            command=lambda:deletar()).grid(row=4, column=0, padx=5, pady=5, sticky="ew")

def atualizar_e_converter(event):
    converter_para_maiusculo_1(event)
    atualizar_filtros(event)

Label(root, text="Buscar produto", bg="blue", fg="white").grid(row=5, column=0, padx=5, pady=5, sticky="ew")
entry_buscar_produto = Entry(root)
entry_buscar_produto.grid(row=6, column=0, padx=5, pady=5)
entry_buscar_produto.bind("<KeyRelease>", atualizar_e_converter)
entry_buscar_produto.bind("<KeyRelease>", converter_para_maiusculo)
# Filtros de data
Label(root, text="Data inicial", bg="blue", fg="white").grid(row=7, column=0, padx=5, pady=5)
entry_data_inicial = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
entry_data_inicial.grid(row=8, column=0, padx=5, pady=5)
entry_data_inicial.set_date(data_menos_360_dias)
entry_data_inicial.bind("<<DateEntrySelected>>", atualizar_filtros)

Label(root, text="Data final", bg="blue", fg="white").grid(row=9, column=0, padx=5, pady=5)
entry_data_final = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
entry_data_final.grid(row=10, column=0, padx=5, pady=5)
entry_data_final.bind("<<DateEntrySelected>>", atualizar_filtros)

tabela_frame = tabela(base_de_dados,root,0,)
tabela_frame.grid(row=3, column=1, columnspan=11, rowspan=20, padx=5, pady=5, sticky="nsew")
carregar_dados(base_de_dados,tabela_frame)

root.mainloop()