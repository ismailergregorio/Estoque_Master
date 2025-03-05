import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import tkinter as tk
from funcoes import*
from tkinter import ttk
from tkcalendar import DateEntry

def adicionar_itens_na_lista(tabela,codigo,nome_produto, setor,data,quantidade,motivo_saida,funcionario,setor_funcionario,valor_total,obs):

    maior = 0

    for item_id in tabela.get_children():  # Itera pelos IDs das linhas
        valores = tabela.item(item_id, "values")  # Obtém os valores da linha
        maior == int(valores[0])
        if int(valores[0]) > maior:
            maior = int(valores[0])
    
    maior = maior + 1

    # print(maior, codigo.get(),nome_produto.get(), setor.get(),data.get_date(),quantidade.get(),tipo_entrada.get(),valor_unitario.get(),valor_total.get(),obs.get())

    tabela.insert("", "end", values=(maior,codigo.get(),nome_produto.get(), setor.get(),data.get_date(),quantidade.get(),motivo_saida.get(),funcionario.get(),setor_funcionario.get(),valor_total.get(),obs.get()))
    itens = [maior,codigo.get(),nome_produto.get(), setor.get(),data.get_date(),quantidade.get(),motivo_saida.get(),funcionario.get(),setor_funcionario.get(),valor_total.get(),obs.get()]
    
    return itens

def calcular_total(codigo,quatidade):
    """
    Calcula o valor total do produto com base na quantidade e no preço unitário.
    """
    dados = abrir_arquivo(dados_de_estoque)

    try:
        for i in dados.iter_rows(min_row=2, values_only=True):
            if int(i[0]) == int(codigo):
                valor_medio = float(i[8])  # Obtém o preço unitário do entry
                total = quatidade * valor_medio  # Calcula o valor total
                valor = float(f"{total:.2f}")

        return valor # Insere o valor calculado
    except ValueError:
        return "0.00"
    
def finalizar_itens_saida(lista,dados):
    workbook_dados = load_workbook(dados)
    sheet_dados = workbook_dados.active  # Seleciona a planilha ativa

    workbook_estoque = load_workbook(dados_de_estoque)
    sheet_estoque = workbook_estoque.active  # Seleciona a planilha ativa

    iten =[]
    for i in lista:
        try:
            for n,x in enumerate(i):
                if not n == 0:
                    iten.append(x)

            sheet_dados.append(iten)
            # Salvar o arquivo
            workbook_dados.save(dados)
        except:
            print(f"Erro ao Salvar o iten na Tabela de Entrada {x}")
            alerta_erro("Erro ao Salvar","Verifique os dados")

        try:
            for linha in sheet_estoque.iter_rows(min_row=2,values_only=False):
                if int(linha[0].value) == int(iten[0]):  # Verifica o ID do item
                    linha[5].value = float(f"{float(linha[5].value) + float(iten[4]):.2}")  # Atualiza o valor na coluna 5 (estoque)
                    linha[6].value = float(f"{float(linha[6].value) - float(iten[4]):.2}")
                    linha[7].value = float(f"{float(linha[7].value) - float(iten[8]):.2}")
                    valores_linha = [celula.value for celula in linha]  # Extrai os valores de todas as células na linha
                    print(f"Item ajustado: {valores_linha}")
                    workbook_estoque.save(dados_de_estoque)
                    break
        except:
            print(f"Erro ao Salvar o iten no estoque {valores_linha}")
            alerta_erro("Erro ao Salvar","Verifique os dados")
            return False

        iten =[]
    return True

def adicionar_itens_na_lista_admissao(tabela,nome,funcao):

    maior = 0

    for item_id in tabela.get_children():  # Itera pelos IDs das linhas
        valores = tabela.item(item_id, "values")  # Obtém os valores da linha
        maior == int(valores[0])
        if int(valores[0]) > maior:
            maior = int(valores[0])
    
    maior = maior + 1

    # print(maior, codigo.get(),nome_produto.get(), setor.get(),data.get_date(),quantidade.get(),tipo_entrada.get(),valor_unitario.get(),valor_total.get(),obs.get())

    tabela.insert("", "end", values=(maior,nome.get(),funcao.get()))
    itens = [maior,nome.get(),funcao.get()]
    
    return itens

def detalhes(event, tree, lista):

    def mostrar_detalhes(tree, lista):
        item_selecionado = tree.selection()[0]

        item_values = list(tree.item(item_selecionado, 'values'))  # Seleciona o número do índice selecionado
        janela_detalhes = tk.Toplevel()
        # Itera sobre uma cópia da lista para evitar problemas de índice

        for item in lista[:]:  # Usar lista[:] cria uma cópia superficial da lista
            if item_values[1] == item[6]:
                print(item)

        tk.Label(janela_detalhes, text=f"Nome do Funcionario:{item_values[1]}", bg="blue", fg="black").grid(row=0, column=0, sticky="w", padx=1, pady=1)
        tk.Label(janela_detalhes, text=f"Setor do funcionario:{item_values[2]}", bg="blue", fg="black").grid(row=2, column=0, sticky="w", padx=1, pady=1)

        tree_itens = ttk.Treeview(janela_detalhes, columns=("Produto","Setor","Data","Quantidade","Valor"), show='headings',)

        tree_itens.heading("Produto", text="Produto")
        tree_itens.heading("Setor", text="Setor")
        tree_itens.heading("Data", text="Data")
        tree_itens.heading("Quantidade", text="Quantidade")
        tree_itens.heading("Valor", text="Valor")

        tree_itens.column("Produto", width=250, anchor="center", stretch=True)
        tree_itens.column("Setor", width=100, stretch=True)
        tree_itens.column("Data", width=100, stretch=True)
        tree_itens.column("Quantidade", width=100, anchor="center", stretch=True)
        tree_itens.column("Valor", width=100, anchor="center", stretch=True)

        tree_itens.grid(row=3, column=0, sticky="w", padx=1, pady=1, columnspan=2)

        for item in lista[:]:  # Usar lista[:] cria uma cópia superficial da lista
            if item_values[1] == item[7]:
                tree_itens.insert("", 0, values=(item[2],item[3],item[4],item[5],item[9]))

    mostrar_detalhes(tree, lista)

