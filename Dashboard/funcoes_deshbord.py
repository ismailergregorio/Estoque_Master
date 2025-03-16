import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from funcoes import*
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import subprocess
# import Entrada_pack.Janela_entrada as Janela_entrada
from datetime import date, timedelta

def coverter_dados_em_lista(dados):
    dados_para_coverter = abrir_arquivo(dados)

    nova_lista = []

    for x in dados_para_coverter.iter_rows(min_row=2, values_only=True):
        nova_lista.append(x)

    return nova_lista



def numero_para_mes(numero):
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    
    # Ajuste para números começarem de 1 (Janeiro é 1, Fevereiro é 2, etc.)
    if 1 <= numero <= 12:
        return meses[numero - 1]
    else:
        return "Número inválido"

def nomerar_lista_meses(lista_meses_para_nomear):
    lista_meses_nomeados = []
    for x in lista_meses_para_nomear:
        nome = numero_para_mes(x)
        lista_meses_nomeados.append(nome)
    
    return lista_meses_nomeados


def padao_tabela_data(dados,numero_coluna):
    lista_mes = []

    for itens in dados:
        mes = itens[numero_coluna].month
        if not mes in lista_mes:
            lista_mes.append(mes)
            mes = 0
        
    lista_mes = sorted(lista_mes)

    return lista_mes

def soma_valor_data(dados,numero_coluna_data,nemero_da_coluna_de_soma,datas):
    lista_valor_soma =[]
    lista_itesns = []

    for data_item in datas:
        valor_soma = 0
        for itens in dados:
            
            if itens[numero_coluna_data].month == data_item:
                lista_itesns.append(itens)
                try:
                    valor_soma = valor_soma + float(itens[nemero_da_coluna_de_soma])
                    valor_soma = float(f"{valor_soma:.2f}")
                except:
                    valor_soma = valor_soma + float(0)

        lista_valor_soma.append(valor_soma)
    
    return lista_valor_soma,lista_itesns

def seletor_de_coluna(dados):
    base_de_dados = abrir_arquivo(dados)

    primeira_linha = list(next(base_de_dados.iter_rows(min_row=1, max_row=1, values_only=True)))
    return primeira_linha

def filtra_por_coluna(dados,lista_coluna,coluna):
    lista_itens_da_reptidos = []
    valor_numero = 0

    for numero,colunas in enumerate(lista_coluna):
        if colunas == coluna:
            for itens_lista in dados:
                if not itens_lista[numero] in lista_itens_da_reptidos:
                    lista_itens_da_reptidos.append(itens_lista[numero])
                    valor_numero = numero
                    
    
    return valor_numero,lista_itens_da_reptidos

def resultado_grafico_vertical(lista_mestre,coluna_personalizada,data_filtro,caminho_dados,indice_soma):
    lista_colunas = seletor_de_coluna(caminho_dados)
    lista_itens = filtra_por_coluna(lista_mestre,lista_colunas,coluna_personalizada)

    valores = soma_valor_personalizado(lista_mestre,lista_itens[1],data_filtro,lista_itens[0],indice_soma)
    rotulos = lista_itens

    return rotulos[1],valores

def soma_valor_personalizado(dados,criterio_da_coluna_personalizada,datas,coluna_personalizada,numero_coluna_soma):
    lista_resultado_soma_ersonalizada = []
    valor_soma = 0
    for colunas in criterio_da_coluna_personalizada:
        for data_personalizada in datas:
            for itens in dados:
                if itens[coluna_personalizada] == colunas:
                    if itens[3].month == data_personalizada:
                        try:
                            valor_soma = valor_soma + float(itens[numero_coluna_soma])
                            valor_soma = float(f"{valor_soma:.2f}")
                        except:
                            valor_soma = valor_soma + float(0)
        lista_resultado_soma_ersonalizada.append(valor_soma)
        valor_soma = 0

    return lista_resultado_soma_ersonalizada

def filtro(dados,seletor,indice, data_inicial, data_final):
    lista_filtro = []
    # try:
    if seletor != "TODOS":
        for x in dados:
            if seletor == x[indice]:
                data = x[3].date()
                if data_inicial <= data <= data_final:  # Filtra os dados dentro do intervalo
                    lista_filtro.append(x)
    else:
        for x in dados:
            data = x[3].date()
            if data_inicial <= data <= data_final:  # Filtra os dados dentro do intervalo
                lista_filtro.append(x)
    # except:
    #     print("erro no filtro")

    return lista_filtro  # Retorna a lista filtrada

def mostrar_detalhes_nf_gerenciamento(event,tree_descrição):

    lista_completa = abrir_arquivo("Base de dados\entrada.xlsx")

    lista_tree = []
    item_selecionado = tree_descrição.selection()[0]

    item_values = list(tree_descrição.item(item_selecionado, 'values'))  # Seleciona o número do índice selecionado
    janela_detalhes = tk.Toplevel()
    # Itera sobre uma cópia da lista para evitar problemas de índice

    tk.Label(janela_detalhes, text=f"Nome do Funcionario: {item_values[0]}", bg="blue", fg="black").grid(row=0, column=0, sticky="w", padx=1, pady=1)
    tk.Label(janela_detalhes, text=f"Data de saida: {item_values[2]}", bg="blue", fg="black").grid(row=1, column=0, sticky="w", padx=1, pady=1)
    tk.Label(janela_detalhes, text=f"Setor do funcionario: {item_values[1]}", bg="blue", fg="black").grid(row=2, column=0, sticky="w", padx=1, pady=1)

    tree_itens = ttk.Treeview(janela_detalhes, columns=("Codigo","Produto","Setor","Data","Quantidade","Motivo","Valor unitario R$","Valor total R$"), show='headings',)

    tree_itens.heading("Codigo", text="Codigo")
    tree_itens.heading("Produto", text="Produto")
    tree_itens.heading("Setor", text="Setor",)
    tree_itens.heading("Data", text="Data")
    tree_itens.heading("Quantidade", text="Quantidade")
    tree_itens.heading("Motivo", text="Motivo")
    tree_itens.heading("Valor unitario R$", text="Valor unitario R$")
    tree_itens.heading("Valor total R$", text="Valor total R$")

    tree_itens.column("Codigo", stretch=True)
    tree_itens.column("Produto", width=250, anchor="center", stretch=True)
    tree_itens.column("Setor", width=100, stretch=True)
    tree_itens.column("Data", width=100, anchor="center", stretch=True)
    tree_itens.column("Quantidade", width=100, anchor="center", stretch=True)
    tree_itens.column("Motivo", width=250, anchor="center", stretch=True)
    tree_itens.column("Valor unitario R$", width=100, stretch=True)
    tree_itens.column("Valor total R$", width=100, anchor="center", stretch=True)

    tree_itens.grid(row=4, column=0, sticky="w", padx=1, pady=1, columnspan=2)
    valor_total = 0

    for item in lista_completa.iter_rows(min_row=2,values_only=True):
        # print(item)
        if f"NF° {item_values[4]},For:{item_values[1]}" == item[8]:
            #ITEN 4 REFERECE A COLUNA ONDE ESTA O NUMERO DA NFº DENTREO DA TABELA DE CADASTRO DE NFº 
            #ITEN 8 REFERECE A COLUNA ONDE ESTA O NUMERO DA NFº DENTREO DA TABELA DE ENTRADA
            print(item)
            tree_itens.insert("", 0, values=(item[0],item[1],item[2],item[3],item[4],item[5],item[6],item[7]))
            valor_atual = item[7]

            valor_total = float(valor_total) + float(valor_atual)

            valor_atual = 0

    tk.Label(janela_detalhes, text=f"Valor da NF: R${valor_total}", bg="blue", fg="black").grid(row=3, column=0, sticky="w", padx=1, pady=1)


if __name__ == "__main__":
    lista_mestre = coverter_dados_em_lista("Base de dados\entrada.xlsx")
    datas =  padao_tabela_data(lista_mestre,3)
    soma = soma_valor_data(lista_mestre,3,7,datas)

    print(resultado_grafico_vertical(lista_mestre,"Motivo",datas))
    