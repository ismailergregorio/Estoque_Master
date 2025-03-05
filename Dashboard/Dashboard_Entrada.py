
import sys
import os

# Adiciona a pasta raiz EstoqueA ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


from tkinter import*
from tkinter import ttk
import subprocess

from funcoes import*
from funcoes_deshbord import*
from Dashboard.confg_grafico_entrada_saida import*

base_de_dados_nf = "Base de dados\Cadastro NF.xlsx"
base_de_dados_entrada = "Base de dados\entrada.xlsx"

lista_mestre = coverter_dados_em_lista(base_de_dados_entrada)
lista_colunas = seletor_de_coluna(base_de_dados_entrada)

DASHBOARD_NF = "Dashboard\Dashboard_NF.py"
DASHBOARD_SAIDA = "Dashboard\Dashboard_Saida.py"
DASHBOARD_ENTRADA = "Dashboard\Dashboard_Entrada.py"
DASHBOARD_PERSONALIZADO = "Dashboard\Dashboard_Saida.py"

lista_itens = []

def abrir_janela(tipo_de_janela):
    if tipo_de_janela == DASHBOARD_NF:
        root.quit() 
        subprocess.Popen(["python", tipo_de_janela])  # Abre a Página 2
    elif tipo_de_janela == DASHBOARD_ENTRADA:
        root.quit()
        subprocess.Popen(["python", tipo_de_janela])  # Abre a Página 2
    elif tipo_de_janela == DASHBOARD_SAIDA:
        root.quit()
        subprocess.Popen(["python", tipo_de_janela])  # Abre a Página 2
    elif tipo_de_janela == DASHBOARD_PERSONALIZADO:
        root.quit()
        subprocess.Popen(["python", tipo_de_janela])
    else:
        print('pagina não encontrada')

def filtros_data(event):
    global lista_mestre
    global lista_itens

    # print(lista_mestre)
    lista = filtro(lista_mestre,seletor.get(),lista_itens[0],entry_data_inicial.get_date(),entry_data_final.get_date())

    tree.delete(*tree.get_children())

    for x in lista:
        tree.insert("", "end",values=x)

    data = padao_tabela_data(lista,3)
    valor = soma_valor_data(lista,3,7,data)
    lista_nomeada = nomerar_lista_meses(data)
    atualiza_grafico(lista_nomeada,valor[0])

    resultado1 = resultado_grafico_vertical(valor[1],"Motivo",data,base_de_dados_entrada,7) 
    atuliza_grafico_vertical1(frame_grafico_vertical1,resultado1[0],resultado1[1])

    resultado2 = resultado_grafico_vertical(valor[1],"Setor",data) 
    atuliza_grafico_vertical2(frame_grafico_vertical1,resultado2[0],resultado2[1],base_de_dados_entrada,7)


root = Tk()
root.title("Deshbord Entrada")
# root.state('zoomed')

Label(root, text="Deshbord Entrada", font=("Arial", 22)).grid(row=0, column=0,)

frame_buttos = Frame(root)
frame_buttos.grid(column=0, row=1 ,columnspan=4)

Button(frame_buttos, text="Deshbord NF", bg="blue", fg="white",width=20,command=lambda:abrir_janela(DASHBOARD_NF)).grid(row=0, column=0,sticky="ew",padx=5,pady=5)
Button(frame_buttos, text="Deshbord Entrada", bg="blue", fg="white",width=20,command=lambda:abrir_janela(DASHBOARD_ENTRADA),state="disabled").grid(row=0, column=1,sticky="ew",padx=5,pady=5)
Button(frame_buttos, text="Deshbord Saida", bg="blue", fg="white",command=lambda:abrir_janela(DASHBOARD_SAIDA)).grid(row=0, column=2,sticky="ew",padx=5,pady=5)
Button(frame_buttos, text="Deshbord Personalizado", bg="blue", fg="white",width=20,command=lambda:abrir_janela(DASHBOARD_PERSONALIZADO)).grid(row=0, column=3,sticky="ew",padx=5,pady=5)


frame_buttos_filtro_tabela = Frame(root)
frame_buttos_filtro_tabela.grid(column=0, row=2 ,columnspan=5)

Label(frame_buttos_filtro_tabela, text="Tabela Compras", font=("Arial", 22)).grid(row=0, column=0,columnspan=5)

coluna = lista_colunas

def atuliza_busca_filtro(event):
    global lista_itens
    lista_itens = filtra_por_coluna(lista_mestre,lista_colunas,coluna.get())
    print(lista_itens)
    seletor['values'] = lista_itens[1]
    

tk.Label(frame_buttos_filtro_tabela, text="Coluna", bg="blue", fg="white").grid(row=1, column=0, padx=5, pady=5)
coluna = ttk.Combobox(frame_buttos_filtro_tabela,values=coluna)
coluna.grid(row=1, column=1, padx=1,)
coluna.set("TODOS")
coluna.bind("<<ComboboxSelected>>", atuliza_busca_filtro)

tk.Label(frame_buttos_filtro_tabela, text="Filtro", bg="blue", fg="white").grid(row=1, column=2, padx=5, pady=5)
seletor = ttk.Combobox(frame_buttos_filtro_tabela)
seletor.grid(row=1, column=3, padx=1,)
seletor.set("TODOS")
seletor.bind("<<ComboboxSelected>>", filtros_data)
# Filtros de data
tk.Label(frame_buttos_filtro_tabela, text="Data inicial", bg="blue", fg="white").grid(row=1, column=4, padx=5, pady=5)
entry_data_inicial = DateEntry(frame_buttos_filtro_tabela, width=20, background='darkblue', foreground='white', borderwidth=2,)
entry_data_inicial.grid(row=1, column=5, padx=5, pady=5)
entry_data_inicial.bind("<<DateEntrySelected>>", filtros_data)

tk.Label(frame_buttos_filtro_tabela, text="Data final", bg="blue", fg="white").grid(row=1, column=6, padx=5, pady=5)
entry_data_final = DateEntry(frame_buttos_filtro_tabela, width=20, background='darkblue', foreground='white', borderwidth=2,)
entry_data_final.grid(row=1, column=7, padx=7, pady=5)
entry_data_final.bind("<<DateEntrySelected>>", filtros_data)

frame_tabela = Frame(root)
frame_tabela.grid(row=3,column=0,columnspan=4)

colunas = ("Codigo", "Produto", "Setor", "Data de Entrada", "Quantidade", "Motivo", "Valor Unitario", "Valor Total", "Obs")
tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings")

# Definição das colunas e seus respectivos cabeçalhos
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
tree.column("Codigo", width=30)
tree.column("Produto", width=100)
tree.column("Setor", width=100)
tree.column("Data de Entrada", width=100)
tree.column("Quantidade", width=100)
tree.column("Motivo", width=100)
tree.column("Valor Unitario", width=100)
tree.column("Valor Total", width=100)
tree.column("Obs", width=100)


tree.grid(row=4, column=1, columnspan=12, rowspan=21, padx=5, pady=5, sticky="nsew")
tree.bind("<Double-1>", lambda event: mostrar_detalhes_nf_gerenciamento(event,tree))

for x in lista_mestre:
    tree.insert("", "end",values=x)

frame_grafico_master = Frame(root)
frame_grafico_master.grid(row=5,column=0)

frame_grafico_vertical1 = Frame(root)
frame_grafico_vertical1.grid(row=2,column=5,rowspan=2)

frame_grafico_vertical2 = Frame(root)
frame_grafico_vertical2.grid(row=5,column=5)

data_filtrada = padao_tabela_data(lista_mestre,3)
valor = soma_valor_data(lista_mestre,3,7,data_filtrada)
lista_nomeada = nomerar_lista_meses(data_filtrada) 
gera_grafico(frame_grafico_master,lista_nomeada,valor[0])


resultado1 = resultado_grafico_vertical(valor[1],"Motivo",data_filtrada,dados_de_entrada,7)
print("inicio",resultado1) 
gera_grafico_vertical1(frame_grafico_vertical1,resultado1[0],resultado1[1])


resultado2 = resultado_grafico_vertical(lista_mestre,"Setor",data_filtrada,dados_de_entrada,7)
print("inicio",resultado2)
gera_grafico_vertical2(frame_grafico_vertical2,resultado2[0],resultado2[1])

root.mainloop()