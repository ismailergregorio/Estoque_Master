import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from tkinter import*
from tkinter import ttk
import subprocess

from confg_grafico_master import*
from funcoes import*
from funcoes_deshbord import*

base_de_dados_nf = "Base de dados\Cadastro NF.xlsx"

DASHBOARD_NF = "Dashboard\Dashboard_NF.py"
DASHBOARD_SAIDA = "Dashboard\Dashboard_Saida.py"
DASHBOARD_ENTRADA = "Dashboard\Dashboard_Entrada.py"
DASHBOARD_PERSONALIZADO = "Dashboard\Dashboard_Saida.py"

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

lista_mestre = coverter_dados_em_lista(base_de_dados_nf)

def filtros_data(event):
    global lista_mestre

    # print(lista_mestre)
    lista = filtro(lista_mestre,seletor.get(),entry_data_inicial.get_date(),entry_data_final.get_date())

    tree.delete(*tree.get_children())

    for x in lista:
        tree.insert("", "end",values=x)

    data = padao_tabela_data(lista)
    valor = soma_valor_data(lista,data)
    lista_nomeada = nomerar_lista_meses(data)

    atualiza_grafico(lista_nomeada,valor)


root = Tk()
root.title("Deshbord Master")
# root.state('zoomed')

Label(root, text="Deshbord Master", font=("Arial", 22)).grid(row=0, column=0,)

frame_buttos = Frame(root)
frame_buttos.grid(column=0, row=1 ,columnspan=4)

Button(frame_buttos, text="Deshbord NF", bg="blue", fg="white",width=20,command=lambda:abrir_janela(DASHBOARD_NF),state="disabled").grid(row=0, column=0,sticky="ew",padx=5,pady=5)
Button(frame_buttos, text="Deshbord Entrada", bg="blue", fg="white",width=20,command=lambda:abrir_janela(DASHBOARD_ENTRADA)).grid(row=0, column=1,sticky="ew",padx=5,pady=5)
Button(frame_buttos, text="Deshbord Saida", bg="blue", fg="white",command=lambda:abrir_janela(DASHBOARD_SAIDA)).grid(row=0, column=2,sticky="ew",padx=5,pady=5)
Button(frame_buttos, text="Deshbord Personalizado", bg="blue", fg="white",width=20,command=lambda:abrir_janela(DASHBOARD_PERSONALIZADO)).grid(row=0, column=3,sticky="ew",padx=5,pady=5)

frame_buttos_filtro_tabela = Frame(root)
frame_buttos_filtro_tabela.grid(column=0, row=2 ,columnspan=5)

Label(frame_buttos_filtro_tabela, text="Tabela Compras", font=("Arial", 22)).grid(row=0, column=0,columnspan=5)
valores = ["TODOS","COMPRA", "DEVOLUÇÃO", "TROCA", "OUTROS"]
tk.Label(frame_buttos_filtro_tabela, text="Filtro", bg="blue", fg="white").grid(row=1, column=0, padx=5, pady=5)
seletor = ttk.Combobox(frame_buttos_filtro_tabela,values=valores)
seletor.grid(row=1, column=1, padx=1,)
seletor.set("TODOS")
seletor.bind("<<ComboboxSelected>>", filtros_data)
# Filtros de data
tk.Label(frame_buttos_filtro_tabela, text="Data inicial", bg="blue", fg="white").grid(row=1, column=2, padx=5, pady=5)
entry_data_inicial = DateEntry(frame_buttos_filtro_tabela, width=20, background='darkblue', foreground='white', borderwidth=2,)
entry_data_inicial.grid(row=1, column=3, padx=5, pady=5)
entry_data_inicial.bind("<<DateEntrySelected>>", filtros_data)

tk.Label(frame_buttos_filtro_tabela, text="Data final", bg="blue", fg="white").grid(row=1, column=4, padx=5, pady=5)
entry_data_final = DateEntry(frame_buttos_filtro_tabela, width=20, background='darkblue', foreground='white', borderwidth=2,)
entry_data_final.grid(row=1, column=5, padx=5, pady=5)
entry_data_final.bind("<<DateEntrySelected>>", filtros_data)

frame_tabela = Frame(root)
frame_tabela.grid(row=3,column=0,columnspan=4)

tree = ttk.Treeview(frame_tabela, columns=["CNPJ","Nome do Fornecedor", "Seguimento","Data de Entrada","N° NF","VALOR NF","Motivo","Obs"], show="headings")

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

tree.grid(row=4, column=1, columnspan=12, rowspan=21, padx=5, pady=5, sticky="nsew")
tree.bind("<Double-1>", lambda event: mostrar_detalhes_nf_gerenciamento(event,tree))

for x in lista_mestre:
    tree.insert("", "end",values=x)

frame_grafico_master = Frame(root)
frame_grafico_master.grid(row=5,column=0,columnspan=4)

data = padao_tabela_data(lista_mestre,3)
valor = soma_valor_data(lista_mestre,3,5,data)
lista_nomeada = nomerar_lista_meses(data) 
gera_grafico(frame_grafico_master,lista_nomeada,valor[0])


root.mainloop()