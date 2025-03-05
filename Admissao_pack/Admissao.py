import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import tkinter as tk
from funcoes import*
from tkinter import ttk
from tkcalendar import DateEntry

from funçoes_admissao import*

# Definindo os caminhos para os arquivos Excel
base_de_dados_saida = "Base de dados\saida.xlsx"
base_de_dados_entrada = "Base de dados\entrada.xlsx"
base_de_dados_cadastro = "Base de dados\CADASTRO.xlsx"

lista_de_funcoes = ["OPERAÇÃO", "PROMOTOR", "SUPERVISOR", "VENDEDOR", "ADMINISTRATIVO", "OUTROS"]

quantidade_camisa = 3
quantidade_calca = 2
quatidade_jaleco = 3
quantidade_epi = 1
quatidade_outros = 1

lista_pre_salva = []

def deleta_entrada():
    entry_nome_funcionario.delete(0, tk.END)
    entry_funcao_funcionario.delete(0, tk.END)
    entry_data_admicao.delete(0, tk.END)

    entry_blusa.delete(0, tk.END)
    entry_calca.delete(0, tk.END)
    entry_jaleco.delete(0, tk.END)
    entry_epi.delete(0, tk.END)


    check_var_tesoura.set(0)
    check_var_flanela.set(0)
    check_var_fita_adesiva.set(0)
    check_var_caneta.set(0)
    check_var_caderno.set(0)
    check_var_espanador.set(0)
    check_var_marcador_permante.set(0)

def verifica_funcao(event):
    label_blusa.grid_forget()
    entry_blusa.grid_forget()

    label_jaleco.grid_forget()
    entry_jaleco.grid_forget()

    label_calca.grid_forget()
    entry_calca.grid_forget()

    label_epi.grid_forget()
    entry_epi.grid_forget()

    frame_checkboxes.grid_forget()

    dados = abrir_arquivo(base_de_dados_cadastro)

    def resultado_lista_itens(chave):
        nomes_itens = []
        for palavra in chave:
            for lista in dados.iter_rows(min_row=2 ,values_only=True):
                if palavra in lista[1]:
                    nomes_itens.append(lista[1])
        
        return nomes_itens

    if entry_funcao_funcionario.get() == "OPERAÇÃO":

        label_blusa.grid(row=5,column=0,columnspan=5,sticky="w")
        entry_blusa.grid(row=6,column=0,columnspan=5)

        label_calca.grid(row=9,column=0,columnspan=5,sticky="w")
        entry_calca.grid(row=10,column=0,columnspan=5)

        label_epi.grid(row=11,column=0,columnspan=5,sticky="w")
        entry_epi.grid(row=12,column=0,columnspan=5)

        blusa = resultado_lista_itens(["CAMISA OPERACIONAL"])
        calca = resultado_lista_itens(["CALÇA OPERAÇÃO"])
        epi = resultado_lista_itens(["BOTINA","SAPATO"])
  
    elif entry_funcao_funcionario.get() == "PROMOTOR":

        label_blusa.grid(row=5,column=0,columnspan=5,sticky="w")
        entry_blusa.grid(row=6,column=0,columnspan=5)

        label_jaleco.grid(row=7,column=0,columnspan=5,sticky="w")
        entry_jaleco.grid(row=8,column=0,columnspan=5)

        label_calca.grid(row=9,column=0,columnspan=5,sticky="w")
        entry_calca.grid(row=10,column=0,columnspan=5)

        label_epi.grid(row=11,column=0,columnspan=5,sticky="w")
        entry_epi.grid(row=12,column=0,columnspan=5)

        blusa = resultado_lista_itens(["CAMISA PROMOTOR"])
        calca = resultado_lista_itens(["CALÇA PROMOTOR"])
        epi = resultado_lista_itens(["BOTINA"])
        jaleco = resultado_lista_itens(["JALECO"])

        frame_checkboxes.grid(row=13,column=0,columnspan=5)

    elif entry_funcao_funcionario.get() == "SUPERVISOR" or entry_funcao_funcionario.get() == "VENDEDOR":

        label_blusa.grid(row=5,column=0,columnspan=5,sticky="w")
        entry_blusa.grid(row=6,column=0,columnspan=5)

        if entry_funcao_funcionario.get() == "VENDEDOR":
            blusa = resultado_lista_itens(["CAMISA VENDAS"])
            frame_checkboxes.grid(row=13,column=0,columnspan=5)
        else:
            blusa = resultado_lista_itens(["CAMISA SUPERVISOR"])
    
    elif entry_funcao_funcionario.get() == "ADMINISTRATIVO":

        label_blusa.grid(row=5,column=0,columnspan=5,sticky="w")
        entry_blusa.grid(row=6,column=0,columnspan=5)
        blusa = resultado_lista_itens(["CAMISA"])
    
    elif entry_funcao_funcionario.get() == "OUTROS":

        label_blusa.grid(row=5,column=0,columnspan=5,sticky="w")
        entry_blusa.grid(row=6,column=0,columnspan=5)

        label_jaleco.grid(row=7,column=0,columnspan=5,sticky="w")
        entry_jaleco.grid(row=8,column=0,columnspan=5)

        label_calca.grid(row=9,column=0,columnspan=5,sticky="w")
        entry_calca.grid(row=10,column=0,columnspan=5)

        label_epi.grid(row=11,column=0,columnspan=5,sticky="w")
        entry_epi.grid(row=12,column=0,columnspan=5)

        frame_checkboxes.grid(row=13,column=0,columnspan=5)

        blusa = resultado_lista_itens(["CAMISA"])
        calca = resultado_lista_itens(["CALÇA"])
        epi = resultado_lista_itens(["BOTINA","SAPATO"])
        jaleco = resultado_lista_itens(["JALECO"])

    if entry_funcao_funcionario.get() == "OUTROS" or entry_funcao_funcionario.get() == "PROMOTOR" or entry_funcao_funcionario.get() == "OPERAÇÃO":
        entry_blusa['values'] = blusa
        entry_calca['values'] = calca
        entry_epi['values'] = epi
        if entry_funcao_funcionario.get() == "OUTROS" or entry_funcao_funcionario.get() == "PROMOTOR":
            entry_jaleco['values'] = jaleco 
    
    if entry_funcao_funcionario.get() == "SUPERVISOR" or entry_funcao_funcionario.get() == "VENDEDOR" or entry_funcao_funcionario.get() == "ADMINISTRATIVO":
        entry_blusa['values'] = blusa

def verifica_itens():
    lista_total = []
    try:
        dados_de_uniformes = [entry_blusa,entry_calca,entry_jaleco,entry_epi]
        dados_kit_basico =   [check_var_tesoura,check_var_fita_adesiva,check_var_flanela,check_var_caneta,check_var_caderno,check_var_marcador_permante,check_var_espanador]

        for dados in dados_de_uniformes:
            if dados.get() != "":
                lista_total.append(dados.get())

        for dados in dados_kit_basico:
            if dados.get() != "":
                lista_total.append(dados.get())
    except:
        print("Problema na selecão do itens")
        
    return lista_total

indice = 0
def montar_lista_itens_admissao(lista):
    global indice
    modelo_lista = []
    lista_oficial = []
    dados = abrir_arquivo(base_de_dados_cadastro)
    try:
        for x in lista:
            for i in dados.iter_rows(min_row=2,values_only=True):
                if i[1] == x:
                    indice = indice + 1
                    modelo_lista.append(indice)
                    modelo_lista.append(f"{i[0]}")
                    modelo_lista.append(i[1])
                    modelo_lista.append(i[2])
                    modelo_lista.append(entry_data_admicao.get_date())

                    if "CAMISA" in i[1]:
                        modelo_lista.append(quantidade_camisa)
                    elif "CALÇA" in i[1]:
                        modelo_lista.append(quantidade_calca)
                    elif "BOTINA" in i[1] or "SAPATO" in i[1]:
                        modelo_lista.append(quantidade_epi)
                    elif "JALECO" in i[1]:
                        modelo_lista.append(quatidade_jaleco)
                    else:
                        modelo_lista.append(quatidade_outros)
                    
                    modelo_lista.append("ADMISSÃO")
                    modelo_lista.append(entry_nome_funcionario.get())
                    modelo_lista.append(entry_funcao_funcionario.get())

                    modelo_lista.append(calcular_total(i[0],modelo_lista[5]))

                    lista_oficial.append(modelo_lista)
            modelo_lista = []
    except:
        print("Erro ao criar a lista de itens do funcionario")
    
    return lista_oficial

def salvar_pre_lista():
    global lista_pre_salva 
    lista_itens = verifica_itens()

    lista_pre_salva = montar_lista_itens_admissao(lista_itens)
    adicionar_itens_na_lista_admissao(tree,entry_nome_funcionario,entry_funcao_funcionario)
    deleta_entrada()
    for i in lista_pre_salva:
        print(i)

def deletar_pre_lista():
    global lista_pre_salva

    selecionado = tree.selection()
    item_selecionado = selecionado[0]
    item_values = list(tree.item(item_selecionado, 'values')) 
    # print(item_values[1]) 

    lista_atualizada = []
    for i in lista_pre_salva:
        if item_values[1] != i[7]:
            lista_atualizada.append(i)
    
    lista_pre_salva = []
    lista_pre_salva = lista_atualizada
    tree.delete(selecionado)
    for i in lista_pre_salva:
        print(i)

def finalizar():
    finalizar_itens_saida(lista_pre_salva,base_de_dados_saida)

# Configuração da janela principal
janela = tk.Tk()
janela.title("Selecionar Itens")

tk.Label(janela, text="Nome", bg="blue", fg="black").grid(row=0, column=0, sticky="w", padx=1, pady=1, columnspan=3)
entry_nome_funcionario = tk.Entry(janela, width=59)
entry_nome_funcionario.grid(row=1, column=0, sticky="w", columnspan=3, padx=1, pady=1)

tk.Label(janela, text="Data", bg="blue", fg="black").grid(row=2, column=1, padx=1, pady=1,sticky="w")
entry_data_admicao = DateEntry(janela, width=12, background='darkblue', foreground='white', borderwidth=2)
entry_data_admicao.grid(row=3, column=1, padx=1, pady=1,sticky="w")

tk.Label(janela, text="Função", bg="blue", fg="black").grid(row=2, column=0, padx=1, pady=1, sticky="w")
entry_funcao_funcionario = ttk.Combobox(janela, values=lista_de_funcoes)
entry_funcao_funcionario.grid(row=3, column=0, padx=1, pady=1, sticky="w")
entry_funcao_funcionario.bind("<<ComboboxSelected>>", verifica_funcao)

# Widgets criados uma vez para serem usados depois
label_blusa = tk.Label(janela, text="Blusa", bg="blue", fg="black")
entry_blusa = ttk.Combobox(janela, width=59)

label_calca = tk.Label(janela, text="Calça", bg="blue", fg="black")
entry_calca = ttk.Combobox(janela, width=59)

label_jaleco = tk.Label(janela, text="Jaleco", bg="blue", fg="black")
entry_jaleco = ttk.Combobox(janela, width=59)

label_epi = tk.Label(janela, text="EPI", bg="blue", fg="black")
entry_epi = ttk.Combobox(janela, width=59)


# Criação dos checkboxes e posicionamento com grid para kit basico
check_var_tesoura = tk.StringVar()
check_var_fita_adesiva = tk.StringVar()
check_var_flanela = tk.StringVar()
check_var_caneta = tk.StringVar()
check_var_caderno = tk.StringVar()
check_var_marcador_permante = tk.StringVar()
check_var_espanador = tk.StringVar()

frame_checkboxes = tk.Frame(janela, bg='#f0f0f0', bd=2, relief="groove",width=350, height=120)

frame_checkboxes.grid_propagate(False)

frame_checkboxes.columnconfigure(0, weight=1)
frame_checkboxes.rowconfigure([0, 1, 2], weight=1)

tk.Label(frame_checkboxes, text="Kit Basico", font=("Arial", 11), fg="black").grid(row=0, column=0, padx=1, pady=1, sticky="w")

chk_tesoura = ttk.Checkbutton(frame_checkboxes, text="Tesoura", variable=check_var_tesoura ,onvalue="TESOURA", offvalue="")
chk_tesoura.grid(row=1, column=0, padx=5, pady=5, sticky="w")

chk_flanela = ttk.Checkbutton(frame_checkboxes, text="Flanela", variable=check_var_flanela ,onvalue="FLANELA", offvalue="")
chk_flanela.grid(row=1, column=1, padx=5, pady=5, sticky="w")

chk_fita_adesiva = ttk.Checkbutton(frame_checkboxes, text="Fita Adesiva", variable=check_var_fita_adesiva, onvalue="FITA ADESIVA(TRASPARENTE) CX", offvalue="")
chk_fita_adesiva.grid(row=1, column=2, padx=5, pady=5, sticky="w")

chk_caneta = ttk.Checkbutton(frame_checkboxes, text="Caneta", variable=check_var_caneta ,onvalue="CANETA", offvalue="")
chk_caneta.grid(row=2, column=0, padx=5, pady=5, sticky="w")

chk_caderno = ttk.Checkbutton(frame_checkboxes, text="Caderno", variable=check_var_caderno ,onvalue="CADERNO  A5 S/ARAME PEQUENO", offvalue="")
chk_caderno.grid(row=2, column=1, padx=5, pady=5, sticky="w")

chk_marcador_permanente = ttk.Checkbutton(frame_checkboxes, text="Marcador Permanente", variable=check_var_marcador_permante ,onvalue="PINCEL PILOTO PERMANENTE", offvalue="")
chk_marcador_permanente.grid(row=2, column=2, padx=5, pady=5, sticky="w")

chk_espanador = ttk.Checkbutton(frame_checkboxes, text="Espanador", variable=check_var_espanador, onvalue="ESPANADOR", offvalue="")
chk_espanador.grid(row=3, column=0, padx=5, pady=5, sticky="w")

check_var_outros = tk.IntVar()

chk_outros = ttk.Checkbutton(frame_checkboxes, text="Outros",)
chk_outros.grid(row=3, column=1, padx=5, pady=5, sticky="w")

entry_outros = ttk.Combobox(frame_checkboxes)

tk.Button(janela, text="Salvar", bg="blue", fg="black",command=salvar_pre_lista).grid(row=15, column=1, padx=1, pady=1, sticky="ew")
tk.Button(janela, text="Deletar", bg="blue", fg="black",command=deletar_pre_lista).grid(row=15, column=2, padx=1, pady=1, sticky="ew")

tree = ttk.Treeview(janela, columns=("ID","Nome funcionario", "Setor"), show='headings',)


tree.heading("ID", text="ID")
tree.heading("Nome funcionario", text="Nome do Produto",)
tree.heading("Setor", text="Setor")

tree.column("ID", width=50, anchor="center", stretch=True)
tree.column("Nome funcionario", width=150, stretch=True)
tree.column("Setor", width=50, anchor="center", stretch=True)

tree.bind("<Double-1>", lambda event: detalhes(event, tree, lista_pre_salva))

# Configurando o Treeview para preencher o espaço restante e ser responsivo
tree.grid(row=16, column=0, columnspan=5, padx=1, pady=1, sticky="nsew",)

tk.Button(janela, text="Finalizar", bg="blue", fg="black",command=finalizar).grid(row=17, column=1, padx=1, pady=1, sticky="ew")
tk.Button(janela, text="Cancela", bg="blue", fg="black",).grid(row=17, column=2, padx=1, pady=1, sticky="ew")

# Executa a janela
janela.mainloop()
