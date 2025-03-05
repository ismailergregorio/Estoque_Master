
from tkinter import*
import subprocess
from style import*


base_de_dados_saida = "Base de dados/saida.xlsx"
base_de_dados_entrada = "Base de dados/entrada.xlsx"
base_de_dados_casdastro = "Base de dados/CADASTRO.xlsx"
base_de_dados_casdastro_nf = "Base de dados/Cadastro NF.xlsx"

def abrir_janela(tipo_de_janela):
    if tipo_de_janela == 'Cadastro_pack\Cadastro.py': 
        subprocess.Popen(["python", tipo_de_janela])  # Abre a Página 2
    elif tipo_de_janela == 'Entrada_pack\Entradai.py':
        subprocess.Popen(["python", tipo_de_janela])  # Abre a Página 2
    elif tipo_de_janela == 'Saida_pack/Saidai.py':
        subprocess.Popen(["python", tipo_de_janela])  # Abre a Página 2
    elif tipo_de_janela == 'Estoque_pack\Estoque.py':
        subprocess.Popen(["python", tipo_de_janela])
    elif tipo_de_janela == 'Fornecedor\Cadastro Fornecedor.py':
        subprocess.Popen(["python", tipo_de_janela])
    elif tipo_de_janela == 'NF\Cadastro_NF.py':
        subprocess.Popen(["python", tipo_de_janela])
    elif tipo_de_janela == 'Dashboard\Dashboard_NF.py':
        subprocess.Popen(["python", tipo_de_janela])
    else:
        print('pagina não encontrada')
    

# Criando a janela principal do Tkinter
root = Tk()
root.title("Gráfico no Tkinter")
root.geometry("1200x900")
root.columnconfigure(0, weight=10)  # Permite expandir a coluna
root.rowconfigure(0, weight=10)    # Permite expandir a linha
# Criando um frame para os outros widgets
frame_widgets = Frame(root)
frame_widgets.grid(column=0, row=0 ,columnspan=4)

b_cadastro = criar_botao_com_icone(frame_widgets, "stilo/cadastre-se.png", "Cadastro Itens",lambda:abrir_janela("Cadastro_pack\Cadastro.py")).grid(row=0, column=0, padx=5, pady=5, sticky='w')
b_entrada = criar_botao_com_icone(frame_widgets, "stilo/Entrada.png", "Entrada Produtos", lambda:abrir_janela("Entrada_pack\Entradai.py")).grid(row=0, column=1, padx=5, pady=5, sticky='w')
b_saida = criar_botao_com_icone(frame_widgets, "stilo/Saida.png", "Saida Produtos", lambda:abrir_janela("Saida_pack/Saidai.py")).grid(row=0, column=2, padx=5, pady=5, sticky='w')
b_estoque = criar_botao_com_icone(frame_widgets, "stilo/estoque.png", "Estoque Produtos", lambda:abrir_janela("Estoque_pack\Estoque.py")).grid(row=0, column=3, padx=5, pady=5, sticky='w')
b_relatorio = criar_botao_com_icone(frame_widgets, "stilo/estoque.png", "Relatorios", lambda:abrir_janela("-")).grid(row=1, column=3, padx=5, pady=5, sticky='w')
b_cadastro_fornecedor = criar_botao_com_icone(frame_widgets, "stilo/entregador.png", "Cadastro Fornecedores", lambda:abrir_janela("Fornecedor\Cadastro Fornecedor.py")).grid(row=1, column=0, padx=5, pady=5, sticky='w')
b_cadastro_nf = criar_botao_com_icone(frame_widgets, "stilo/fatura.png", "Cadastro NF", lambda:abrir_janela("NF\Cadastro_NF.py")).grid(row=1, column=1, padx=5, pady=5, sticky='w')
b_deshibord = criar_botao_com_icone(frame_widgets, "stilo/relatorio.png", "Deshibord", lambda:abrir_janela("Dashboard\Dashboard_NF.py")).grid(row=1, column=2, padx=5, pady=5, sticky='w')

root.mainloop()