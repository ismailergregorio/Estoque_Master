import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from funcoes import*
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import subprocess

base_de_dados = "Base de dados/Cadastro Fornecedor.xlsx"

# Função de inicialização da interface
def deletar_fornecedor():
        # Obtém o item selecionado no Treeview
    wb = load_workbook(base_de_dados)
    ws = wb.active

    item_selecionado = tree.selection()
    if item_selecionado:  # Verifica se algum item foi selecionado
        # Deleta o item selecionado
        valores = tree.item(item_selecionado, "values")
        decisao = mostrar_alerta(f"Deseja Realmente excluir estes registro{valores[0],valores[1]}")
        try:
            if decisao == True:
                rows = list(enumerate(ws.iter_rows(min_row=2, values_only=True), start=2))

                # Itera sobre as linhas de forma reversa
                for index, row in reversed(rows):
                    # Compara o valor da célula (primeira coluna) com o valor do Treeview
                    print(row[0],valores[0])
                    if row[0] == valores[0]:  # Supondo que o identificador está na primeira coluna
                        ws.delete_rows(index)  # Deleta a linha correspondente
                        print(f"Item {valores[0]} deletado da planilha principal.")
                        tree.delete(item_selecionado)
                        wb.save(base_de_dados)
                        break  # Encerra o loop após encontrar a linha
        except:
            print("Erro ao Excluir")

def abrir_main():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "main.py"])
def abrir_janel_cadastro_fornrcedor():
    root.quit()  # Fecha a janela atual
    subprocess.Popen(["python", "Fornecedor\Janela_entrada_fornecedor.py"])

base_de_dados = "Base de dados\Cadastro Fornecedor.xlsx"
root = tk.Tk()
root.title("Sistema de Cadastro de Produtos")

# # Definindo o layout
# root.geometry("1000x500")
# root.configure(bg="white")

# Labels e entradas para o cadastro de produto
tk.Label(root, text="Cadastro de Fornecedor", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, columnspan=2)
tk.Button(root, text="Sair", bg="blue", fg="white",
            command=abrir_main).grid(row=0, column=6, padx=5, pady=5, sticky="ew")

tk.Button(root, text="Adicionar Fornecedor", bg="blue", fg="white",command=abrir_janel_cadastro_fornrcedor).grid(row=1, column=0, padx=5, pady=5, sticky="ew")


tk.Button(root, text="Deletar Fornecedor", bg="blue", fg="white",command=deletar_fornecedor).grid(row=1, column=1, padx=5, pady=5, sticky="ew")


tk.Button(root, text="Abri excel", bg="blue", fg="white").grid(row=1, column=2, padx=5, pady=5, sticky="ew")

tk.Button(root, text="Atulizar dados", bg="blue", fg="white",command=lambda:carregar_dados(base_de_dados,tree)).grid(row=1, column=12, padx=5, pady=5, sticky="ew")

tk.Label(root, text="Buscar produto", bg="blue", fg="white").grid(row=5, column=0, padx=5, pady=5, sticky="ew")
entry_buscar_produto = tk.Entry(root)
entry_buscar_produto.grid(row=6, column=0, padx=5, pady=5)
entry_buscar_produto.bind("<KeyRelease>")
entry_buscar_produto.bind("<KeyRelease>", converter_para_maiusculo)

# Filtros de data
tk.Label(root, text="Data inicial", bg="blue", fg="white").grid(row=7, column=0, padx=5, pady=5)
entry_data_inicial = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, year=2024)
entry_data_inicial.grid(row=8, column=0, padx=5, pady=5)

tk.Label(root, text="Data final", bg="blue", fg="white").grid(row=9, column=0, padx=5, pady=5)
entry_data_final = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2, year=2024)
entry_data_final.grid(row=10, column=0, padx=5, pady=5)

tk.Button(root, text="Filtrar", bg="blue", fg="white",command=lambda:filtro(tree,entry_data_inicial,entry_data_final,base_de_dados)).grid(row=11, column=0, padx=5, pady=5, sticky="ew")
tk.Button(root, text="Deletar filtro", bg="blue", fg="white",command=lambda:carregar_dados(base_de_dados,tree)).grid(row=12, column=0, padx=5, pady=5, sticky="ew")

# Tabela (Treeview) para exibição dos produtos cadastrados
colunas = ("CNPJ","nome_fornecedor", "Seguimento", "telefone", "e-mail","OBS")
tree = ttk.Treeview(root, columns=colunas, show="headings")

tree.heading("CNPJ", text="CNPJ")
tree.heading("nome_fornecedor", text="Nome do Fornecedor")
tree.heading("Seguimento", text="Seguimento")
tree.heading("telefone", text="Telefone")
tree.heading("e-mail", text="E-mail")
tree.heading("OBS", text="Obs")

tree.column("CNPJ", width=100)
tree.column("nome_fornecedor", width=100)
tree.column("Seguimento", width=100)
tree.column("telefone", width=100)
tree.column("e-mail", width=100)
tree.column("OBS", width=100)

tree.grid(row=3, column=1, columnspan=12, rowspan=21, padx=5, pady=5, sticky="nsew")

carregar_dados(base_de_dados,tree)

if __name__ == "__main__":
    root.mainloop()
