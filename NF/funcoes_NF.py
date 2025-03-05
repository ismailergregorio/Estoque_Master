import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from funcoes import*

base_de_dados_cadastro = 'Base de dados\CADASTRO.xlsx'

base_de_dados = 'Base de dados\entrada.xlsx'

dados_de_estoque = "Base de dados\Estoque.xlsx"

dados_de_nf = "Base de dados\Cadastro NF.xlsx"

def calcular_media(codigo_iten):
    """
    Função para calcular a média do valor dos itens com base nos registros na planilha de entrada.
    
    Parâmetros:
    - codigo_iten: Código do item a ser analisado.
    - dados_de_entrada: Caminho do arquivo Excel com os dados de entrada.
    - abrir_arquivo: Função para carregar o arquivo Excel.
    
    Retorna:
    - Média do valor do item formatada para duas casas decimais.
    """
    
    # Carrega os dados da planilha de entrada
    dados_m = abrir_arquivo(dados_de_entrada)
    
    # Inicializa variáveis para cálculo
    quantidade = 0
    soma_do_valor = 0
    valor_media_final = 0
    
    try:
        # Itera sobre as linhas da planilha a partir da segunda linha
        for itens in dados_m.iter_rows(min_row=2, values_only=True):
            if itens[0] == codigo_iten[0]:  # Verifica se o código do item corresponde
                quantidade += int(itens[4])  # Soma a quantidade total
                soma_do_valor += float(itens[7])  # Soma os valores totais
                
                # Calcula a média do valor do item
                valor_media_final = soma_do_valor / quantidade
    except Exception as e:
        print(f"Erro ao calcular o valor médio: {e}")  # Exibe uma mensagem de erro em caso de falha
    
    # Retorna o valor médio formatado para duas casas decimais
    return float(f"{valor_media_final:.2f}")


def deletar_item_entrada(freme_da_tabela,dados):
    """
    Função para deletar um item selecionado em um Treeview e remover as correspondentes linhas de registros 
    em diferentes planilhas do Excel, atualizando os valores no estoque.
    """
    
    # Carrega os arquivos Excel
    wb = load_workbook(dados_de_nf)
    ws = wb.active  # Seleciona a planilha ativa

    wbe = load_workbook(dados_de_entrada)
    sheet_entrada = wbe.active  # Seleciona a planilha ativa

    workbook_estoque = load_workbook(dados_de_estoque)
    sheet_estoque = workbook_estoque.active  # Seleciona a planilha ativa

    # Obtém o item selecionado no Treeview
    item_selecionado = freme_da_tabela.selection()
    
    if item_selecionado:  # Verifica se algum item foi selecionado
        item_selecionado = item_selecionado[0]  # Obtém o primeiro item selecionado
        valores = freme_da_tabela.item(item_selecionado, "values")  # Obtém os valores do item selecionado
        
        # Exibe um alerta de confirmação
        decisao = mostrar_alerta(f"Deseja realmente excluir este registro? NF: {valores[4]}, Fornecedor: {valores[1]}")
        
        if decisao:  # Se o usuário confirmar a exclusão
            
            try:
                # Deleta a linha correspondente na planilha de NF
                rows = list(enumerate(ws.iter_rows(min_row=2, values_only=True), start=2))
                for index, row in reversed(rows):  # Itera sobre as linhas de forma reversa
                    if int(row[4]) == int(valores[4]):  # Compara o número da NF com o da tabela
                        ws.delete_rows(index)  # Deleta a linha correspondente
                        freme_da_tabela.delete(item_selecionado)  # Remove o item da Treeview
                        wb.save(dados_de_nf)  # Salva a alteração no arquivo Excel
                        break  # Encerra o loop após encontrar e deletar a linha

            except Exception as e:
                alerta_erro("Erro ao Deletar NF", f"Verifique os dados. Erro: {e}")

            try:
                # Deleta a linha correspondente na planilha de entrada
                rows = list(enumerate(sheet_entrada.iter_rows(min_row=2, values_only=True), start=2))
                for index, linha_E in reversed(rows):
                    if linha_E[8] == f"NF° {valores[4]},For:{valores[1]}":  # Verifica se a entrada corresponde à NF
                        
                        # Atualiza os valores na planilha de estoque
                        for linha_estoque in sheet_estoque.iter_rows(min_row=2, values_only=False):
                            if int(linha_estoque[0].value) == int(linha_E[0]):  # Compara o ID do item
                                
                                # Atualiza os valores das colunas relacionadas ao estoque
                                linha_estoque[4].value = float(linha_estoque[4].value) - float(linha_E[4])  # Diminui a quantidade no estoque
                                linha_estoque[6].value = float(linha_estoque[6].value) - float(linha_E[4])
                                linha_estoque[7].value = float(linha_estoque[7].value) - float(linha_E[7])
                                linha_estoque[8].value = calcular_media(linha_E)  # Recalcula a média de valores
                                
                                # Salva a atualização no estoque
                                workbook_estoque.save(dados_de_estoque)
                                
                        # Deleta a linha correspondente na planilha de entrada
                        sheet_entrada.delete_rows(index)
                        wbe.save(dados_de_entrada)  # Salva a alteração na planilha de entrada
                        
            except Exception as e:
                alerta_erro("Erro ao Deletar Atulisar o estoque", f"Verifique os dados. Erro: {e}")
        
        else:
            print("Operação cancelada pelo usuário.")
    else:
        print("Nenhum item selecionado para deletar.")
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
