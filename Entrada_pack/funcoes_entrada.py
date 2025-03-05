import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from funcoes import*

base_de_dados_cadastro = 'Base de dados\CADASTRO.xlsx'

base_de_dados = 'Base de dados\entrada.xlsx'

dados_de_estoque = "Base de dados\Estoque.xlsx"



def atualizar_campos1(dados, entry,set1,set2,set3):
    """
    Função para atualizar o campo baseado no valor digitado ou alterado.
    - dados: Caminho para o arquivo ou base de dados.
    - entry1: Combobox ou Entry widget a ser monitorado.
    """
    # Carrega os dados do arquivo
    base_dados = abrir_arquivo(dados)

        # Obtém o valor atual do widget
    valor = entry.get().strip()
    # Verifica se o valor está presente na base de dados
    for i in base_dados.iter_rows(min_row=2, values_only=True):
        if valor == str(i[1]):  # Comparação com a coluna de índice 0
            entry.delete(0, tk.END)
            # Aqui você pode adicionar lógica para atualizar outros campos
            set1.set(i[0])
            set2.set(i[1])
            set3.set(i[2])
            break
    
def calcular_media(codigo_iten):
    dados_m = abrir_arquivo(dados_de_entrada)

    quantidade = 0
    soma_do_valor = 0
    valor_media_final = 0
    try:
        for itens in dados_m.iter_rows(min_row=2, values_only=True):
            if itens[0] == codigo_iten[0]:
                quantidade = quantidade + int(itens[4])
                soma_do_valor = soma_do_valor + float(itens[7])

                valor_media_final = soma_do_valor / quantidade
    except:
        print("Erro ao calcular o valor Medio") 

    print(valor_media_final)
    try:
        return float(f"{valor_media_final:.3}")
    except:
        return float(valor_media_final)



def finalizar_itens_entrada(lista,dados):
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

                    linha[4].value = float(linha[4].value) + float(iten[4])  # Atualiza o valor na coluna 5 (estoque)
                    linha[6].value = float(linha[6].value) + float(iten[4])
                    linha[7].value = float(linha[7].value) + float(iten[7])
                    linha[8].value = calcular_media(iten) # ajusta o valor medio 

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

def deletar_item_entrada(freme_da_tabela,dados):
    # Obtém o item selecionado no Treeview
    wb = load_workbook(dados)
    ws = wb.active

    workbook_estoque = load_workbook(dados_de_estoque)
    sheet_estoque = workbook_estoque.active  # Seleciona a planilha ativa

    item_selecionado = freme_da_tabela.selection()
    
    if item_selecionado:  # Verifica se algum item foi selecionado
        # Deleta o item selecionado
        valores = freme_da_tabela.item(item_selecionado, "values")
        decisao = mostrar_alerta(f"Deseja Realmente excluir estes registro{valores[0],valores[1]}")
        # try:
        if decisao == True:
            rows = list(enumerate(ws.iter_rows(min_row=2, values_only=True), start=2))

            # Itera sobre as linhas de forma reversa
            for index, row in reversed(rows):
                # Compara o valor da célula (primeira coluna) com o valor do Treeview
                if int(row[0]) == int(valores[0]):  # Supondo que o identificador está na primeira coluna
                    ws.delete_rows(index)  # Deleta a linha correspondente
                    print(f"Item {valores[0]} deletado da planilha principal.")
                    wb.save(dados)
                    break  # Encerra o loop após encontrar a linha
            
            for linha in sheet_estoque.iter_rows(min_row=2,values_only=False):
                if int(linha[0].value) == int(valores[0]):  # Verifica o ID do item
                    linha[4].value = float(linha[4].value) - float(valores[4])  # Atualiza o valor na coluna 5 (estoque)
                    linha[6].value = float(linha[6].value) - float(valores[4])
                    linha[7].value = float(linha[7].value) - float(valores[7])
                    linha[8].value = calcular_media(valores)
                    # print(linha[8].value)
                    valores_linha = [celula.value for celula in linha]  # Extrai os valores de todas as células na linha
                    workbook_estoque.save(dados_de_estoque)
                    break
        else:
            print("Operação cancelada pelo usuário.")
        # except Exception as e:
        #     alerta_erro(f"Erro ao Deletar{e}",f"Verifique os dados.")
    else:
        print("Nenhum item selecionado para deletar.")

        

def adicionar_itens_na_lista(tabela,codigo,nome_produto, setor,data,quantidade,tipo_entrada,valor_unitario,valor_total,obs):

    maior = 0

    for item_id in tabela.get_children():  # Itera pelos IDs das linhas
        valores = tabela.item(item_id, "values")  # Obtém os valores da linha
        maior == int(valores[0])
        if int(valores[0]) > maior:
            maior = int(valores[0])
    
    maior = maior + 1

    # print(maior, codigo.get(),nome_produto.get(), setor.get(),data.get_date(),quantidade.get(),tipo_entrada.get(),valor_unitario.get(),valor_total.get(),obs.get())

    tabela.insert("", "end", values=(maior, codigo.get(),nome_produto.get(), setor.get(),data.get_date(),quantidade.get(),tipo_entrada.get(),valor_unitario.get(),valor_total.get(),obs.get()))
    itens = [maior, codigo.get(),nome_produto.get(), setor.get(),data.get_date(),quantidade.get(),tipo_entrada.get(),valor_unitario.get(),valor_total.get(),obs.get()]
    
    return itens



