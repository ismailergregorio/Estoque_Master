import sys
import os

# Adiciona o diretório do arquivo de funções ao sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from funcoes import*

base_de_dados_cadastro = 'Base de dados\CADASTRO.xlsx'

base_de_dados = 'Base de dados\entrada.xlsx'

dados_de_estoque = "Base de dados\Estoque.xlsx"


def carregar_dados_entry(arquivo, indice):
    """
    Carrega dados de uma planilha e retorna uma lista contendo os valores exclusivos
    de uma coluna específica, identificada pelo índice fornecido.
    
    :param arquivo: Caminho para o arquivo (normalmente Excel) que será aberto.
    :param indice: Índice da coluna cujos valores serão extraídos.
    :return: Lista de valores únicos da coluna especificada ou uma mensagem de erro.
    """
    # Abre o arquivo fornecido e carrega os dados
    dados = abrir_arquivo(arquivo)  # Presume que abrir_arquivo é uma função previamente definida.

    # Inicializa a lista que armazenará os valores da coluna especificada
    lista_dados = []
    try:
        # Itera sobre as linhas da planilha, ignorando o cabeçalho (a primeira linha)
        for x in dados.iter_rows(min_row=2, values_only=True):
            # Verifica se o valor da coluna ainda não está na lista; se não, adiciona
            if x[indice] not in lista_dados:
                lista_dados.append(x[indice])
    except:
        # Em caso de erro (como arquivo inválido ou índice fora do alcance), retorna mensagem de erro
        lista_dados = "Dados não Disponível" 

    # Retorna a lista com os dados únicos ou a mensagem de erro
    return lista_dados


def configurar_busca_combobox(combobox, lista_valores):
    """
    Configura o comportamento de busca dinâmica para um Combobox.
    
    :param combobox: O Combobox a ser configurado.
    :param lista_valores: A lista original de valores do Combobox.
    :param delay: Tempo de espera (em milissegundos) após a digitação antes de filtrar.
    """

    delay=600
    filtro_id = None  # Armazena o ID do after para evitar execuções desnecessárias

    def atualizar_lista(event=None):
        nonlocal filtro_id
        if filtro_id:
            combobox.after_cancel(filtro_id)  # Cancela o último after programado
        
        # Aguarda um pouco para processar a entrada
        filtro_id = combobox.after(delay, lambda: filtrar_lista(combobox, lista_valores))

    def filtrar_lista(combobox, lista_valores):
        entrada = combobox.get().lower()
        if not entrada:  # Se a entrada estiver vazia, restaura a lista original
            combobox['values'] = lista_valores
            combobox.event_generate('<Down>')
        else:  # Filtra a lista com base na entrada
            filtrado = [item for item in lista_valores if entrada in str(item).lower()]
            combobox['values'] = filtrado

         # Reabre o Combobox automaticamente (opcional)
            combobox.event_generate('<Down>')

    # Adiciona o evento de keyrelease ao Combobox
    combobox.bind("<KeyRelease>", atualizar_lista)



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
                    linha[5].value = float(linha[5].value) + float(iten[4])  # Atualiza o valor na coluna 5 (estoque)
                    linha[6].value = float(linha[6].value) - float(iten[4])
                    linha[7].value = float(linha[7].value) - float(iten[8])
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

def deletar_item_saida(freme_da_tabela,dados):
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
        try:
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
                        linha[5].value = float(linha[5].value) - float(valores[4])  # Atualiza o valor na coluna 5 (estoque)
                        linha[6].value = float(linha[6].value) + float(valores[4])
                        linha[7].value = float(linha[7].value) + float(valores[8])
                        valores_linha = [celula.value for celula in linha]  # Extrai os valores de todas as células na linha
                        workbook_estoque.save(dados_de_estoque)
                        break
            else:
                print("Operação cancelada pelo usuário.")
        except:
            alerta_erro("Erro ao Deletar","Verifique os dados")
    else:
        print("Nenhum item selecionado para deletar.")

        

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
