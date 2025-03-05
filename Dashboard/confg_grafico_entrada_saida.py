import tkinter as tk
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


def gera_grafico(root,dados_x,dados_y,):
    global ax,fig ,canvas
    
    fig = Figure(figsize=(8, 3), dpi=100)
    ax = fig.add_subplot(111)

    ax.clear()  # Limpa qualquer gráfico anterior
    bars = ax.bar(dados_x, dados_y)  # Cria as barras

    for bar in bars:
        height = bar.get_height()  # Obtém o valor da barra
        ax.text(bar.get_x() + bar.get_width()/2, height, f'{height}', 
                ha='center', va='bottom', fontsize=10)

    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.get_tk_widget().pack()
    canvas.draw()

def atualiza_grafico(dados_x,dados_y):
    global ax,fig ,canvas
    ax.clear()
    bars = ax.bar(dados_x, dados_y)  # Cria as barras

    for bar in bars:
        height = bar.get_height()  # Obtém o valor da barra
        ax.text(bar.get_x() + bar.get_width()/2, height, f'{height}', 
                ha='center', va='bottom', fontsize=10)

    canvas.draw()  # Atualiza a exibição do gráfico no Tkinter


def gera_grafico_vertical1(root,grupos,valores):
    global ax_horizontal,fig_horizontal ,canvas_horizontal

    y_pos = np.arange(len(grupos))  # Posições no eixo Y

    # Criar o gráfico
    fig_horizontal, ax_horizontal = plt.subplots(figsize=(4, 3))  # Ajuste de tamanho
    fig_horizontal.subplots_adjust(left=0.3, right=0.9, top=0.9, bottom=0.2)

    # Criar o gráfico de barras horizontais
    bars = ax_horizontal.barh(y_pos, valores, color='skyblue', edgecolor='black')

    # Adicionar os valores ao lado das barras
    for bar, valor in zip(bars, valores):
        ax_horizontal.text(valor + 0.3, bar.get_y() + bar.get_height()/2,  # Posição
                        str(valor), va='center', ha='left', color='black', fontsize=10)

    # Configurar os rótulos e título
    ax_horizontal.set_yticks(y_pos)
    ax_horizontal.set_yticklabels(grupos)
    ax_horizontal.set_xlabel('Valores')
    ax_horizontal.set_title('Gráfico de Barras Horizontais')

    # Integrar o gráfico ao Tkinter
    canvas_horizontal = FigureCanvasTkAgg(fig_horizontal, master=root)
    canvas_horizontal.get_tk_widget().pack(expand=True, fill=tk.BOTH)
    canvas_horizontal.draw()
    
def atuliza_grafico_vertical1(root,grupos,valores):
    global ax_horizontal,fig_horizontal ,canvas_horizontal
    ax_horizontal.clear()

    y_pos = np.arange(len(grupos))  # Posições no eixo Y

    bars = ax_horizontal.barh(y_pos, valores, color='skyblue', edgecolor='black')

    # Adicionar os valores ao lado das barras
    for bar, valor in zip(bars, valores):
        ax_horizontal.text(valor + 0.3, bar.get_y() + bar.get_height()/2,  # Posição
                        str(valor), va='center', ha='left', color='black', fontsize=10)
    
    ax_horizontal.set_yticks(y_pos)
    ax_horizontal.set_yticklabels(grupos)
    ax_horizontal.set_xlabel('Valores')
    ax_horizontal.set_title('Gráfico de Barras Horizontais')

    canvas_horizontal.draw()  # Atualiza a exibição do gráfico no Tkinter


def gera_grafico_vertical2(root,grupos,valores):
    global ax_horizontal2,fig_horizontal2 ,canvas_horizontal2

    y_pos = np.arange(len(grupos))  # Posições no eixo Y

    # Criar o gráfico
    fig_horizontal2, ax_horizontal2 = plt.subplots(figsize=(4, 3))  # Ajuste de tamanho
    fig_horizontal2.subplots_adjust(left=0.3, right=0.9, top=0.9, bottom=0.2)

    # Criar o gráfico de barras horizontais
    bars = ax_horizontal2.barh(y_pos, valores, color='skyblue', edgecolor='black')

    # Adicionar os valores ao lado das barras
    for bar, valor in zip(bars, valores):
        ax_horizontal2.text(valor + 0.3, bar.get_y() + bar.get_height()/2,  # Posição
                        str(valor), va='center', ha='left', color='black', fontsize=10)

    # Configurar os rótulos e título
    ax_horizontal2.set_yticks(y_pos)
    ax_horizontal2.set_yticklabels(grupos)
    ax_horizontal2.set_xlabel('Valores')
    ax_horizontal2.set_title('Gráfico de Barras Horizontais')

    # Integrar o gráfico ao Tkinter
    canvas_horizontal2 = FigureCanvasTkAgg(fig_horizontal2, master=root)
    canvas_horizontal2.get_tk_widget().pack(expand=True, fill=tk.BOTH)
    canvas_horizontal2.draw()

def atuliza_grafico_vertical2(root,grupos,valores):
    global ax_horizontal2,fig_horizontal2 ,canvas_horizontal2
    ax_horizontal2.clear()

    y_pos = np.arange(len(grupos))  # Posições no eixo Y

    bars = ax_horizontal2.barh(y_pos, valores, color='skyblue', edgecolor='black')

    # Adicionar os valores ao lado das barras
    for bar, valor in zip(bars, valores):
        ax_horizontal2.text(valor + 0.3, bar.get_y() + bar.get_height()/2,  # Posição
                        str(valor), va='center', ha='left', color='black', fontsize=10)
    
    ax_horizontal2.set_yticks(y_pos)
    ax_horizontal2.set_yticklabels(grupos)
    ax_horizontal2.set_xlabel('Valores')
    ax_horizontal2.set_title('Gráfico de Barras Horizontais')

    canvas_horizontal2.draw()  # Atualiza a exibição do gráfico no Tkinter