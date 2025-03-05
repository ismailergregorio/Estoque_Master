import tkinter as tk
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
