import tkinter as tk
from PIL import Image, ImageTk  # Biblioteca PIL para manipular imagens

def Bunton_custo(caminho_imagem,root,text,comando):
    # Carregar a imagem (certifique-se de que a imagem está no mesmo diretório ou forneça o caminho completo)
    imagem_original = Image.open(caminho_imagem)  # Substitua por sua imagem
    imagem_redimensionada = imagem_original.resize((20, 20))  # Ajuste o tamanho da imagem
    icone = ImageTk.PhotoImage(imagem_redimensionada)

    # Criar botão com ícone e texto
    botao = tk.Button(root, text=text, image=icone, compound="left", command=comando)
    return botao

def criar_botao_com_icone(root, imagem_caminho, texto, comando=None):
    # Carregar e redimensionar a imagem
    imagem_original = Image.open(imagem_caminho)
    imagem_redimensionada = imagem_original.resize((50, 50))  # Ajuste o tamanho conforme necessário
    icone = ImageTk.PhotoImage(imagem_redimensionada)

    # Criar o botão com imagem e texto
    botao = tk.Button(root, text=texto, image=icone, compound="left", command=comando,width=200,height=100)
    botao.image = icone  # Referência para evitar garbage collection
    return botao

