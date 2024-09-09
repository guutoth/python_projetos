### ATUALIZAÇÃO 09/09/2024 - 10:00 ####

import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, messagebox
import os
import webbrowser
import pandas as pd

# Função para obter nome e preço do produto


def obter_nome_e_preco(url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
    except requests.RequestException as e:
        return 'Nome não encontrado', 'Preço não encontrado', url

    soup = BeautifulSoup(response.content, 'html.parser')
    nome_element = soup.select_one(
        '.vtex-store-components-3-x-productBrand--quickview')
    preco_element = soup.select_one(
        '.vtex-product-price-1-x-currencyContainer--product-price')

    nome = nome_element.get_text(
        strip=True) if nome_element else 'Nome não encontrado'
    preco = preco_element.get_text(
        strip=True) if preco_element else 'Preço não encontrado'

    return nome, preco, url

# Função para ler links do arquivo


def ler_links_arquivo(arquivo):
    caminho_arquivo = os.path.join(os.path.dirname(__file__), arquivo)
    if not os.path.isfile(caminho_arquivo):
        return []
    with open(caminho_arquivo, 'r') as f:
        urls = f.read().splitlines()
    return urls

# Função para adicionar link ao arquivo


def adicionar_link():
    novo_link = caixa_link.get().strip()  # Remove espaços em branco extras
    if novo_link:
        arquivo_links = 'produtos.txt'
        caminho_arquivo = os.path.join(
            os.path.dirname(__file__), arquivo_links)

        # Adiciona quebra de linha se não houver
        with open(caminho_arquivo, 'a') as f:
            if not novo_link.endswith('\n'):
                novo_link += '\n'
            f.write(novo_link)

        caixa_link.delete(0, tk.END)  # Limpa a caixa de texto após adicionar
        atualizar_lista()  # Atualiza a lista com o novo link
    else:
        messagebox.showwarning(
            "Entrada inválida", "Por favor, insira um link válido.")

# Função para abrir o link no navegador


def abrir_link(event):
    item = tree.selection()[0]
    link = tree.item(item, "values")[2]
    webbrowser.open(link)

# Função para atualizar a lista de produtos


def atualizar_lista():
    arquivo_links = 'produtos.txt'
    urls_produtos = ler_links_arquivo(arquivo_links)

    for item in tree.get_children():
        tree.delete(item)

    if not urls_produtos:
        tree.insert('', 'end', values=(
            "Nenhum link de produto encontrado", "", ""))
        return

    for url in urls_produtos:
        nome, preco, link = obter_nome_e_preco(url)
        tree.insert('', 'end', values=(nome, preco, link))

# Função para exportar para Excel


def exportar_para_excel():
    dados = []
    for item in tree.get_children():
        dados.append(tree.item(item, "values"))

    df = pd.DataFrame(dados, columns=["Produto", "Preço", "Link"])
    df.to_excel("produtos_precos.xlsx", index=False)
    messagebox.showinfo("Exportação concluída",
                        "Dados exportados para produtos_precos.xlsx")


# Criação da interface gráfica
janela = tk.Tk()
janela.title("Lista de Produtos e Preços")

style = ttk.Style()
style.configure("Treeview.Heading", font=("Helvetica", 14))
style.configure("Treeview", rowheight=30, font=("Helvetica", 12))

# Função para centralizar colunas


def centralizar_colunas():
    tree.column('Produto', anchor='center')
    tree.column('Preço', anchor='center')
    tree.column('Link', anchor='center')


# Definição das colunas e criação da Treeview
colunas = ('Produto', 'Preço', 'Link')
tree = ttk.Treeview(janela, columns=colunas, show='headings', style="Treeview")

tree.heading('Produto', text='Produto')
tree.heading('Preço', text='Preço')
tree.heading('Link', text='Link')

tree.column('Produto', width=250, anchor='center')
tree.column('Preço', width=75, anchor='center')
tree.column('Link', width=350, anchor='center')

tree.pack(fill=tk.BOTH, expand=True)

tree.bind("<Double-1>", abrir_link)

# Adicionar Link
frame_adicionar = tk.Frame(janela)
frame_adicionar.pack(pady=10)

tk.Label(frame_adicionar, text="Adicionar Link:",
         font=("Helvetica", 12)).pack(side=tk.LEFT)

caixa_link = tk.Entry(frame_adicionar, width=50, font=("Helvetica", 12))
caixa_link.pack(side=tk.LEFT, padx=5)

botao_adicionar = tk.Button(
    frame_adicionar, text="Adicionar", command=adicionar_link, font=("Helvetica", 12))
botao_adicionar.pack(side=tk.LEFT)

# Botões de Atualizar e Exportar
botao_atualizar = tk.Button(
    janela, text="Atualizar Preços", command=atualizar_lista, font=("Helvetica", 12))
botao_atualizar.pack(pady=5)

botao_exportar = tk.Button(janela, text="Exportar para Excel",
                           command=exportar_para_excel, font=("Helvetica", 12))
botao_exportar.pack(pady=5)

centralizar_colunas()

# Ajusta a resolução para 800x600
janela.geometry("800x600")
janela.mainloop()
