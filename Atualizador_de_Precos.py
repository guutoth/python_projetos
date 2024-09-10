import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import webbrowser
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

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
    preco = preco_element.get_text(strip=True).replace(
        "R$", "").strip() if preco_element else 'Preço não encontrado'

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
    novo_link = caixa_link.get().strip()
    if novo_link:
        arquivo_links = 'C:\\Atualizador de Preços (Quero-Quero)\\produtos.txt'
        caminho_arquivo = os.path.join(
            os.path.dirname(__file__), arquivo_links)

        with open(caminho_arquivo, 'a') as f:
            if not novo_link.endswith('\n'):
                novo_link += '\n'
            f.write(novo_link)

        caixa_link.delete(0, tk.END)
        atualizar_lista()
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
    arquivo_links = 'C:\\Atualizador de Preços (Quero-Quero)\\produtos.txt'
    urls_produtos = ler_links_arquivo(arquivo_links)

    for item in tree.get_children():
        tree.delete(item)

    if not urls_produtos:
        tree.insert('', 'end', values=(
            "Nenhum link de produto encontrado", "", ""))
        return

    # Usando ThreadPoolExecutor para realizar requisições em paralelo
    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_url = {executor.submit(
            obter_nome_e_preco, url): url for url in urls_produtos}
        for future in as_completed(future_to_url):
            try:
                nome, preco, link = future.result()
                tree.insert('', 'end', values=(nome, preco, link))
            except Exception as e:
                print(f"Erro ao obter dados para {future_to_url[future]}: {e}")

# Função para exportar para Excel


def exportar_para_excel():
    try:
        dados = []
        for item in tree.get_children():
            dados.append(tree.item(item, "values"))

        if not dados:
            messagebox.showwarning("Exportação falhou",
                                   "Não há dados para exportar.")
            return

        df = pd.DataFrame(dados, columns=["Produto", "Preço", "Link"])

        caminho_excel = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Salvar como"
        )

        if not caminho_excel:
            return

        df.to_excel(caminho_excel, index=False)
        messagebox.showinfo("Exportação concluída",
                            f"Dados exportados para {caminho_excel}")
    except Exception as e:
        messagebox.showerror("Erro na exportação",
                             f"Ocorreu um erro ao exportar: {e}")

# Função para excluir item selecionado


def excluir_item():
    selecionado = tree.selection()
    if not selecionado:
        messagebox.showwarning("Nenhum item selecionado",
                               "Selecione um item para excluir.")
        return

    for item in selecionado:
        link = tree.item(item, "values")[2]
        excluir_do_arquivo(link)
        tree.delete(item)
    atualizar_lista()

# Função para excluir link do arquivo produtos.txt


def excluir_do_arquivo(link):
    caminho_arquivo = 'C:\\Atualizador de Preços (Quero-Quero)\\produtos.txt'
    with open(caminho_arquivo, 'r') as f:
        linhas = f.readlines()

    with open(caminho_arquivo, 'w') as f:
        for linha in linhas:
            if linha.strip() != link:
                f.write(linha)


# Criação da interface gráfica
janela = tk.Tk()
janela.title("Lista de Produtos e Preços - Versão 1.1")

style = ttk.Style()
style.configure("Treeview.Heading", font=("Helvetica", 14))
style.configure("Treeview", rowheight=30, font=("Helvetica", 12))


def centralizar_colunas():
    tree.column('Produto', anchor='center')
    tree.column('Preço', anchor='center')
    tree.column('Link', anchor='center')


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

frame_adicionar = tk.Frame(janela)
frame_adicionar.pack(pady=10)

tk.Label(frame_adicionar, text="Adicionar Link:",
         font=("Helvetica", 12)).pack(side=tk.LEFT)

caixa_link = tk.Entry(frame_adicionar, width=50, font=("Helvetica", 12))
caixa_link.pack(side=tk.LEFT, padx=5)

botao_adicionar = tk.Button(
    frame_adicionar, text="Adicionar", command=adicionar_link, font=("Helvetica", 12))
botao_adicionar.pack(side=tk.LEFT)

botao_atualizar = tk.Button(
    janela, text="Atualizar Preços", command=atualizar_lista, font=("Helvetica", 12))
botao_atualizar.pack(pady=5)

botao_exportar = tk.Button(janela, text="Exportar para Excel",
                           command=exportar_para_excel, font=("Helvetica", 12))
botao_exportar.pack(pady=5)

botao_excluir = tk.Button(janela, text="Excluir Selecionado",
                          command=excluir_item, font=("Helvetica", 12))
botao_excluir.pack(pady=5)

centralizar_colunas()

janela.geometry("800x600")
janela.mainloop()
