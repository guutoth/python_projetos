### VERSÃO 2.3 ###
# UPDATES:
# 1 - REFORMULAÇÃO DAS BARRAS E BOTÕES
# 2 - OPÇÃO DE ADICIONAR PRODUTOS DO MERCADO LIVRE E LOJAS QUERO QUERO

import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import webbrowser
import pandas as pd
import concurrent.futures
import shutil
from datetime import datetime
import time

# Constantes de configuração
CAMINHO_ARQUIVO = 'C:\\Atualizador de Preços (ML e Quero-Quero)\\produtos.txt'
DIRETORIO_BACKUP = 'C:\\Atualizador de Preços (ML e Quero-Quero)\\backup\\'
COLUNAS = ('Produto', 'Preço', 'Link')

# Funções específicas para diferentes sites


def obter_nome_e_preco_queroquero(url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers, timeout=20)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Erro ao acessar {url}: {e}")
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


def obter_nome_e_preco_mercadolivre(url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers, timeout=20)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Erro ao acessar {url}: {e}")
        return 'Nome não encontrado', 'Preço não encontrado', url

    soup = BeautifulSoup(response.content, 'html.parser')
    nome_element = soup.select_one('h1.ui-pdp-title')
    preco_element = soup.select_one('.andes-money-amount__fraction')

    nome = nome_element.get_text(
        strip=True) if nome_element else 'Nome não encontrado'
    preco = preco_element.get_text(strip=True).replace(
        "R$", "").strip() if preco_element else 'Preço não encontrado'

    return nome, preco, url


# Função que escolhe a função de parsing com base no domínio
def obter_nome_e_preco(url):
    if "queroquero.com" in url:
        return obter_nome_e_preco_queroquero(url)
    elif "mercadolivre.com" in url:
        return obter_nome_e_preco_mercadolivre(url)
    else:
        return 'Site não suportado', 'Preço não encontrado', url


def ler_links_arquivo(arquivo):
    caminho_arquivo = os.path.join(os.path.dirname(__file__), arquivo)
    if not os.path.isfile(caminho_arquivo):
        return []
    with open(caminho_arquivo, 'r') as f:
        return f.read().splitlines()


def fazer_backup(caminho_arquivo):
    if not os.path.exists(DIRETORIO_BACKUP):
        os.makedirs(DIRETORIO_BACKUP, exist_ok=True)
        print(f"Pasta de backup criada em: {DIRETORIO_BACKUP}")

    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    backup_arquivo = os.path.join(
        DIRETORIO_BACKUP, f'produtos_backup_{timestamp}.txt')
    try:
        shutil.copy(caminho_arquivo, backup_arquivo)
        print(f"Backup criado em: {backup_arquivo}")
    except Exception as e:
        print(f"Erro ao criar o backup: {e}")


def adicionar_link():
    novo_link = caixa_link.get().strip()
    if novo_link:
        if verificar_link_existente(novo_link):
            messagebox.showwarning("Produto duplicado",
                                   "O produto já está na lista.")
            return

        fazer_backup(CAMINHO_ARQUIVO)
        salvar_link(novo_link)
        caixa_link.delete(0, tk.END)
        atualizar_lista()
    else:
        messagebox.showwarning(
            "Entrada inválida", "Por favor, insira um link válido.")


def verificar_link_existente(link):
    urls_existentes = ler_links_arquivo(CAMINHO_ARQUIVO)
    return link in urls_existentes


def salvar_link(link):
    with open(CAMINHO_ARQUIVO, 'a') as f:
        if not link.endswith('\n'):
            link += '\n'
        f.write(link)


def excluir_do_arquivo(link):
    fazer_backup(CAMINHO_ARQUIVO)
    linhas = ler_arquivo(CAMINHO_ARQUIVO)
    salvar_linhas(linhas, link)


def ler_arquivo(caminho_arquivo):
    with open(caminho_arquivo, 'r') as f:
        return f.readlines()


def salvar_linhas(linhas, link):
    with open(CAMINHO_ARQUIVO, 'w') as f:
        for linha in linhas:
            if linha.strip() != link:
                f.write(linha)


def abrir_link(event):
    item = tree.selection()[0]
    link = tree.item(item, "values")[2]
    webbrowser.open(link)


def atualizar_lista(ordenar_por=None):
    global progress_bar, status_label
    urls_produtos = ler_links_arquivo(CAMINHO_ARQUIVO)
    tree.delete(*tree.get_children())
    progress_bar['value'] = 0
    progress_bar['maximum'] = len(urls_produtos)
    status_label.config(text="Atualizando...")

    if not urls_produtos:
        tree.insert('', 'end', values=(
            "Nenhum link de produto encontrado", "", ""))
        progress_bar['value'] = 0
        status_label.config(text="Atualização concluída.")
        root.after(5000, limpar_status)
        return

    produtos = []

    def processar_urls(urls):
        nonlocal produtos
        urls_já_exibidas = set()
        with concurrent.futures.ThreadPoolExecutor() as executor:
            futuros = [executor.submit(obter_nome_e_preco, url)
                       for url in urls]
            for futuro in concurrent.futures.as_completed(futuros):
                nome, preco, url = futuro.result()
                if url not in urls_já_exibidas:
                    produtos.append((nome, preco, url))
                    urls_já_exibidas.add(url)
                progress_bar['value'] += 1
                root.update_idletasks()

    processar_urls(urls_produtos)
    produtos = ordenar_produtos(produtos, ordenar_por)

    for nome, preco, link in produtos:
        tree.insert('', 'end', values=(nome, preco, link))

    progress_bar['value'] = 0
    status_label.config(text="Atualização concluída.")
    root.after(5000, limpar_status)


def limpar_status():
    status_label.config(text="Pronto")


def ordenar_produtos(produtos, ordenar_por):
    if ordenar_por == 'nome':
        return sorted(produtos, key=lambda x: x[0])
    elif ordenar_por == 'preco':
        return sorted(produtos, key=lambda x: float(x[1].replace(',', '.').replace('R$', '').strip()) if x[1].replace(',', '.').replace('R$', '').strip().replace('.', '', 1).isdigit() else float('inf'))
    return produtos


def exportar_para_excel():
    try:
        dados = [tree.item(item, "values") for item in tree.get_children()]
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
        if caminho_excel:
            df.to_excel(caminho_excel, index=False)
            messagebox.showinfo("Exportação concluída",
                                f"Dados exportados para {caminho_excel}")
    except Exception as e:
        messagebox.showerror("Erro na exportação",
                             f"Ocorreu um erro ao exportar: {e}")


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


def configurar_interface():
    global root, tree, caixa_link, progress_bar, status_label
    root = tk.Tk()
    root.title("Atualizador de Preços (ML e Quero-Quero) - Versão 2.3")

    # Configuração do estilo
    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
    style.configure("Treeview", font=("Arial", 9), rowheight=25)

    # Frame principal
    frame = ttk.Frame(root, padding="10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    # Configuração da TreeView com Scrollbar
    tree_scroll = ttk.Scrollbar(frame)
    tree_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))

    tree = ttk.Treeview(frame, columns=COLUNAS, show="headings",
                        selectmode="browse", yscrollcommand=tree_scroll.set)
    tree.heading('Produto', text='Produto')
    tree.heading('Preço', text='Preço')
    tree.heading('Link', text='Link')
    tree.column('Produto', width=400)
    tree.column('Preço', width=80)
    tree.column('Link', width=400)
    tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    tree_scroll.config(command=tree.yview)

    tree.bind("<Double-1>", abrir_link)

    # Frame para o campo de link e botão adicionar
    link_frame = ttk.Frame(frame, padding="5")
    link_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E))
    link_frame.grid_columnconfigure(0, weight=1)

    # Caixa de entrada para adicionar links
    caixa_link = ttk.Entry(link_frame, width=70)
    caixa_link.grid(row=0, column=0, padx=0, pady=10, sticky=(tk.W, tk.E))

    # Botão "Adicionar Novo Produto"
    btn_adicionar = ttk.Button(
        link_frame, text="Adicionar Novo Produto", command=adicionar_link)
    btn_adicionar.grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)

    # Frame para o botão "Atualizar Lista" e a barra de progresso
    atualizar_frame = ttk.Frame(frame, padding="10")
    atualizar_frame.grid(row=1, column=0, columnspan=2,
                         pady=10, sticky=(tk.W, tk.E))
    atualizar_frame.grid_columnconfigure(0, weight=1)

    # Botão "Atualizar Lista"
    btn_atualizar_lista = ttk.Button(
        atualizar_frame, text="Atualizar Lista", command=atualizar_lista, width=20)
    btn_atualizar_lista.pack(padx=10, pady=5)

    # Barra de progresso logo abaixo do botão "Atualizar Lista"
    progress_bar = ttk.Progressbar(
        atualizar_frame, orient="horizontal", length=200, mode="determinate")
    progress_bar.pack(padx=10, pady=5)

    # Status
    status_label = ttk.Label(atualizar_frame, text="Pronto", anchor="center")
    status_label.pack(padx=10, pady=10)

    # Frame para os botões de ação, centralizado
    botoes_frame = ttk.Frame(frame, padding="10")
    botoes_frame.grid(row=5, column=0, columnspan=2,
                      pady=10, sticky=(tk.W, tk.E))
    botoes_frame.grid_rowconfigure(0, weight=1)
    botoes_frame.grid_columnconfigure(0, weight=1)
    botoes_frame.grid_columnconfigure(1, weight=1)
    botoes_frame.grid_columnconfigure(2, weight=1)
    botoes_frame.grid_columnconfigure(3, weight=1)

    # Botões de ações, centralizados
    largura_botao = 20
    btn_ordenar_nome = ttk.Button(botoes_frame, text="Ordenar por Nome",
                                  command=lambda: atualizar_lista('nome'), width=largura_botao)
    btn_ordenar_nome.grid(row=0, column=0, padx=5)

    btn_ordenar_preco = ttk.Button(botoes_frame, text="Ordenar por Preço",
                                   command=lambda: atualizar_lista('preco'), width=largura_botao)
    btn_ordenar_preco.grid(row=0, column=1, padx=5)

    btn_exportar_excel = ttk.Button(
        botoes_frame, text="Exportar para Excel", command=exportar_para_excel, width=largura_botao)
    btn_exportar_excel.grid(row=0, column=2, padx=5)

    btn_excluir = ttk.Button(
        botoes_frame, text="Excluir Selecionado", command=excluir_item, width=largura_botao)
    btn_excluir.grid(row=0, column=3, padx=5)

    # Configurando o redimensionamento
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    atualizar_lista()

    root.mainloop()


if __name__ == "__main__":
    configurar_interface()
