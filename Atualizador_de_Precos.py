### VERSÃO 2.2 ###
# OTIMIZAÇÃO E REFATORAÇÃO #
# ADIÇÃO DA BARRA DE CARREGAMENTO #
# 10/09/2024 #


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
CAMINHO_ARQUIVO = 'C:\\Atualizador de Preços (Quero-Quero)\\produtos.txt'
DIRETORIO_BACKUP = 'C:\\Atualizador de Preços (Quero-Quero)\\backup\\'
COLUNAS = ('Produto', 'Preço', 'Link')


def obter_nome_e_preco(url):
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
    root.title("Lista de Produtos e Preços - Versão 2.2")

    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Helvetica", 14))
    style.configure("Treeview", rowheight=30, font=("Helvetica", 10))

    tree = ttk.Treeview(root, columns=COLUNAS,
                        show='headings', style="Treeview")
    for coluna in COLUNAS:
        tree.heading(coluna, text=coluna)
        tree.column(coluna, width=250 if coluna ==
                    'Produto' else 75, anchor='center')

    tree.pack(fill=tk.BOTH, expand=True)
    tree.bind("<Double-1>", abrir_link)

    frame_adicionar = tk.Frame(root)
    frame_adicionar.pack(pady=10)

    tk.Label(frame_adicionar, text="Adicionar Produto:",
             font=("Helvetica", 12)).pack(side=tk.LEFT)
    caixa_link = tk.Entry(frame_adicionar, width=50, font=("Helvetica", 12))
    caixa_link.pack(side=tk.LEFT, padx=5)

    tk.Button(frame_adicionar, text="Adicionar", command=adicionar_link,
              font=("Helvetica", 12)).pack(side=tk.LEFT)

    frame_botoes = tk.Frame(root)
    frame_botoes.pack(pady=5)

    tk.Button(frame_botoes, text="Atualizar Preços", command=lambda: atualizar_lista(
    ), font=("Helvetica", 12)).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes, text="Exportar para Excel", command=exportar_para_excel, font=(
        "Helvetica", 12)).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes, text="Excluir Selecionado", command=excluir_item, font=(
        "Helvetica", 12)).pack(side=tk.LEFT, padx=5)

    frame_ordenacao = tk.Frame(root)
    frame_ordenacao.pack(pady=10)

    tk.Button(frame_ordenacao, text="Ordenar por Nome", command=lambda: atualizar_lista(
        'nome'), font=("Helvetica", 12)).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_ordenacao, text="Ordenar por Preço", command=lambda: atualizar_lista(
        'preco'), font=("Helvetica", 12)).pack(side=tk.LEFT, padx=5)

    frame_progresso = tk.Frame(root)
    frame_progresso.pack(pady=10, fill=tk.X)

    progress_bar = ttk.Progressbar(
        frame_progresso, orient="horizontal", length=400, mode="determinate")
    progress_bar.pack(pady=5, fill=tk.X)

    status_label = tk.Label(
        frame_progresso, text="Pronto", font=("Helvetica", 12))
    status_label.pack()

    root.geometry("800x600")
    root.mainloop()


if __name__ == "__main__":
    configurar_interface()
