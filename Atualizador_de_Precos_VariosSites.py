### VERSÃO 2.4.4 ###
# UPDATES:
# 1 - REFORMULAÇÃO DAS BARRAS E BOTÕES
# 2 - OPÇÃO DE ADICIONAR PRODUTOS DO MERCADO LIVRE E LOJAS QUERO QUERO
# 3 - ADICIONADA UMA COLUNA QUE MOSTRA O SITE 
# 4 - MELHORIAS NA INTERFACE
# 5 - MELHORIAS NO DESEMPENHO
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
import threading
import logging

# Configuração do logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Constantes de configuração
NOME_PROGRAMA = "Atualizador de Preços (ML e Quero-Quero) - Versão 2.4."
CAMINHO_ARQUIVO = f'C:\\Atualizador de Preços (ML e Quero-Quero)\\produtos.txt'
DIRETORIO_BACKUP = f'C:\\Atualizador de Preços (ML e Quero-Quero)\\backup\\'
COLUNAS = ('Produto', 'Preço', 'Link', 'Site')

# Variáveis globais
root = None  # Inicializa como None e será definido na função criar_interface


def obter_nome_e_preco_queroquero(url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers, timeout=20)
        response.raise_for_status()
    except requests.RequestException as e:
        logging.error(f"Erro ao acessar {url}: {e}")
        return 'Nome não encontrado', 'Preço não encontrado', url, 'queroquero.com'

    soup = BeautifulSoup(response.content, 'html.parser')
    nome_element = soup.select_one('.vtex-store-components-3-x-productBrand--quickview')
    preco_element = soup.select_one('.vtex-product-price-1-x-currencyContainer--product-price')

    nome = nome_element.get_text(strip=True) if nome_element else 'Nome não encontrado'
    preco = preco_element.get_text(strip=True).replace("R$", "").strip() if preco_element else 'Preço não encontrado'

    return nome, preco, url, 'queroquero.com'


def obter_nome_e_preco_mercadolivre(url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers, timeout=20)
        response.raise_for_status()
    except requests.RequestException as e:
        logging.error(f"Erro ao acessar {url}: {e}")
        return 'Nome não encontrado', 'Preço não encontrado', url, 'mercadolivre.com'

    soup = BeautifulSoup(response.content, 'html.parser')
    nome_element = soup.select_one('h1.ui-pdp-title')
    preco_element = soup.select_one('.andes-money-amount__fraction')

    nome = nome_element.get_text(strip=True) if nome_element else 'Nome não encontrado'
    preco = preco_element.get_text(strip=True).replace("R$", "").strip() if preco_element else 'Preço não encontrado'

    return nome, preco, url, 'mercadolivre.com'


def obter_nome_e_preco(url):
    if "queroquero.com" in url:
        return obter_nome_e_preco_queroquero(url)
    elif "mercadolivre.com" in url:
        return obter_nome_e_preco_mercadolivre(url)
    else:
        return 'Site não suportado', 'Preço não encontrado', url, 'desconhecido'


def ler_links_arquivo(arquivo):
    caminho_arquivo = os.path.join(os.path.dirname(__file__), arquivo)
    if not os.path.isfile(caminho_arquivo):
        logging.warning(f"O arquivo {caminho_arquivo} não existe.")
        return []
    with open(caminho_arquivo, 'r') as f:
        return f.read().splitlines()


def fazer_backup(caminho_arquivo):
    if not os.path.exists(DIRETORIO_BACKUP):
        os.makedirs(DIRETORIO_BACKUP, exist_ok=True)
        logging.info(f"Pasta de backup criada em: {DIRETORIO_BACKUP}")

    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    backup_arquivo = os.path.join(DIRETORIO_BACKUP, f'produtos_backup_{timestamp}.txt')
    try:
        shutil.copy(caminho_arquivo, backup_arquivo)
        logging.info(f"Backup criado em: {backup_arquivo}")
    except Exception as e:
        logging.error(f"Erro ao criar o backup: {e}")


def adicionar_link():
    novo_link = caixa_link.get().strip()
    if novo_link:
        if verificar_link_existente(novo_link):
            messagebox.showwarning("Produto duplicado", "O produto já está na lista.")
            return

        fazer_backup(CAMINHO_ARQUIVO)
        salvar_link(novo_link)
        caixa_link.delete(0, tk.END)
        atualizar_lista()
    else:
        messagebox.showwarning("Entrada inválida", "Por favor, insira um link válido.")


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
    def processar_lista():
        urls_produtos = ler_links_arquivo(CAMINHO_ARQUIVO)
        tree.delete(*tree.get_children())
        progress_bar['value'] = 0
        progress_bar['maximum'] = len(urls_produtos)
        status_label.config(text="Atualizando...")

        if not urls_produtos:
            tree.insert('', 'end', values=("Nenhum link de produto encontrado", "", "", ""))
            progress_bar['value'] = 0
            status_label.config(text="Atualização concluída.")
            root.after(5000, limpar_status)
            return

        produtos = []

        def processar_urls(urls):
            nonlocal produtos
            urls_já_exibidas = set()
            with concurrent.futures.ThreadPoolExecutor() as executor:
                futuros = [executor.submit(obter_nome_e_preco, url) for url in urls]
                for futuro in concurrent.futures.as_completed(futuros):
                    nome, preco, url, site = futuro.result()
                    if url not in urls_já_exibidas:
                        produtos.append((nome, preco, url, site))
                        urls_já_exibidas.add(url)
                    root.after(0, lambda: progress_bar.configure(value=progress_bar['value'] + 1))

        processar_urls(urls_produtos)
        produtos = ordenar_produtos(produtos, ordenar_por)

        for nome, preco, link, site in produtos:
            tree.insert('', 'end', values=(nome, preco, link, site))

        ajustar_largura_colunas()

        progress_bar['value'] = 0
        status_label.config(text="Atualização concluída.")
        root.after(5000, limpar_status)

    threading.Thread(target=processar_lista).start()


def limpar_status():
    status_label.config(text="Pronto")


def ordenar_produtos(produtos, ordenar_por):
    if ordenar_por == 'nome':
        return sorted(produtos, key=lambda x: x[0])
    elif ordenar_por == 'preco':
        return sorted(produtos, key=lambda x: float(x[1].replace(',', '.').replace('R$', '').strip()) if x[1].replace(',', '.').replace('R$', '').strip().replace('.', '', 1).isdigit() else float('inf'))
    return produtos


def ajustar_largura_colunas():
    for coluna in COLUNAS:
        largura_max = max([len(tree.item(item, 'values')[COLUNAS.index(coluna)]) for item in tree.get_children()] + [len(coluna)])
        if coluna == 'Link':
            largura_max = max(int(largura_max * 0.20), 20)
        else:
            tree.column(coluna, width=largura_max * 10)
        tree.column(coluna, width=largura_max, anchor=tk.CENTER)  # Ajusta a largura de cada coluna


def exportar_para_excel():
    try:
        caminho_excel = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if caminho_excel:
            dados = [(tree.item(item, "values")) for item in tree.get_children()]
            df = pd.DataFrame(dados, columns=COLUNAS)
            df.to_excel(caminho_excel, index=False)
            logging.info(f"Dados exportados para {caminho_excel}")
    except Exception as e:
        logging.error(f"Erro ao exportar para Excel: {e}")
        messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")


def criar_interface():
    global root, caixa_link, tree, progress_bar, status_label

    root = tk.Tk()
    root.title(NOME_PROGRAMA)
    root.geometry("1000x600")

    # Frame superior
    frame_superior = tk.Frame(root)
    frame_superior.pack(pady=10)

    tk.Label(frame_superior, text="Link do produto:").grid(row=0, column=0, padx=10)
    caixa_link = tk.Entry(frame_superior, width=80)
    caixa_link.grid(row=0, column=1, padx=10)
    tk.Button(frame_superior, text="Adicionar Novo Produto", command=adicionar_link).grid(row=0, column=2, padx=10)

    # Frame inferior
    frame_inferior = tk.Frame(root)
    frame_inferior.pack(pady=10)

    tk.Button(frame_inferior, text="Atualizar Lista", command=lambda: atualizar_lista()).grid(row=0, column=0, padx=10)
    tk.Button(frame_inferior, text="Exportar para Excel", command=exportar_para_excel).grid(row=0, column=1, padx=10)
    tk.Button(frame_inferior, text="Excluir Selecionado", command=lambda: excluir_do_arquivo(tree.item(tree.selection()[0], "values")[2])).grid(row=0, column=2, padx=10)
    tk.Button(frame_inferior, text="Ordenar por Nome", command=lambda: atualizar_lista(ordenar_por='nome')).grid(row=1, column=0, padx=10)
    tk.Button(frame_inferior, text="Ordenar por Preço", command=lambda: atualizar_lista(ordenar_por='preco')).grid(row=1, column=1, padx=10)

    # Status e Barra de progresso
    status_label = tk.Label(root, text="Pronto")
    status_label.pack(pady=5)
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=800, mode="determinate")
    progress_bar.pack(pady=10)

    # TreeView
    tree = ttk.Treeview(root, columns=COLUNAS, show='headings')
    for coluna in COLUNAS:
        tree.heading(coluna, text=coluna)
        tree.column(coluna, width=50, anchor=tk.CENTER)
    tree.bind('<Double-1>', abrir_link)
    tree.pack(expand=True, fill='both', padx=10)

    atualizar_lista()

    root.mainloop()


if __name__ == "__main__":
    criar_interface()

