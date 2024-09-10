import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import webbrowser
import pandas as pd
import concurrent.futures

# Função para obter nome e preço do produto


def obter_nome_e_preco(url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
    except requests.RequestException as e:
        return 'Nome não encontrado', 'Preço não encontrado', url

    soup = BeautifulSoup(response.content, 'html.parser')
    nome_element = soup.select_one(
        '.ui-pdp-title')
    preco_element = soup.select_one(
        '.ui-pdp-price__second-line')

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
    novo_link = caixa_link.get().strip()  # Remove espaços em branco extras
    if novo_link:
        caminho_arquivo = criar_arquivo_links()

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

# Função para atualizar a lista de produtos usando threads


def atualizar_lista(ordenar_por=None):
    caminho_arquivo = criar_arquivo_links()
    urls_produtos = ler_links_arquivo(caminho_arquivo)

    for item in tree.get_children():
        tree.delete(item)

    if not urls_produtos:
        tree.insert('', 'end', values=(
            "Nenhum link de produto encontrado", "", ""))
        return

    produtos = []

    # Usa threads para obter os dados dos produtos em paralelo
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futuros = [executor.submit(obter_nome_e_preco, url)
                   for url in urls_produtos]
        for futuro in concurrent.futures.as_completed(futuros):
            produtos.append(futuro.result())

    if ordenar_por == 'nome':
        produtos.sort(key=lambda x: x[0])
    elif ordenar_por == 'preco':
        produtos.sort(key=lambda x: float(x[1].replace(',', '.').replace('R$', '').strip()) if x[1].replace(
            ',', '.').replace('R$', '').strip().replace('.', '', 1).isdigit() else float('inf'))

    for nome, preco, link in produtos:
        tree.insert('', 'end', values=(nome, preco, link))

# Função para exportar para Excel


def exportar_para_excel():
    try:
        dados = []
        for item in tree.get_children():
            dados.append(tree.item(item, "values"))

        # Verifica se há dados para exportar
        if not dados:
            messagebox.showwarning("Exportação falhou",
                                   "Não há dados para exportar.")
            return

        # Cria um DataFrame a partir dos dados
        df = pd.DataFrame(dados, columns=["Produto", "Preço", "Link"])

        # Abre uma caixa de diálogo para o usuário escolher onde salvar o arquivo
        caminho_excel = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Salvar como"
        )

        # Verifica se o usuário cancelou a operação
        if not caminho_excel:
            return

        # Salva o DataFrame como um arquivo Excel
        df.to_excel(caminho_excel, index=False)

        # Mostra uma mensagem de sucesso
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
    caminho_arquivo = criar_arquivo_links()
    with open(caminho_arquivo, 'r') as f:
        linhas = f.readlines()

    with open(caminho_arquivo, 'w') as f:
        for linha in linhas:
            if linha.strip() != link:
                f.write(linha)

# Função para garantir que o diretório e o arquivo TXT existam


def criar_arquivo_links():
    diretorio = 'C:\\Atualizador de Preços (Mercado Livre)'
    caminho_arquivo = os.path.join(diretorio, 'produtos.txt')

    if not os.path.exists(diretorio):
        os.makedirs(diretorio)

    if not os.path.isfile(caminho_arquivo):
        with open(caminho_arquivo, 'w') as f:
            pass  # Apenas cria o arquivo, não escreve nada

    return caminho_arquivo


# Criação da interface gráfica
janela = tk.Tk()
janela.title("Lista de Produtos e Preços - Versão 0.1")

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

# Botões de Atualizar, Exportar e Excluir
frame_botoes = tk.Frame(janela)
frame_botoes.pack(pady=5)

botao_atualizar = tk.Button(frame_botoes, text="Atualizar Preços",
                            command=lambda: atualizar_lista(), font=("Helvetica", 12))
botao_atualizar.pack(side=tk.LEFT, padx=5)

botao_exportar = tk.Button(frame_botoes, text="Exportar para Excel",
                           command=exportar_para_excel, font=("Helvetica", 12))
botao_exportar.pack(side=tk.LEFT, padx=5)

botao_excluir = tk.Button(frame_botoes, text="Excluir Selecionado",
                          command=excluir_item, font=("Helvetica", 12))
botao_excluir.pack(side=tk.LEFT, padx=5)

# Botões de Ordenação
frame_ordenacao = tk.Frame(janela)
frame_ordenacao.pack(pady=10)

botao_ordenar_nome = tk.Button(frame_ordenacao, text="Ordenar por Nome",
                               command=lambda: atualizar_lista('nome'), font=("Helvetica", 12))
botao_ordenar_nome.pack(side=tk.LEFT, padx=5)

botao_ordenar_preco = tk.Button(frame_ordenacao, text="Ordenar por Preço",
                                command=lambda: atualizar_lista('preco'), font=("Helvetica", 12))
botao_ordenar_preco.pack(side=tk.LEFT, padx=5)

centralizar_colunas()

# Ajusta a resolução para 800x600
janela.geometry("800x600")
janela.mainloop()
