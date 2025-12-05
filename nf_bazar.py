import pandas as pd
import win32print
import win32ui
import datetime as dt
import re
from num2words import num2words
import textwrap
import numpy as np
import tkinter as tk

# ================= DATAS =================

day = dt.datetime.now().day
month = dt.datetime.now().month
year = dt.datetime.now().year

# ================= FUNÇÕES AUXILIARES =================

def convert_google_sheet_url(url):
    pattern = r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)(/edit#gid=(\d+)|/edit.*)?'
    replacement = lambda m: f'https://docs.google.com/spreadsheets/d/{m.group(1)}/export?' + (f'gid={m.group(3)}&' if m.group(3) else '') + 'format=csv'
    return re.sub(pattern, replacement, url)

# ================= BASE DE DADOS =================

cpf_banco = pd.read_csv(
    convert_google_sheet_url(
        'https://docs.google.com/spreadsheets/d/1R2ziBev9t4c8xJpWbf5rtzjFMpCgfEko7B3I2iBMhG4/edit?gid=0#gid=0'
    ), dtype=str
)

produto_banco = pd.read_excel('estoque.xlsx')
produto_banco.columns = produto_banco.columns.str.strip()
produto_banco['PREÇO DE VENDA'] = np.around(produto_banco['PREÇO DE VENDA'].astype(float), 2)

produtos_disponiveis = produto_banco['PRODUTO'].dropna().astype(str).tolist()
nomes_disponiveis = cpf_banco['NOME'].dropna().astype(str).tolist()

# ================= VARIÁVEIS GLOBAIS =================

PRODUTO = None
VALOR = None
VENDEDOR = None
CPF = None
NOME = None

carrinho = []

# ================= FUNÇÕES CLIENTE =================

def atualizar_lista_cpf(*args):
    termo = entrada_cpf.get().lower()
    lista_cpf.delete(0, tk.END)

    for nome in nomes_disponiveis:
        if termo in nome.lower():
            lista_cpf.insert(tk.END, nome)

def selecionar_cpf(event):
    if lista_cpf.curselection():
        selecionado = lista_cpf.get(lista_cpf.curselection())
        entrada_cpf.delete(0, tk.END)
        entrada_cpf.insert(0, selecionado)

# ================= FUNÇÕES PRODUTO =================

def atualizar_lista(*args):
    termo = entrada_produto.get().lower()
    lista_produto.delete(0, tk.END)

    for produto in produtos_disponiveis:
        if termo in produto.lower():
            lista_produto.insert(tk.END, produto)

def selecionar_produto(event):
    if lista_produto.curselection():
        selecionado = lista_produto.get(lista_produto.curselection())
        entrada_produto.delete(0, tk.END)
        entrada_produto.insert(0, selecionado)

# ================= CARRINHO =================

def atualizar_carrinho_na_tela():
    lista_carrinho.delete(0, tk.END)
    total = 0

    for item in carrinho:
        texto = f"{item['produto']} | {item['quantidade']} un | R$ {item['subtotal']:.2f}"
        lista_carrinho.insert(tk.END, texto)
        total += item['subtotal']

    total_var.set(f"Total: R$ {total:.2f}")

def remover_item():
    if lista_carrinho.curselection():
        index = lista_carrinho.curselection()[0]
        carrinho.pop(index)
        atualizar_carrinho_na_tela()

# ================= CONTROLE PRODUTO =================

def confirmar_produto():
    global PRODUTO, VALOR, CPF, NOME

    produto_final = entrada_produto.get()

    if produto_final in produtos_disponiveis:
        PRODUTO = produto_final

        linha = produto_banco[produto_banco['PRODUTO'] == PRODUTO]
        if not linha.empty:
            VALOR = float(linha['PREÇO DE VENDA'].values[0])

        QNTD = int(qntd_spinbox.get())
        VENDEDOR = vendedor_var.get()

        nome_final = entrada_cpf.get()
        linha_cpf = cpf_banco[cpf_banco['NOME'] == nome_final]

        if not linha_cpf.empty:
            CPF = linha_cpf['CPF'].values[0]
            NOME = nome_final
        else:
            return

        item = {
            "produto": PRODUTO,
            "valor_unitario": VALOR,
            "quantidade": QNTD,
            "subtotal": VALOR * QNTD,
            "vendedor": VENDEDOR
        }

        carrinho.append(item)
        atualizar_carrinho_na_tela()
        entrada_produto.delete(0, tk.END)

# ================= IMPRESSÃO =================

def imprimir_cupom():
    total_geral = sum(item["subtotal"] for item in carrinho)

    lista_produtos_texto = ""
    for item in carrinho:
        linha = (
            f"- {item['produto']} | "
            f"Qtd: {item['quantidade']} | "
            f"Unid: R$ {item['valor_unitario']:.2f} | "
            f"Subtotal: R$ {item['subtotal']:.2f}\n"
        )
        lista_produtos_texto += linha

    valor_extenso = num2words(total_geral, lang='pt_BR', to='currency')

    texto = f"""
RECIBO FISCAL

Declaro, para os devidos fins, que recebi de {NOME},
inscrito no CPF sob o nº {CPF},
a importância total de R$ {total_geral:.2f}
({valor_extenso}),

referente ao pagamento correspondente à aquisição dos seguintes produtos
no Bazar Beneficente de itens doados pela Receita Federal:

{lista_produtos_texto}

Nova Serrana - MG, {day} de {month} de {year}.

Assinatura:

Responsável pelo atendimento:

{VENDEDOR}
"""

    printer_name = win32print.GetDefaultPrinter()
    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)

    font = win32ui.CreateFont({
        "name": "Consolas",
        "height": 28,
        "width": 12,
        "weight": 400,
    })

    hDC.SelectObject(font)
    hDC.StartDoc("Recibo")
    hDC.StartPage()

    LIMITE = 45

    linhas_formatadas = []
    for linha in texto.split("\n"):
        if linha.strip():
            linhas_formatadas.extend(textwrap.wrap(linha, width=LIMITE))
        else:
            linhas_formatadas.append("")

    y = 10
    for linha in linhas_formatadas:
        hDC.TextOut(10, y, linha)
        y += 40

    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()

# ================= RESET PARA NOVA VENDA =================

def resetar_tela():
    global carrinho, PRODUTO, VALOR, VENDEDOR, CPF, NOME

    carrinho = []

    PRODUTO = None
    VALOR = None
    VENDEDOR = None
    CPF = None
    NOME = None

    entrada_produto.delete(0, tk.END)
    entrada_cpf.delete(0, tk.END)

    lista_carrinho.delete(0, tk.END)
    lista_produto.delete(0, tk.END)
    lista_cpf.delete(0, tk.END)

    qntd_spinbox.delete(0, tk.END)
    qntd_spinbox.insert(0, 1)

    vendedor_var.set(lista_vendedores[0])

    total_var.set("Total: R$ 0.00")

# ================= FINALIZAÇÃO =================

def finalizar_compra():
    if cart:=carrinho:
        imprimir_cupom()
        resetar_tela()

# ================= INTERFACE =================

janela = tk.Tk()
janela.title("Sistema de Vendas")
janela.geometry("900x550")
janela.tk.call('tk', 'scaling', 0.9)

frame_esquerdo = tk.Frame(janela)
frame_esquerdo.grid(row=0, column=0, padx=10, pady=10, sticky="n")

frame_direito = tk.Frame(janela)
frame_direito.grid(row=0, column=1, padx=10, pady=10, sticky="n")

# COLUNA ESQUERDA

tk.Label(frame_esquerdo, text="Produto").grid(row=0, column=0, sticky="w")

entrada_produto = tk.Entry(frame_esquerdo, width=40)
entrada_produto.grid(row=1, column=0)
entrada_produto.bind("<KeyRelease>", atualizar_lista)

lista_produto = tk.Listbox(frame_esquerdo, width=55, height=6)
lista_produto.grid(row=2, column=0)
lista_produto.bind("<<ListboxSelect>>", selecionar_produto)

tk.Label(frame_esquerdo, text="Quantidade").grid(row=3, column=0, sticky="w")
qntd_spinbox = tk.Spinbox(frame_esquerdo, from_=1, to=100, width=5)
qntd_spinbox.grid(row=4, column=0, sticky="w")

tk.Label(frame_esquerdo, text="Carrinho").grid(row=5, column=0, sticky="w")
lista_carrinho = tk.Listbox(frame_esquerdo, width=55, height=8)
lista_carrinho.grid(row=6, column=0)

total_var = tk.StringVar()
total_var.set("Total: R$ 0.00")

tk.Label(frame_esquerdo, textvariable=total_var).grid(row=7, column=0, sticky="e")

# COLUNA DIREITA

tk.Label(frame_direito, text="Cliente (Nome)").grid(row=0, column=0, sticky="w")

entrada_cpf = tk.Entry(frame_direito, width=35)
entrada_cpf.grid(row=1, column=0)
entrada_cpf.bind("<KeyRelease>", atualizar_lista_cpf)

lista_cpf = tk.Listbox(frame_direito, width=55, height=6)
lista_cpf.grid(row=2, column=0)
lista_cpf.bind("<<ListboxSelect>>", selecionar_cpf)

tk.Label(frame_direito, text="Vendedor").grid(row=3, column=0)

lista_vendedores = ["vendedor1", "vendedor2", "vendedor3"]
vendedor_var = tk.StringVar()
vendedor_var.set(lista_vendedores[0])

menu_vendedores = tk.OptionMenu(frame_direito, vendedor_var, *lista_vendedores)
menu_vendedores.grid(row=4, column=0)

btn_add = tk.Button(frame_direito, text="Adicionar ao carrinho", width=25, command=confirmar_produto)
btn_add.grid(row=5, column=0, pady=10)

btn_remover = tk.Button(frame_direito, text="Remover item", width=25, command=remover_item)
btn_remover.grid(row=6, column=0)

btn_finalizar = tk.Button(frame_direito, text="Finalizar venda", width=25, command=finalizar_compra)
btn_finalizar.grid(row=7, column=0, pady=20)

janela.mainloop()
