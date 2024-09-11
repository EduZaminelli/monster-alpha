import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
from fpdf import FPDF
from PIL import Image


# Função para obter o caminho absoluto do recurso (útil para PyInstaller)
def resource_path(relative_path):
    """ Obtém o caminho absoluto do recurso, funciona para dev e para PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# Obtém o caminho da área de trabalho do usuário
desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
report_file = os.path.join(desktop_path, 'relatorio_mensal.pdf')

# Verifica e cria a pasta e os arquivos Excel
path = 'C:/Monster_Dados/'
if not os.path.exists(path):
    os.makedirs(path)

excel_file = path + 'estoque_vendas.xlsx'
historico_entradas_file = path + 'historico_entradas.xlsx'
historico_saidas_file = path + 'historico_saidas.xlsx'

# Se os arquivos não existirem, cria novos
if not os.path.isfile(excel_file):
    df = pd.DataFrame(columns=[
        'ID', 'Produto', 'Marca', 'Sabor', 'Valor Compra', 'Valor Venda', 'Entrada', 'Saída', 'Quant. Estoque',
        'Situação', 'Data Entrada', 'Data Saída'
    ])
    df.to_excel(excel_file, index=False)

if not os.path.isfile(historico_entradas_file):
    historico_entradas = pd.DataFrame(columns=[
        'ID', 'Produto', 'Marca', 'Sabor', 'Valor Compra', 'Quantidade', 'Data Entrada'
    ])
    historico_entradas.to_excel(historico_entradas_file, index=False)

if not os.path.isfile(historico_saidas_file):
    historico_saidas = pd.DataFrame(columns=[
        'ID', 'Produto', 'Marca', 'Sabor', 'Valor Venda', 'Quantidade', 'Data Saída'
    ])
    historico_saidas.to_excel(historico_saidas_file, index=False)


# Funções auxiliares
def carregar_dados():
    return pd.read_excel(excel_file, dtype={'Data Saída': str})


def salvar_dados(df):
    df.to_excel(excel_file, index=False)


def carregar_historico_entradas():
    return pd.read_excel(historico_entradas_file)


def salvar_historico_entradas(df):
    df.to_excel(historico_entradas_file, index=False)


def carregar_historico_saidas():
    return pd.read_excel(historico_saidas_file)


def salvar_historico_saidas(df):
    df.to_excel(historico_saidas_file, index=False)


def atualizar_situacao(df):
    def situacao(row):
        if row['Quant. Estoque'] == 0:
            return 'Esgotado'
        elif row['Quant. Estoque'] < (row['Entrada'] * 0.2):
            return 'Acabando'
        else:
            return 'Normal'

    df['Situação'] = df.apply(situacao, axis=1)
    return df


def calcular_saldo():
    historico_entradas = carregar_historico_entradas()
    historico_saidas = carregar_historico_saidas()

    total_gasto = (historico_entradas['Quantidade'] * historico_entradas['Valor Compra']).sum()
    total_vendas = (historico_saidas['Quantidade'] * historico_saidas['Valor Venda']).sum()

    return total_gasto, total_vendas


def gerar_relatorio():
    total_gasto, total_vendas = calcular_saldo()
    lucro = total_vendas - total_gasto
    status = 'Lucro' if lucro >= 0 else 'Prejuízo'

    # Configuração do PDF
    pdf = FPDF()
    pdf.add_page()

    # Adicionando a imagem como marca d'água
    img_path = resource_path('wsalpha.png')
    with Image.open(img_path) as img:
        img = img.convert("RGBA")
        img.putalpha(50)  # Tornar a imagem semi-transparente para marca d'água
        img.save('temp_mark.png')

    pdf.image('temp_mark.png', x=20, y=50, w=170, h=170)

    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Relatorio Mensal - Controle de Estoque e Vendas", 0, 1, "C")

    pdf.set_font("Arial", size=12)
    pdf.ln(10)
    pdf.cell(0, 10, f"Total Gasto: R${total_gasto:.2f}", 0, 1)
    pdf.cell(0, 10, f"Total Vendas: R${total_vendas:.2f}", 0, 1)
    pdf.cell(0, 10, f"Lucro/Prejuizo: R${lucro:.2f}", 0, 1)
    pdf.cell(0, 10, f"Status: {status}", 0, 1)

    pdf.output(report_file)
    os.remove('temp_mark.png')  # Remover a imagem temporária

    messagebox.showinfo("Relatório", f"Relatório gerado com sucesso!\nStatus: {status}\nSalvo em: {report_file}")


def preencher_campos_produto(produto):
    df = carregar_dados()
    dados_produto = df[df['Produto'] == produto]
    if not dados_produto.empty:
        dados_produto = dados_produto.iloc[0]
        entry_marca_entrada.delete(0, tk.END)
        entry_marca_entrada.insert(0, dados_produto['Marca'])
        entry_sabor_entrada.delete(0, tk.END)
        if pd.notna(dados_produto['Sabor']):
            entry_sabor_entrada.insert(0, dados_produto['Sabor'])
        entry_valor_compra_entrada.delete(0, tk.END)
        entry_valor_compra_entrada.insert(0, f"R$ {dados_produto['Valor Compra']:.2f}")
        entry_valor_venda_entrada.delete(0, tk.END)
        entry_valor_venda_entrada.insert(0, f"R$ {dados_produto['Valor Venda']:.2f}")


def preencher_campos_saida(produto):
    df = carregar_dados()
    dados_produto = df[df['Produto'] == produto]
    if not dados_produto.empty:
        dados_produto = dados_produto.iloc[0]
        entry_marca_saida.delete(0, tk.END)
        entry_marca_saida.insert(0, dados_produto['Marca'])
        entry_sabor_saida.delete(0, tk.END)
        if pd.notna(dados_produto['Sabor']):
            entry_sabor_saida.insert(0, dados_produto['Sabor'])
        entry_valor_venda_saida.delete(0, tk.END)
        entry_valor_venda_saida.insert(0, f"R$ {dados_produto['Valor Venda']:.2f}")


# Função para registrar entrada de produtos
def registrar_entrada():
    df = carregar_dados()
    historico_entradas = carregar_historico_entradas()
    produto = entry_nome_produto.get()
    marca = entry_marca_entrada.get()
    sabor = entry_sabor_entrada.get() if entry_sabor_entrada.get().strip() else None
    valor_compra = float(entry_valor_compra_entrada.get().replace('R$', '').strip())
    valor_venda = float(entry_valor_venda_entrada.get().replace('R$', '').strip())
    quantidade = int(entry_quantidade_entrada.get())
    data_entrada = datetime.now().strftime('%d/%m/%Y')

    produto_existente = df[
        (df['Produto'] == produto) & ((df['Sabor'] == sabor) | (df['Sabor'].isna() & pd.isna(sabor)))]
    if produto_existente.empty:
        novo_id = df['ID'].max() + 1 if not df.empty else 1
        novo_produto = pd.DataFrame({
            'ID': [novo_id],
            'Produto': [produto],
            'Marca': [marca],
            'Sabor': [sabor],
            'Valor Compra': [valor_compra],
            'Valor Venda': [valor_venda],
            'Entrada': [quantidade],
            'Saída': [0],
            'Quant. Estoque': [quantidade],
            'Situação': ['Normal'],
            'Data Entrada': [data_entrada],
            'Data Saída': ['']
        }).dropna(how='all')
        df = pd.concat([df, novo_produto], ignore_index=True)
    else:
        df.loc[(df['Produto'] == produto) & (
                    (df['Sabor'] == sabor) | (df['Sabor'].isna() & pd.isna(sabor))), 'Entrada'] += quantidade
        df.loc[(df['Produto'] == produto) & (
                    (df['Sabor'] == sabor) | (df['Sabor'].isna() & pd.isna(sabor))), 'Quant. Estoque'] += quantidade
        df.loc[(df['Produto'] == produto) & (
                    (df['Sabor'] == sabor) | (df['Sabor'].isna() & pd.isna(sabor))), 'Data Entrada'] = data_entrada

        novo_id = df.loc[
            (df['Produto'] == produto) & ((df['Sabor'] == sabor) | (df['Sabor'].isna() & pd.isna(sabor))), 'ID'].values[
            0]

    novo_entrada = pd.DataFrame({
        'ID': [novo_id],
        'Produto': [produto],
        'Marca': [marca],
        'Sabor': [sabor],
        'Valor Compra': [valor_compra],
        'Quantidade': [quantidade],
        'Data Entrada': [data_entrada]
    }).dropna(how='all')
    historico_entradas = pd.concat([historico_entradas, novo_entrada], ignore_index=True)

    df = atualizar_situacao(df)
    salvar_dados(df)
    salvar_historico_entradas(historico_entradas)
    atualizar_lista_produtos()
    atualizar_saldo()
    messagebox.showinfo("Sucesso", "Entrada registrada com sucesso!")


# Função para registrar saída de produtos
def registrar_saida():
    df = carregar_dados()
    historico_saidas = carregar_historico_saidas()
    produto = entry_produto_saida.get()
    sabor = entry_sabor_saida.get() if entry_sabor_saida.get().strip() else None
    quantidade = int(entry_quantidade_saida.get())
    data_venda = datetime.now().strftime('%d/%m/%Y')

    # Verifica se o produto existe no DataFrame com o sabor especificado
    filtro = (df['Produto'] == produto) & ((df['Sabor'] == sabor) | (df['Sabor'].isna() & pd.isna(sabor)))
    if not df[filtro].empty:
        estoque_atual = df.loc[filtro, 'Quant. Estoque'].values[0]
        valor_venda = df.loc[filtro, 'Valor Venda'].values[0]
        if quantidade > estoque_atual:
            messagebox.showerror("Erro", "Quantidade de saída maior que o estoque disponível!")
        else:
            df.loc[filtro, 'Saída'] += quantidade
            df.loc[filtro, 'Quant. Estoque'] -= quantidade
            df.loc[filtro, 'Data Saída'] = str(data_venda)
            saldo = quantidade * valor_venda

            novo_saida = pd.DataFrame({
                'ID': df.loc[filtro, 'ID'].values[0],
                'Produto': [produto],
                'Marca': df.loc[filtro, 'Marca'].values[0],
                'Sabor': [sabor],
                'Valor Venda': [valor_venda],
                'Quantidade': [quantidade],
                'Data Saída': [data_venda]
            }).dropna(how='all')
            historico_saidas = pd.concat([historico_saidas, novo_saida], ignore_index=True)

            df = atualizar_situacao(df)
            salvar_dados(df)
            salvar_historico_saidas(historico_saidas)
            atualizar_lista_produtos()
            atualizar_saldo()
            messagebox.showinfo("Sucesso", f"Saída registrada com sucesso! Saldo: R${saldo:.2f}")
    else:
        messagebox.showerror("Erro", "Produto ou sabor não encontrado!")


# Função para atualizar a lista de produtos na aba Produtos
def atualizar_lista_produtos(filtros=None, ordenar_por=None, reverso=False):
    df = carregar_dados()

    # Substitui NaN por string vazia
    df['Sabor'] = df['Sabor'].fillna('')

    if filtros:
        for key, value in filtros.items():
            if value:
                df = df[df[key].astype(str).str.contains(value, case=False, na=False)]
    if ordenar_por:
        df = df.sort_values(by=ordenar_por, ascending=not reverso)
    for i in tree_produtos.get_children():
        tree_produtos.delete(i)
    for _, row in df.iterrows():
        tree_produtos.insert("", "end", values=list(row))


def atualizar_saldo():
    total_gasto, total_vendas = calcular_saldo()
    label_total_gasto.config(text=f"Total Gasto: R${total_gasto:.2f}")
    label_total_vendas.config(text=f"Total Vendas: R${total_vendas:.2f}")


def limpar_campos_entrada():
    entry_nome_produto.delete(0, tk.END)
    entry_marca_entrada.delete(0, tk.END)
    entry_sabor_entrada.delete(0, tk.END)
    entry_valor_compra_entrada.delete(0, tk.END)
    entry_valor_venda_entrada.delete(0, tk.END)
    entry_quantidade_entrada.delete(0, tk.END)


def limpar_campos_saida():
    entry_produto_saida.delete(0, tk.END)
    entry_marca_saida.delete(0, tk.END)
    entry_sabor_saida.delete(0, tk.END)
    entry_valor_venda_saida.delete(0, tk.END)
    entry_quantidade_saida.delete(0, tk.END)


def limpar_todos_os_campos():
    # Desativar temporariamente o autocompletar
    original_autocomplete = entry_nome_produto._completion_list
    entry_nome_produto.set_completion_list([])
    entry_sabor_entrada.set_completion_list([])
    entry_produto_saida.set_completion_list([])
    entry_sabor_saida.set_completion_list([])

    limpar_campos_entrada()
    limpar_campos_saida()

    # Reativar o autocompletar com a lista original
    entry_nome_produto.set_completion_list(original_autocomplete)
    entry_sabor_entrada.set_completion_list(sabores)
    entry_produto_saida.set_completion_list(original_autocomplete)
    entry_sabor_saida.set_completion_list(sabores)


def deletar_item():
    selected_item = tree_produtos.selection()[0]
    item_id = tree_produtos.item(selected_item, 'values')[0]
    df = carregar_dados()
    df = df[df['ID'] != int(item_id)]
    salvar_dados(df)
    atualizar_lista_produtos()
    messagebox.showinfo("Sucesso", "Item deletado com sucesso!")


class AutocompleteEntry(ttk.Entry):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list, key=str.lower)
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self.bind('<FocusOut>', self.handle_focusout)
        self.bind('<Return>', self.handle_enter)
        self.bind('<Escape>', self.handle_escape)

    def autocomplete(self, delta=0):
        if delta:
            self.delete(self.position, tk.END)
        else:
            self.position = len(self.get())

        _hits = [item for item in self._completion_list if item.lower().startswith(self.get().lower())]

        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits

        if _hits:
            self._hit_index = (self._hit_index + delta) % len(_hits)
            self.delete(0, tk.END)
            self.insert(0, _hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        if event.keysym in ('BackSpace', 'Left', 'Right', 'Up', 'Down'):
            return
        self.autocomplete()

    def handle_focusout(self, event):
        produto = self.get()
        if produto in self._completion_list:
            if self.master == aba_entrada:
                preencher_campos_produto(produto)
            elif self.master == aba_saida:
                preencher_campos_saida(produto)

    def handle_enter(self, event):
        produto = self.get()
        if produto in self._completion_list:
            if self.master == aba_entrada:
                preencher_campos_produto(produto)
            elif self.master == aba_saida:
                preencher_campos_saida(produto)

    def handle_escape(self, event):
        limpar_todos_os_campos()


# Interface gráfica
root = tk.Tk()
root.title("Monster Alpha Suplements")
root.iconbitmap(resource_path('wsalpha.ico'))
root.option_add('*TCombobox*Listbox.font', 'Century Gothic')

# Adicionar texto de cabeçalho
header_frame = tk.Frame(root)
header_frame.pack(fill='x')

header_label = tk.Label(header_frame, text="Controle de Estoque e Vendas", font=("Century Gothic", 16, "bold"))
header_label.pack(pady=10)

# Adicionar imagem como cabeçalho
img_path = resource_path('wsalpha.png')
img = tk.PhotoImage(file=img_path).subsample(3, 3)
img_label = tk.Label(root, image=img)
img_label.pack(pady=1)

# Notebook (abas)
notebook = ttk.Notebook(root)
notebook.pack(pady=5, expand=True)
notebook.configure(style='TNotebook')

style = ttk.Style()
style.configure('TNotebook.Tab', font=('Century Gothic', 12))
style.map('TNotebook.Tab')

# Aba Produtos
aba_produtos = ttk.Frame(notebook)
notebook.add(aba_produtos, text='Produtos')

columns = ['ID', 'Produto', 'Marca', 'Sabor', 'Valor Compra', 'Valor Venda', 'Entrada', 'Saída', 'Quant. Estoque',
           'Situação']
tree_produtos = ttk.Treeview(aba_produtos, columns=columns, show='headings')
tree_produtos.heading('ID', text='ID')
tree_produtos.column('ID', width=50, minwidth=50, anchor='center')
for col in columns[1:]:
    tree_produtos.heading(col, text=col)
    tree_produtos.column(col, width=100, minwidth=100, anchor='center')
tree_produtos.pack(expand=True, fill='both')

# Função para ordenar colunas
sort_column = None
sort_reverse = False


def ordenar_coluna(col):
    global sort_column, sort_reverse
    sort_reverse = not sort_reverse if sort_column == col else False
    sort_column = col
    atualizar_lista_produtos(ordenar_por=col, reverso=sort_reverse)


for col in columns:
    tree_produtos.heading(col, text=col, command=lambda _col=col: ordenar_coluna(_col))

# Botão para deletar item
btn_deletar_item = tk.Button(aba_produtos, text="Deletar Item", command=deletar_item, font=('Century Gothic', 10))
btn_deletar_item.pack(pady=5)

# Aba Entrada
aba_entrada = ttk.Frame(notebook)
notebook.add(aba_entrada, text='Entrada')

df = carregar_dados()
produtos = list(df['Produto'].dropna().unique())
sabores = list(df['Sabor'].dropna().unique())

tk.Label(aba_entrada, text="Nome do Produto", font=('Century Gothic', 10)).grid(row=0, column=0, padx=5, pady=2)
entry_nome_produto = AutocompleteEntry(aba_entrada, font=('Century Gothic', 10))
entry_nome_produto.set_completion_list(produtos)
entry_nome_produto.grid(row=0, column=1, padx=5, pady=2)

tk.Label(aba_entrada, text="Marca", font=('Century Gothic', 10)).grid(row=1, column=0, padx=5, pady=2)
entry_marca_entrada = ttk.Entry(aba_entrada, font=('Century Gothic', 10))
entry_marca_entrada.grid(row=1, column=1, padx=5, pady=2)

tk.Label(aba_entrada, text="Sabor", font=('Century Gothic', 10)).grid(row=2, column=0, padx=5, pady=2)
entry_sabor_entrada = AutocompleteEntry(aba_entrada, font=('Century Gothic', 10))
entry_sabor_entrada.set_completion_list(sabores)
entry_sabor_entrada.grid(row=2, column=1, padx=5, pady=2)

tk.Label(aba_entrada, text="Valor Compra Uni. (R$)", font=('Century Gothic', 10)).grid(row=3, column=0, padx=5, pady=2)
entry_valor_compra_entrada = ttk.Entry(aba_entrada, font=('Century Gothic', 10))
entry_valor_compra_entrada.grid(row=3, column=1, padx=5, pady=2)

tk.Label(aba_entrada, text="Valor Venda Uni. (R$)", font=('Century Gothic', 10)).grid(row=4, column=0, padx=5, pady=2)
entry_valor_venda_entrada = ttk.Entry(aba_entrada, font=('Century Gothic', 10))
entry_valor_venda_entrada.grid(row=4, column=1, padx=5, pady=2)

tk.Label(aba_entrada, text="Quantidade", font=('Century Gothic', 10)).grid(row=5, column=0, padx=5, pady=2)
entry_quantidade_entrada = ttk.Entry(aba_entrada, font=('Century Gothic', 10))
entry_quantidade_entrada.grid(row=5, column=1, padx=5, pady=2)

tk.Button(aba_entrada, text="Registrar Entrada", command=registrar_entrada, font=('Century Gothic', 10)).grid(row=6,
                                                                                                              columnspan=2,
                                                                                                              pady=5)

# Aba Saída
aba_saida = ttk.Frame(notebook)
notebook.add(aba_saida, text='Saída')

tk.Label(aba_saida, text="Nome do Produto", font=('Century Gothic', 10)).grid(row=0, column=0, padx=5, pady=2)
entry_produto_saida = AutocompleteEntry(aba_saida, font=('Century Gothic', 10))
entry_produto_saida.set_completion_list(produtos)
entry_produto_saida.grid(row=0, column=1, padx=5, pady=2)

tk.Label(aba_saida, text="Quantidade", font=('Century Gothic', 10)).grid(row=1, column=0, padx=5, pady=2)
entry_quantidade_saida = ttk.Entry(aba_saida, font=('Century Gothic', 10))
entry_quantidade_saida.grid(row=1, column=1, padx=5, pady=2)

tk.Label(aba_saida, text="Marca", font=('Century Gothic', 10)).grid(row=2, column=0, padx=5, pady=2)
entry_marca_saida = ttk.Entry(aba_saida, font=('Century Gothic', 10))
entry_marca_saida.grid(row=2, column=1, padx=5, pady=2)

tk.Label(aba_saida, text="Sabor", font=('Century Gothic', 10)).grid(row=3, column=0, padx=5, pady=2)
entry_sabor_saida = AutocompleteEntry(aba_saida, font=('Century Gothic', 10))
entry_sabor_saida.set_completion_list(sabores)
entry_sabor_saida.grid(row=3, column=1, padx=5, pady=2)

tk.Label(aba_saida, text="Valor Venda Uni. (R$)", font=('Century Gothic', 10)).grid(row=4, column=0, padx=5, pady=2)
entry_valor_venda_saida = ttk.Entry(aba_saida, font=('Century Gothic', 10))
entry_valor_venda_saida.grid(row=4, column=1, padx=5, pady=2)

tk.Button(aba_saida, text="Registrar Saída", command=registrar_saida, font=('Century Gothic', 10)).grid(row=5,
                                                                                                        columnspan=2,
                                                                                                        pady=5)

tk.Button(aba_saida, text="Gerar Relatório Atual", command=gerar_relatorio, font=('Century Gothic', 10)).grid(row=6,
                                                                                                              columnspan=2,
                                                                                                              pady=5)

# Labels de saldo total
frame_saldo = ttk.Frame(root)
frame_saldo.pack(pady=5)

label_total_gasto = tk.Label(frame_saldo, text="Total Gasto: R$0.00", font=('Century Gothic', 10))
label_total_gasto.grid(row=0, column=0, padx=5)

label_total_vendas = tk.Label(frame_saldo, text="Total Vendas: R$0.00", font=('Century Gothic', 10))
label_total_vendas.grid(row=0, column=1, padx=5)

# Rodapé
footer_frame = tk.Frame(root)
footer_frame.pack(fill='x')

footer_label = tk.Label(footer_frame, text="Developed by Edu Zaminelli", font=('Century Gothic', 8))
footer_label.pack(pady=2)

# Verifica se é o primeiro dia do mês para gerar o relatório
if datetime.now().day == 1:
    gerar_relatorio()

# Carregar dados na lista de produtos e atualizar saldo ao iniciar
atualizar_lista_produtos()
atualizar_saldo()

# Rodar a aplicação
root.mainloop()
