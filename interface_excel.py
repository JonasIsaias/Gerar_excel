import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox
import matplotlib
matplotlib.use('TkAgg')

def criar_excel(nome_arquivo, colunas, dados):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Adiciona cabeçalhos
    for col_num, coluna in enumerate(colunas, start=1):
        sheet.cell(row=1, column=col_num, value=coluna)

    # Adiciona dados
    for linha in dados:
        sheet.append(linha)

    workbook.save(nome_arquivo)
    messagebox.showinfo("Sucesso", f'Arquivo {nome_arquivo} criado com sucesso!')

def carregar_excel(nome_arquivo):
    try:
        df = pd.read_excel(nome_arquivo)
        return df
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {e}")
        return None

def gerar_grafico(df):
    tipos_grafico = {
        '1': 'bar',
        '2': 'line',
        '3': 'pie'
    }

    def plotar_grafico():
        tipo_grafico = tipo_grafico_var.get()
        coluna_x = coluna_x_entry.get()
        coluna_y = coluna_y_entry.get()

        if coluna_x not in df.columns or coluna_y not in df.columns:
            messagebox.showerror("Erro", "Uma ou mais colunas não existem no DataFrame.")
            return

        if tipo_grafico == '1':
            df.plot(kind='bar', x=coluna_x, y=coluna_y)
        elif tipo_grafico == '2':
            df.plot(kind='line', x=coluna_x, y=coluna_y)
        elif tipo_grafico == '3':
            df.plot(kind='pie', y=coluna_y, labels=df[coluna_x], autopct='%1.1f%%')
        else:
            messagebox.showerror("Erro", "Tipo de gráfico inválido.")
            return

        plt.title(f'Gráfico de {tipo_grafico} de {coluna_y} vs {coluna_x}')
        plt.show()

    tipo_grafico_var = Tk()
    tipo_grafico_var.title("Escolha Tipo de Gráfico")

    Label(tipo_grafico_var, text="Escolha o tipo de gráfico").pack()
    Label(tipo_grafico_var, text="1. Gráfico de Barras").pack()
    Label(tipo_grafico_var, text="2. Gráfico de Linhas").pack()
    Label(tipo_grafico_var, text="3. Gráfico de Pizza").pack()

    tipo_grafico_entry = Entry(tipo_grafico_var)
    tipo_grafico_entry.pack()

    Label(tipo_grafico_var, text="Coluna para o eixo X:").pack()
    coluna_x_entry = Entry(tipo_grafico_var)
    coluna_x_entry.pack()

    Label(tipo_grafico_var, text="Coluna para o eixo Y:").pack()
    coluna_y_entry = Entry(tipo_grafico_var)
    coluna_y_entry.pack()

    Button(tipo_grafico_var, text="Gerar Gráfico", command=plotar_grafico).pack()

    tipo_grafico_var.mainloop()

def selecionar_arquivo_entrada():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo:
        df = carregar_excel(arquivo)
        if df is not None:
            gerar_grafico(df)

def criar_arquivo():
    nome_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    if nome_arquivo:
        colunas = entry_colunas.get().split(',')
        colunas = [coluna.strip() for coluna in colunas]
        dados = []
        while True:
            linha_dados = entry_dados.get()
            if linha_dados.lower() == 'sair':
                break
            linha_dados = linha_dados.split(',')
            linha_dados = [dado.strip() for dado in linha_dados]
            if len(linha_dados) != len(colunas):
                messagebox.showerror("Erro", f"Esperado {len(colunas)} valores, mas recebeu {len(linha_dados)}. Tente novamente.")
                continue
            dados.append(linha_dados)
        criar_excel(nome_arquivo, colunas, dados)

app = Tk()
app.title("Gerenciador de Excel")

Label(app, text="Digite os nomes das colunas, separados por vírgula:").pack()
entry_colunas = Entry(app)
entry_colunas.pack()

Label(app, text="Digite os dados para cada linha, separados por vírgula (digite 'sair' para encerrar):").pack()
entry_dados = Entry(app)
entry_dados.pack()

Button(app, text="Criar Novo Arquivo Excel", command=criar_arquivo).pack()
Button(app, text="Carregar e Analisar Arquivo Excel", command=selecionar_arquivo_entrada).pack()
Button(app, text="Sair", command=app.quit).pack()

app.mainloop()
