import pandas as pd
from datetime import datetime
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

def obter_area_de_trabalho_padrao():
    return str(Path.home() / "Documents")

def Planilha(arquivo):
    df = pd.read_excel(arquivo, skiprows=7)
    df.columns.values[2:7] = ['Subgrupo', 'Codigo', None, None, 'Produto']

   
    colunas_selecionadas = df.loc[:, ['Subgrupo', 'Codigo', 'Produto', 'PrecoCompra', 'PrecoVenda']]
  
    colunas_selecionadas = colunas_selecionadas.copy()
    colunas_selecionadas.insert(0, ' ', None)
    colunas_selecionadas.loc[:, 'CMV'] = colunas_selecionadas['PrecoCompra'] / colunas_selecionadas['PrecoVenda']
    colunas_selecionadas.loc[:, 'CLASSIFICACAO'] = None
    colunas_selecionadas.loc[:, 'CMV 15%'] = colunas_selecionadas['PrecoCompra'] / 0.15
    colunas_selecionadas.loc[:, 'CMV 20%'] = colunas_selecionadas['PrecoCompra'] / 0.20
    colunas_selecionadas.loc[:, 'CMV 25%'] = colunas_selecionadas['PrecoCompra'] / 0.25
    colunas_selecionadas.loc[:, 'CMV 30%'] = colunas_selecionadas['PrecoCompra'] / 0.30
    colunas_selecionadas.loc[:, 'CMV 40%'] = colunas_selecionadas['PrecoCompra'] / 0.40
    colunas_selecionadas.loc[:, 'CMV 45%'] = colunas_selecionadas['PrecoCompra'] / 0.45
    colunas_selecionadas.loc[:, 'NOVO PRECO'] = None
    colunas_selecionadas.loc[:, 'NOVO CMV'] = None
    colunas_selecionadas.loc[:, 'NOVA CLASSIFICACAO'] = None



    num_arquivo = 1
    pasta_destino = obter_area_de_trabalho_padrao()
    while True:
        timestamp = datetime.now().strftime("%d-%m-%Y")
        nome_arquivo = f'CMV {timestamp} {num_arquivo}.xlsx'
        caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
        if not os.path.exists(caminho_arquivo):
            break
        num_arquivo += 1

    colunas_selecionadas.to_excel(caminho_arquivo, index=False)


def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw() # Esconde a janela principal
    arquivo = filedialog.askopenfilename(filetypes=[("", "*.xls;*.xlsx")])
    if arquivo:
        Planilha(arquivo)

if __name__ == "__main__":
    selecionar_arquivo()
