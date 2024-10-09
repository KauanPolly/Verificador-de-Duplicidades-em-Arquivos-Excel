import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from openpyxl import Workbook

def converter_xls_para_xlsx(filepath):
    """Converte um arquivo .xls para .xlsx ou tenta ler como HTML se necessário."""
    try:
        # Primeiro tenta carregar o arquivo como um Excel verdadeiro usando xlrd
        try:
            xls = pd.read_excel(filepath, engine='xlrd', header=0)  # Garante que a primeira linha seja o header
        except Exception as e:
            # Se der erro, tenta ler como HTML
            messagebox.showinfo("Aviso", "Tentando converter o arquivo como HTML.")
            dfs = pd.read_html(filepath, header=0)  # Ler como HTML, usando a primeira linha como header
            xls = dfs[0]  # Pega a primeira tabela
            
        # Definir o novo caminho com extensão .xlsx
        new_filepath = filepath.replace('.xls', '.xlsx')
        
        # Salvar o DataFrame no formato .xlsx
        xls.to_excel(new_filepath, index=False)
        
        return new_filepath  # Retornar o caminho do novo arquivo .xlsx
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao converter o arquivo .xls: {str(e)}")
        return None

def verificar_duplicidade_login_data(df):
    resultados = []
    
    # Converter a coluna de datas para datetime
    df[df.columns[15]] = pd.to_datetime(df[df.columns[15]], errors='coerce')
    
    for login, group in df.groupby(df.columns[1]):  # Usar a coluna de índice 1 (Login)
        if group.duplicated(subset=[df.columns[15]], keep=False).any():  # Usar a coluna de índice 15 (datavenc)
            # Obter datas duplicadas e formatá-las corretamente
            datas_duplicadas = group[group.duplicated(subset=[df.columns[15]], keep=False)][df.columns[15]]
            datas_formatadas = [d.strftime('%d/%m/%Y') for d in datas_duplicadas if pd.notnull(d)]  # Ignorar valores NaT
            resultados.append(f"Login {login} possui duplicidade nas datas de vencimento: {', '.join(datas_formatadas)}")
    return resultados

def selecionar_arquivo():
    filepath = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filepath:
        # Verificar se o arquivo é .xls
        if filepath.endswith('.xls'):
            # Converter para .xlsx
            filepath = converter_xls_para_xlsx(filepath)
            if not filepath:
                return  # Se a conversão falhar, pare a execução
        
        # Carregar o arquivo Excel e verificar duplicidades
        try:
            df = pd.read_excel(filepath, header=0)  # Usar a primeira linha como header
            
            # Garantir que os nomes das colunas são strings e remover espaços em branco
            df.columns = df.columns.astype(str).str.strip()
            
            # Exibir as colunas disponíveis no arquivo
            print("Colunas disponíveis no arquivo:", df.columns)
            
            # Verificar duplicidades
            resultados = verificar_duplicidade_login_data(df)
            if resultados:
                messagebox.showinfo("Duplicidades encontradas", "\n".join(resultados))
            else:
                messagebox.showinfo("Sem duplicidades", "Nenhuma duplicidade encontrada.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo: {str(e)}")

# Configuração da interface Tkinter
root = tk.Tk()
root.title("Verificar Duplicidade de Datas de Vencimento")
root.geometry("400x200")

# Botão para selecionar o arquivo
botao_arquivo = tk.Button(root, text="Selecionar Arquivo Excel", command=selecionar_arquivo, height=2, width=30)
botao_arquivo.pack(pady=50)

# Iniciar a interface
root.mainloop()
