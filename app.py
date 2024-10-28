import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Função para carregar o arquivo Excel
def carregar_arquivo():
    filepath = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if filepath:
        try:
            # Pega a última página do Excel
            sheets = pd.read_excel(filepath, sheet_name=None)
            ultima_planilha = list(sheets.keys())[-1]
            df = sheets[ultima_planilha]
            messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")
            return df
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {str(e)}")
            return None
    return None

# Função para salvar o DataFrame modificado em um novo arquivo Excel
def salvar_arquivo(df):
    save_path = filedialog.asksaveasfilename(
        title="Salvar Arquivo",
        defaultextension=".xlsx",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if save_path:
        try:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o arquivo: {str(e)}")

# Função para modelar os dados
def modelar_dados(df):
    if df is not None:
        df_modelado = df.groupby('CENTRO DE CUSTO', as_index=False)['TOTAL'].sum()
        df_modelado['Total das obras'] = df_modelado['TOTAL'].sum()
        return df_modelado
    else:
        messagebox.showwarning("Aviso", "Nenhum dado carregado para modelar.")
        return None

# Função principal para a interface
def iniciar_interface():
    root = tk.Tk()
    root.title("Automatizador de Modelagem de Excel")
    root.geometry("300x150")
    
    df = None  # Variável para armazenar os dados do Excel
    
    def carregar_dados():
        nonlocal df
        df = carregar_arquivo()

    def processar_dados():
        nonlocal df
        if df is not None:
            df = modelar_dados(df)
            messagebox.showinfo("Sucesso", "Dados modelados com sucesso!")
        else:
            messagebox.showwarning("Aviso", "Nenhum arquivo foi carregado.")

    def salvar_dados():
        nonlocal df
        if df is not None:
            salvar_arquivo(df)
        else:
            messagebox.showwarning("Aviso", "Nenhum dado para salvar.")
    
    # Botão para carregar o arquivo Excel
    btn_carregar = tk.Button(root, text="Carregar Excel", command=carregar_dados)
    btn_carregar.pack(pady=10)
    
    # Botão para modelar os dados
    btn_modelar = tk.Button(root, text="Modelar Dados", command=processar_dados)
    btn_modelar.pack(pady=10)
    
    # Botão para salvar o arquivo modificado
    btn_salvar = tk.Button(root, text="Salvar Arquivo", command=salvar_dados)
    btn_salvar.pack(pady=10)
    
    root.mainloop()

# Iniciar a interface gráfica
iniciar_interface()