import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import logging
import os
from openpyxl import load_workbook

# Configuração do logger
logging.basicConfig(filename='email_metrics.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Função para calcular as métricas
def calcular_metricas(df):
    try:
        # Verificar se as colunas necessárias existem
        colunas_necessarias = ['Enviados', 'Abertos', 'Cliques', 'Erros', 'Descadastros']
        for coluna in colunas_necessarias:
            if (coluna not in df.columns) or (not pd.api.types.is_numeric_dtype(df[coluna])):
                messagebox.showerror('Erro', f'A coluna necessária "{coluna}" não foi encontrada na planilha ou não possui valores numéricos.')
                return None

        # Calcular métricas
        df['Taxa de Abertura (%)'] = (df['Abertos'] / df['Enviados']) * 100
        df['Taxa de Cliques (%)'] = (df['Cliques'] / df['Enviados']) * 100
        df['Taxa de Descadastros (%)'] = (df['Descadastros'] / df['Enviados']) * 100
        df['Feedback Score'] = 1 - ((df['Descadastros'] + df['Erros']) / df['Enviados'])

        # Adicionar coluna de relatório de desempenho
        df['Relatório de Desempenho'] = df.apply(gerar_relatorio_desempenho, axis=1)

        return df
    except Exception as e:
        logging.error(f'Erro ao calcular métricas: {str(e)}')
        messagebox.showerror('Erro', f'Ocorreu um erro ao calcular as métricas: {str(e)}')
        return None

# Função para gerar relatório de desempenho
def gerar_relatorio_desempenho(row):
    status = []

    # Análise da Taxa de Abertura
    if row['Taxa de Abertura (%)'] > 25:
        status.append("Taxa de Abertura: Excelente")
    elif row['Taxa de Abertura (%)'] > 15:
        status.append("Taxa de Abertura: Boa")
    else:
        status.append("Taxa de Abertura: Precisa de Melhorias")

    # Análise da Taxa de Cliques
    if row['Taxa de Cliques (%)'] > 5:
        status.append("Taxa de Cliques: Excelente")
    elif row['Taxa de Cliques (%)'] > 2:
        status.append("Taxa de Cliques: Boa")
    else:
        status.append("Taxa de Cliques: Precisa de Melhorias")

    # Análise da Taxa de Descadastros
    if row['Taxa de Descadastros (%)'] < 0.5:
        status.append("Taxa de Descadastros: Excelente")
    elif row['Taxa de Descadastros (%)'] < 1:
        status.append("Taxa de Descadastros: Boa")
    else:
        status.append("Taxa de Descadastros: Precisa de Melhorias")

    # Análise do Feedback Score
    if row['Feedback Score'] > 0.95:
        status.append("Feedback Score: Excelente")
    elif row['Feedback Score'] > 0.9:
        status.append("Feedback Score: Boa")
    else:
        status.append("Feedback Score: Precisa de Melhorias")

    return "; ".join(status)

# Função para ajustar a largura das colunas
def ajustar_largura_colunas(nome_arquivo):
    wb = load_workbook(nome_arquivo)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    wb.save(nome_arquivo)

# Função para processar a planilha
def processar_planilha(nome_arquivo):
    try:
        # Carregar planilha Excel ou CSV
        if nome_arquivo.lower().endswith('.csv'):
            df = pd.read_csv(nome_arquivo)
        elif nome_arquivo.lower().endswith('.xlsx') or nome_arquivo.lower().endswith('.xls'):
            df = pd.read_excel(nome_arquivo)
        else:
            messagebox.showerror('Erro', 'Formato de arquivo não suportado. Por favor, selecione um arquivo CSV (.csv) ou Excel (.xlsx, .xls).')
            return

        # Calcular as métricas
        df_com_metricas = calcular_metricas(df)

        if df_com_metricas is not None:
            # Salvar como uma nova planilha Excel
            novo_nome_arquivo = os.path.splitext(nome_arquivo)[0] + '_com_metricas.xlsx'
            df_com_metricas.to_excel(novo_nome_arquivo, index=False)
            
            # Ajustar a largura das colunas
            ajustar_largura_colunas(novo_nome_arquivo)
            
            messagebox.showinfo('Sucesso', f'Planilha com métricas salva em {novo_nome_arquivo}')
    except FileNotFoundError:
        logging.error(f'Arquivo não encontrado: {nome_arquivo}')
        messagebox.showerror('Erro', f'Arquivo não encontrado: {nome_arquivo}')
    except Exception as e:
        logging.error(f'Erro ao processar planilha: {str(e)}')
        messagebox.showerror('Erro', f'Ocorreu um erro ao processar a planilha: {str(e)}')

# Função para lidar com o clique do botão
def selecionar_arquivo():
    arquivo_selecionado = filedialog.askopenfilename(filetypes=[('Arquivos CSV', '*.csv'), ('Planilhas Excel', '*.xlsx;*.xls')])
    if arquivo_selecionado:
        processar_planilha(arquivo_selecionado)

# Configuração da interface gráfica
root = tk.Tk()
root.title('Análise de Métricas de Email')

# Criar e posicionar os widgets
label_instrucoes = tk.Label(root, text='Selecione o arquivo CSV ou Excel para processar:')
label_instrucoes.pack(pady=10)

btn_selecionar_arquivo = tk.Button(root, text='Selecionar Arquivo', command=selecionar_arquivo)
btn_selecionar_arquivo.pack(pady=10)

# Executar o loop principal da interface gráfica
root.mainloop()
