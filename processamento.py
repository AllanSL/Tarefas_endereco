from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from tkinter import messagebox
import os
from utils import resource_path

def processar_arquivo(caminho_tarefas):
    try:
        # Lê os arquivos
        df_tarefas = pd.read_excel(caminho_tarefas, dtype={'Produto': str})
        caminho_giro = resource_path('GIRO.xlsx')
        df_giro = pd.read_excel(caminho_giro, dtype={'Produto': str})

        # Limpa espaços
        df_tarefas['Produto'] = df_tarefas['Produto'].str.strip()
        df_giro['Produto'] = df_giro['Produto'].str.strip()

        # Merge e preenche GIRO
        df = df_tarefas.merge(df_giro, on='Produto', how='left')
        df['GIRO'] = df['GIRO'].fillna('c2').str.strip().str.lower()

        # Agrupa por Local_Origem
        agrupado = df.groupby('Local_Origem')

        # Cria ranking
        ranking = agrupado.size().reset_index(name='# TAREFAS GERAL')
        ranking['# TAREFAS A1'] = agrupado.apply(lambda g: (g['GIRO'] == 'a1').sum()).values
        ranking['# TAREFAS A1 e A2'] = agrupado.apply(lambda g: g['GIRO'].isin(['a1', 'a2']).sum()).values
        ranking = ranking.rename(columns={'Local_Origem': 'Endereço'})
        ranking = ranking.sort_values(by='# TAREFAS GERAL', ascending=False)

        # Salva arquivo Excel com a data no nome
        data_atual = datetime.now().strftime('%d_%m_%y')
        caminho_saida = f'locais_tarefas_{data_atual}.xlsx'
        ranking.to_excel(caminho_saida, index=False)

        # Formatação do Excel
        formatar_excel(caminho_saida)

        # Abre o arquivo gerado
        os.startfile(caminho_saida)
        messagebox.showinfo("Sucesso", f"Arquivo gerado e aberto com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

def formatar_excel(caminho_saida):
    wb = load_workbook(caminho_saida)
    ws = wb.active

    # Adiciona autofiltro
    ws.auto_filter.ref = ws.dimensions

    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    thin_border_bottom = Border(bottom=Side(style='thin'))

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for i, cell in enumerate(col):
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
            if i == 0:
                cell.font = bold_font
                cell.border = thin_border_bottom
            cell.alignment = center_align
        ws.column_dimensions[col_letter].width = max_length + 6

    wb.save(caminho_saida)