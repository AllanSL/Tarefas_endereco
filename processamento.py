from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side
from tkinter import messagebox
import os
from utils import resource_path

def verificar_arquivo_em_uso(caminho):
    """Verifica se o arquivo está em uso."""
    try:
        with open(caminho, 'r+'):
            pass
    except IOError:
        return True
    return False

def processar_arquivo(caminho_tarefas, progresso_callback=None):
    try:
        # Verifica se o arquivo de entrada está em uso
        if verificar_arquivo_em_uso(caminho_tarefas):
            messagebox.showerror("Erro", "O arquivo selecionado está em uso. Feche o arquivo e tente novamente.")
            return

        # Etapa 1: Leitura dos arquivos
        if progresso_callback: progresso_callback(10)
        df_tarefas = pd.read_excel(caminho_tarefas, dtype={'Produto': str})
        caminho_giro = resource_path('GIRO.xlsx')
        df_giro = pd.read_excel(caminho_giro, dtype={'Produto': str})

        # Etapa 2: Limpeza
        if progresso_callback: progresso_callback(25)
        df_tarefas['Produto'] = df_tarefas['Produto'].str.strip()
        df_giro['Produto'] = df_giro['Produto'].str.strip()

        # Etapa 3: Merge e preenchimento de GIRO
        if progresso_callback: progresso_callback(40)
        df = df_tarefas.merge(df_giro, on='Produto', how='left')
        df['GIRO'] = df['GIRO'].fillna('c2').str.strip().str.lower()

        # Etapa 4: Agrupamento e criação de ranking
        if progresso_callback: progresso_callback(60)
        agrupado = df.groupby('Local_Origem')
        ranking = agrupado.size().reset_index(name='# TAREFAS GERAL')
        ranking['# TAREFAS A1'] = agrupado.apply(lambda g: (g['GIRO'] == 'a1').sum()).values
        ranking['# TAREFAS A1 e A2'] = agrupado.apply(lambda g: g['GIRO'].isin(['a1', 'a2']).sum()).values
        ranking = ranking.rename(columns={'Local_Origem': 'Endereço'})
        ranking = ranking.sort_values(by='# TAREFAS GERAL', ascending=False)

        # Etapa 5: Salvando e formatando
        if progresso_callback: progresso_callback(80)
        data_atual = datetime.now().strftime('%d_%m_%y')
        caminho_saida = f'locais_tarefas_{data_atual}.xlsx'
        ranking.to_excel(caminho_saida, index=False)
        formatar_excel(caminho_saida)

        # Etapa 6: Abrindo e finalizando
        messagebox.showinfo("Concluído", "Processamento finalizado com sucesso!")
        if progresso_callback: progresso_callback(100)
        os.startfile(caminho_saida)

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