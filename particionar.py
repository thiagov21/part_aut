import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment

# Caminhos
planilha_base = r'\\mao-fs01.technos.local\Arquivos\Departamental\PÓS VENDA\Laboratorio\COBRANÇA POSTOS Thiago V\automação\FILIAL\FILIAL.xlsx'
caminho_pasta_os_filial = r'\\mao-fs01.technos.local\Arquivos\Departamental\PÓS VENDA\Laboratorio\COBRANÇA POSTOS Thiago V\automação\FILIAL\FILIAL PART'

# Ler a planilha base
df = pd.read_excel(planilha_base)

# Obter a lista de postos únicos
filial_unica = df['Filial'].unique()

# Processar cada posto
for filial in filial_unica:
    filial_df = df[df['Filial'] == filial]
    
    # Verificar e criar diretório, se necessário
    if not os.path.exists(caminho_pasta_os_filial):
        os.makedirs(caminho_pasta_os_filial)
    
    # Criar o arquivo Excel separado para o posto
    arquivo_filial = os.path.join(caminho_pasta_os_filial, f'{filial}.xlsx')
    filial_df.to_excel(arquivo_filial, index=False)
    
    # Carregar a planilha criada com openpyxl
    wb = load_workbook(arquivo_filial)
    ws = wb.active
    
    # Centralizar o conteúdo das células
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Ajustar automaticamente a largura das colunas
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Letra da coluna (A, B, C, etc.)
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        # Adicionar margem ao comprimento e ajustar a largura
        adjusted_width = max_length + 2
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Adicionar formatação de tabela
    tab = Table(displayName=f"Table_{filial}", ref=ws.dimensions)
    
    # Estilo da tabela
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    ws.add_table(tab)
    
    # Salvar o arquivo com a formatação
    wb.save(arquivo_filial)
    
print('Processo concluído e planilhas formatadas com colunas ajustadas.')
