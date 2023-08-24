import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import openpyxl.styles

src_folder = os.getcwd() + '\\'
extensions = ('.xlsx', '.xls')
folders = [folder for folder in os.listdir(src_folder) if os.path.splitext(folder)[-1] in extensions]


print('Selecione o arquivo da previa')
for i, v in enumerate(folders):
    print(f'{i+1} - {v}')

resp = int(input('Digite o número do arquivo: '))

# ler o arquivo
df = pd.read_excel(folders[resp-1], header=4)

df = df[df['Tipo'].str.contains('CCEAR')]


# ajustar as colunas com date
df['Vencimento Parcela'] = df['Vencimento Parcela'].dt.date
df['Competência Processo'] = df['Competência Processo'].dt.date
df['Data Emissão NF'] = df['Data Emissão NF'].dt.date

df.to_excel('PREVIA CCEAR.xlsx', index=False)

# separar phanilhas CCEAR
bradesco_df = df.loc[(df['Banco'] == 'Bradesco')].copy()
bb_df = df.loc[(df['Banco'] == 'Banco Brasil')].copy()

# remover linhas que não contenham "CCEAR" na coluna tipo
bradesco_df = bradesco_df.loc[bradesco_df['Tipo'].str.contains('CCEAR')]
bb_df = bb_df.loc[bb_df['Tipo'].str.contains('CCEAR')]

# exportar arquivos
bradesco_df.to_excel('CCEAR - BRA.xlsx', index=False)
bb_df.to_excel('CCEAR - BB.xlsx', index=False)

arquivos = ["CCEAR - BB.xlsx", "CCEAR - BRA.xlsx", "PREVIA CCEAR.xlsx"]

# ajustar o cabeçalho e adicionar bordas
for i in arquivos:
    book = load_workbook(i)
    sheet = book.active

    # definir a altura da linha do cabeçalho
    sheet.row_dimensions[1].height = 30

    # ajustar a largura da coluna para o tamanho do conteúdo
    for col in range(1, sheet.max_column + 1):
        column_letter = get_column_letter(col)
        max_length = 0
        for cell in sheet[column_letter]:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception as e:
                pass
        adjusted_width = (max_length + 2) * 1.2 # ajustar o tamanho da coluna para acomodar o conteúdo
        sheet.column_dimensions[column_letter].width = adjusted_width

    # alinhar o cabeçalho no centro
    for cell in sheet[1]:
        cell.alignment = cell.alignment.copy(horizontal='center', vertical='center')

    # Adicionar bordas em todas as bordas das células
    for row in sheet.iter_rows():
        for cell in row:
            cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(border_style='thin', color='000000'),
                                                 right=openpyxl.styles.Side(border_style='thin', color='000000'),
                                                 top=openpyxl.styles.Side(border_style='thin', color='000000'),
                                                 bottom=openpyxl.styles.Side(border_style='thin', color='000000'))
    # salvar as alterações na planilha
    book.save(i)