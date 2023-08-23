import pandas as pd
import pandera as pa
import os
from glob import glob

def espera(n, men):
    from time import sleep
    print('-='*40)
    print(men)
    print('-='*40)
    sleep(n)


pasta_atual = os.getcwd() + '\\*.xlsx'
folders = sorted(glob(pasta_atual))
if len(folders) == 0:
    pasta_atual = os.getcwd() + '\\*.xls'
    folders = sorted(glob(pasta_atual))

try:
    df = pd.concat((pd.read_excel(cont) for cont in folders), ignore_index=True)
except:
    espera(2, 'ERRO! Arquivo corrompido ou formato invalido')

df['Válido'] = 'S'

schema = pa.DataFrameSchema(
  columns = {
    "Vln_a_mean": pa.Column(pa.Float, nullable=False),
    "Vln_b_mean": pa.Column(pa.Float, nullable=False),
    "Vln_c_mean": pa.Column(pa.Float, nullable=False),
    "Ia_mean": pa.Column(pa.Float, nullable=False),
    "Ib_mean": pa.Column(pa.Float, nullable=False),
    "Ic_mean": pa.Column(pa.Float, nullable=False),
    "Active_Power": pa.Column(pa.Float, nullable=False),
    "Frequency": pa.Column(pa.Float, coerce=pa.Float, nullable=False), 
  }
)

try:
    schema.validate(df)
except:
    espera(3, 'ERRO! Verifique se os arquivos contém valores vazios, ou Nome das colunas não encotrado. ')

filtro_01_r0 = df['Vln_a_mean'] == 0 
filtro_02_r0 = df['Vln_b_mean'] == 0
filtro_03_r0 = df['Vln_c_mean'] == 0
filtro_04_r0 = df['Ia_mean'] == 0
filtro_05_r0 = df['Ib_mean'] == 0
filtro_06_r0 = df['Ic_mean'] == 0
filtro_07_r0 = df['Frequency'] == 0
filtro_08_r0 = df['Active_Power'] == 0
filtro_09_r0 = df['Active_Energy'] == 0	
filtro_10_r0 = df['Reactive_Power'] == 0	
filtro_11_r0 = df['Reactive_Energy'] == 0	    

# REGRA 00
r0 = df.loc[filtro_01_r0 & filtro_02_r0 & filtro_03_r0 & filtro_04_r0 & filtro_05_r0 & filtro_06_r0 & filtro_07_r0 & filtro_08_r0 & filtro_09_r0 & filtro_10_r0 & filtro_11_r0]

if len(r0) > 0:
    df.loc[filtro_01_r0 & filtro_02_r0 & filtro_03_r0 & filtro_04_r0 & filtro_05_r0 & filtro_06_r0 & filtro_07_r0 & filtro_08_r0 & filtro_09_r0 & filtro_10_r0 & filtro_11_r0, ['Válido', 'Regra 00']] = ['N', 'Grandezas iguais a zero']

# REGRA 01

filtro_00_r1 = (df['Vln_a_mean'] == 0) & (df['Ia_mean'] > 0)
filtro_01_r1 = (df['Vln_b_mean'] == 0) & (df['Ib_mean'] > 0)
filtro_02_r1 = (df['Vln_c_mean'] == 0) & (df['Ic_mean'] > 0)


r1 = df.loc[filtro_00_r1 | filtro_01_r1 | filtro_02_r1]

if len(r1) > 0:
    df.loc[filtro_00_r1 | filtro_01_r1 | filtro_02_r1, ['Válido', 'Regra 01']] = ['N', 'Tensão A, B ou C igual zero e Corrente de fase correspondente maior que zero']

# REGRA 02

filtro_01_r2 = df['Vln_a_mean'] < 0 
filtro_02_r2 = df['Vln_b_mean'] < 0
filtro_03_r2 = df['Vln_c_mean'] < 0
filtro_04_r2 = df['Ia_mean'] < 0
filtro_05_r2 = df['Ib_mean'] < 0
filtro_06_r2 = df['Ic_mean'] < 0
filtro_07_r2 = df['Active_Power'] < 0
filtro_08_r2 = df['Frequency'] < 0
filtro_09_r2 = df['Active_Energy'] < 0	
filtro_10_r2 = df['Reactive_Power'] < 0	
filtro_11_r2 = df['Reactive_Energy'] < 0	    
    

r2 = df.loc[filtro_01_r2 | filtro_02_r2 | filtro_03_r2 | filtro_04_r2 | filtro_05_r2 | filtro_06_r2 | filtro_07_r2 | filtro_08_r2 | filtro_09_r2 | filtro_10_r2 | filtro_11_r2]

if len(r2) > 0:
    df.loc[filtro_01_r2 | filtro_02_r2 | filtro_03_r2 | filtro_04_r2 | filtro_05_r2 | filtro_06_r2 | filtro_07_r2 | filtro_08_r2 | filtro_09_r2 | filtro_10_r2 | filtro_11_r2, ['Válido', 'Regra 02']] = ['N', 'Grandezas com valores negativos']

# REGRA 04
filtro_00_r4 = df['Vln_a_mean'] == 0
filtro_01_r4 = df['Vln_b_mean'] == 0
filtro_02_r4 = df['Vln_c_mean'] == 0
filtro_03_r4 = df['Active_Power'] != 0

r4 = df.loc[(filtro_00_r4 | filtro_01_r4 | filtro_02_r4) & filtro_03_r4]

if len(r4) > 0:
    df.loc[(filtro_00_r4 | filtro_01_r4 | filtro_02_r4) & filtro_03_r4, ['Válido', 'Regra 04']] = ['N', 'Tensão A, B e C igual a zero, e Potência Ativa diferente de zero']
# REGRA 05
filtro_00_r5 = df['Ia_mean'] == 0
filtro_01_r5 = df['Ib_mean'] == 0
filtro_02_r5 = df['Ic_mean'] == 0
filtro_03_r5 = df['Active_Power'] != 0

r5 = df.loc[(filtro_00_r5 | filtro_01_r5 | filtro_02_r5) & filtro_03_r5]

if len(r5) > 0:
    df.loc[(filtro_00_r5 | filtro_01_r5 | filtro_02_r5) & filtro_03_r5, ['Válido', 'Regra 05']] = ['N', 'Corrente A, B e C igual a zero e Potência Ativa diferente de zero']
# REGRA 06
filtro_00_r6 = df['Active_Power'] == 0
filtro_01_r6 = df['Vln_a_mean'] != 0
filtro_02_r6 = df['Vln_b_mean'] != 0
filtro_03_r6 = df['Vln_c_mean'] != 0
filtro_04_r6 = df['Ia_mean'] != 0
filtro_05_r6 = df['Ib_mean'] != 0
filtro_06_r6 = df['Ic_mean'] != 0
filtro_07_r6 = df['Frequency'] != 0
filtro_08_r6 = df['Active_Energy'] < 0	
filtro_09_r6 = df['Reactive_Power'] < 0	
filtro_10_r6 = df['Reactive_Energy'] < 0	    




r6 = df.loc[filtro_00_r6 & (filtro_01_r6 | filtro_02_r6 | filtro_03_r6 | filtro_04_r6 | filtro_05_r6 | filtro_06_r6 | filtro_08_r6 | filtro_09_r6 | filtro_10_r6)]

if len(r6) > 0:
    df.loc[filtro_00_r6 & (filtro_01_r6 | filtro_02_r6 | filtro_03_r6 | filtro_04_r6 | filtro_05_r6 | filtro_06_r6 | filtro_08_r6 | filtro_09_r6 | filtro_10_r6), ['Válido', 'Regra 06']] = ['N', 'Potência Ativa igual a zero e todas as outras grandezas deferente de zero']
# REGRA 07
filtro_00_r7 = df['Frequency'] == 0
filtro_01_r7 = df['Vln_a_mean'] != 0
filtro_02_r7 = df['Vln_b_mean'] != 0 
filtro_03_r7 = df['Vln_c_mean'] != 0

r6 = df.loc[filtro_00_r7 & (filtro_01_r7 | filtro_02_r7 | filtro_03_r7)]

if len(r6) > 0:
  df.loc[filtro_00_r7 & (filtro_01_r7 | filtro_02_r7 | filtro_03_r7), ['Válido', 'Regra 07']] = ['N', 'Frequência igual a zero e qualquer umas das tensões diferente de zero']
# REGRA 08
filtro_00_r8 = df['Frequency'] != 0
filtro_01_r8 = df['Vln_a_mean'] == 0
filtro_02_r8 = df['Vln_b_mean'] == 0
filtro_03_r8 = df['Vln_c_mean'] == 0


r8 = df.loc[filtro_00_r8 & (filtro_01_r8 | filtro_02_r8 | filtro_03_r8)]

if len(r8) > 0:
    df.loc[filtro_00_r8 & (filtro_01_r8 | filtro_02_r8 | filtro_03_r8), ['Válido', 'Regra 08']] = ['N', 'Frequência diferente de zero e qualquer uma das tensões iguais a zero']

# REGRA 09
filtro_00_r9 = df['Frequency'] == 0 
filtro_01_r9 = df['Vln_a_mean'] == 0
filtro_02_r9 = df['Vln_b_mean'] == 0
filtro_03_r9 = df['Vln_c_mean'] == 0
filtro_04_r9 = df['Ia_mean'] != 0
filtro_05_r9 = df['Ib_mean'] != 0
filtro_06_r9 = df['Ic_mean'] != 0

r9 = df.loc[(filtro_00_r9 & filtro_01_r9 & filtro_02_r9 & filtro_03_r9) & (filtro_04_r9 & filtro_05_r9 & filtro_06_r9)]

if len(r9) > 0:
    df.loc[(filtro_00_r9 & filtro_01_r9 & filtro_02_r9 & filtro_03_r9) & (filtro_04_r9 & filtro_05_r9 & filtro_06_r9), ['Válido', 'Regra 09']] = ['N', 'Frequência e tensões iguais a zero e todas as correntes diferentes de zero']

#EXPORTANDO ERROS
erros = df[df['Válido'] == 'N']

if len(erros) > 0:
    erros.to_excel('Memoria-Invalida.xlsx', index='False')
    espera(3, 'CONCLUIDO COM SUCESSO!, as ocorrências inválidas estão em "Mémorias inválidas"')
else: 
    espera(3, 'Não foram encontrados ocorrências inválidas')