import pandas as pd
import pandera as pa
import os
from glob import glob
from time import sleep

def espera(n, men):
    print('-='*40)
    print(men)
    sleep(n)

pasta_atual = os.getcwd() + '\\*.xlsx'
folders = sorted([f for f in glob(pasta_atual) if 'Código Medidores' not in f])
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
    "Data e Hora da Medição": pa.Column(pa.DateTime, nullable=False),
    "Tensão Fase A Consistida (kV)": pa.Column(pa.Float, nullable=False),
    "Tensao Fase B Consistida(kV)": pa.Column(pa.Float, nullable=False),
    "Tensao Fase C Consistida (kV)": pa.Column(pa.Float, nullable=False),
    "Corrente Fase A Consistida (A)": pa.Column(pa.Float, nullable=False),
    "Corrente Fase B Consistida (A)": pa.Column(pa.Float, nullable=False),
    "Corrente Fase C Consistida (A)": pa.Column(pa.Float, nullable=False),
    "Potencia Ativa Consistida (kW)": pa.Column(pa.Float, nullable=False),
    "Frequencia Consistida (Hz)": pa.Column(pa.Float, coerce=pa.Float, nullable=False),
    
  }
)
try:
    schema.validate(df)
except:
    espera(3, 'ERRO! Verifique se os arquivos contém valores vazios, ou Nome das colunas não encotrado. ')


filtro_01_r0 = df['Tensão Fase A Consistida (kV)'] == 0 
filtro_02_r0 = df['Tensao Fase B Consistida(kV)'] == 0
filtro_03_r0 = df['Tensao Fase C Consistida (kV)'] == 0
filtro_04_r0 = df['Corrente Fase A Consistida (A)'] == 0
filtro_05_r0 = df['Corrente Fase B Consistida (A)'] == 0
filtro_06_r0 = df['Corrente Fase C Consistida (A)'] == 0
filtro_07_r0 = df['Potencia Ativa Consistida (kW)'] == 0
filtro_08_r0 = df['Potência Reativa Consistida (kVAr)'] == 0
filtro_09_r0 = df['Energia Ativa Consistida (kWh)'] == 0
filtro_10_r0 = df['Energia Reativa Consistida (kVArh)'] == 0
filtro_11_r0 = df['Frequencia Consistida (Hz)'] == 0
filtro_12_r0 = df['Potencia Ativa (kW)'] == 0	
filtro_13_r0 = df['Energia Ativa (kWh)'] == 0	
filtro_14_r0 = df['Potencia Reativa (kVAr)'] == 0	
filtro_15_r0 = df['Energia Reativa (kVArh)'] == 0	    
filtro_16_r0 = df['Tensao Fase A (kV)'] == 0	    
filtro_17_r0 = df['Tensao Fase B (kV)'] == 0	    	
filtro_18_r0 = df['Tensao Fase C (kV)'] == 0	    
filtro_19_r0 = df['Corrente Fase A (A)'] == 0	
filtro_20_r0 = df['Corrente Fase B (A)'] == 0	
filtro_21_r0 = df['Corrente Fase C (A)'] == 0	    	
filtro_22_r0 = df['Frequencia (Hz)'] == 0	

r0 = df.loc[filtro_01_r0 & filtro_02_r0 & filtro_03_r0 & filtro_04_r0 & filtro_05_r0 & filtro_06_r0 & filtro_07_r0 & filtro_08_r0 & filtro_09_r0 & filtro_10_r0 & filtro_11_r0 & filtro_12_r0 & filtro_13_r0 & filtro_14_r0 & filtro_15_r0 & filtro_16_r0 & filtro_17_r0 & filtro_18_r0 & filtro_19_r0 & filtro_20_r0 & filtro_21_r0 & filtro_22_r0]

if len(r0) > 0:
    df.loc[filtro_01_r0 & filtro_02_r0 & filtro_03_r0 & filtro_04_r0 & filtro_05_r0 & filtro_06_r0 & filtro_07_r0 & filtro_08_r0 & filtro_09_r0 & filtro_10_r0 & filtro_11_r0 & filtro_12_r0 & filtro_13_r0 & filtro_14_r0 & filtro_15_r0 & filtro_16_r0 & filtro_17_r0 & filtro_18_r0 & filtro_19_r0 & filtro_20_r0 & filtro_21_r0 & filtro_22_r0, ['Válido', 'Regra 03']] = ['N', 'Grandezas iguais a zero']


filtro_00_r1 = (df['Tensão Fase A Consistida (kV)'] == 0) & (df['Corrente Fase A Consistida (A)'] > 0)
filtro_01_r1 = (df['Tensao Fase B Consistida(kV)'] == 0) & (df['Corrente Fase B Consistida (A)'] > 0)
filtro_02_r1 = (df['Tensao Fase C Consistida (kV)'] == 0) & (df['Corrente Fase C Consistida (A)'] > 0)


r1 = df.loc[filtro_00_r1 | filtro_01_r1 | filtro_02_r1]

if len(r1) > 0:
    df.loc[filtro_00_r1 | filtro_01_r1 | filtro_02_r1, ['Válido', 'Regra 01']] = ['N', 'Tensão A, B ou C igual zero e Corrente de fase correspondente maior que zero']

filtro_01_r2 = df['Tensão Fase A Consistida (kV)'] < 0 
filtro_02_r2 = df['Tensao Fase B Consistida(kV)'] < 0
filtro_03_r2 = df['Tensao Fase C Consistida (kV)'] < 0
filtro_04_r2 = df['Corrente Fase A Consistida (A)'] < 0
filtro_05_r2 = df['Corrente Fase B Consistida (A)'] < 0
filtro_06_r2 = df['Corrente Fase C Consistida (A)'] < 0
filtro_07_r2 = df['Potencia Ativa Consistida (kW)'] < 0
filtro_08_r2 = df['Potência Reativa Consistida (kVAr)'] < 0
filtro_09_r2 = df['Energia Ativa Consistida (kWh)'] < 0
filtro_10_r2 = df['Energia Reativa Consistida (kVArh)'] < 0
filtro_11_r2 = df['Frequencia Consistida (Hz)'] < 0
filtro_12_r2 = df['Potencia Ativa (kW)'] < 0	
filtro_13_r2 = df['Energia Ativa (kWh)'] < 0	
filtro_14_r2 = df['Potencia Reativa (kVAr)'] < 0	
filtro_15_r2 = df['Energia Reativa (kVArh)'] < 0	    
filtro_16_r2 = df['Tensao Fase A (kV)'] < 0	    
filtro_17_r2 = df['Tensao Fase B (kV)'] < 0	    	
filtro_18_r2 = df['Tensao Fase C (kV)'] < 0	    
filtro_19_r2 = df['Corrente Fase A (A)'] < 0	
filtro_20_r2 = df['Corrente Fase B (A)'] < 0	
filtro_21_r2 = df['Corrente Fase C (A)'] < 0	    	
filtro_22_r2 = df['Frequencia (Hz)'] < 0	
    



r2 = df.loc[filtro_01_r2 | filtro_02_r2 | filtro_03_r2 | filtro_04_r2 | filtro_05_r2 | filtro_06_r2 | filtro_07_r2 | filtro_08_r2 | filtro_09_r2 | filtro_10_r2 | filtro_11_r2 | filtro_12_r2 |
filtro_13_r2 | filtro_14_r2 | filtro_15_r2 | filtro_16_r2 | filtro_17_r2 | filtro_18_r2 | filtro_19_r2 | filtro_20_r2 | filtro_21_r2 | filtro_22_r2]

if len(r2) > 0:
    df.loc[filtro_01_r2 | filtro_02_r2 | filtro_03_r2 | filtro_04_r2 | filtro_05_r2 | filtro_06_r2 | filtro_07_r2 | filtro_08_r2 | filtro_09_r2 | filtro_10_r2 | filtro_11_r2 | filtro_12_r2 |
    filtro_13_r2 | filtro_14_r2 | filtro_15_r2 | filtro_16_r2 | filtro_17_r2 | filtro_18_r2 | filtro_19_r2 | filtro_20_r2 | filtro_21_r2 | filtro_22_r2, ['Válido', 'Regra 02']] = ['N', 'Grandezas com valores negativos']

filtro_00_r4 = df['Tensão Fase A Consistida (kV)'] == 0
filtro_01_r4 = df['Tensao Fase B Consistida(kV)'] == 0
filtro_02_r4 = df['Tensao Fase C Consistida (kV)'] == 0
filtro_03_r4 = df['Potencia Ativa Consistida (kW)'] != 0

r4 = df.loc[(filtro_00_r4 | filtro_01_r4 | filtro_02_r4) & filtro_03_r4]

if len(r4) > 0:
    df.loc[(filtro_00_r4 | filtro_01_r4 | filtro_02_r4) & filtro_03_r4, ['Válido', 'Regra 04']] = ['N', 'Tensão A, B e C igual a zero, e Potência Ativa diferente de zero']

filtro_00_r5 = df['Corrente Fase A Consistida (A)'] == 0
filtro_01_r5 = df['Corrente Fase B Consistida (A)'] == 0
filtro_02_r5 = df['Corrente Fase C Consistida (A)'] == 0
filtro_03_r5 = df['Potencia Ativa Consistida (kW)'] != 0

r5 = df.loc[(filtro_00_r5 | filtro_01_r5 | filtro_02_r5) & filtro_03_r5]

if len(r5) > 0:
    df.loc[(filtro_00_r5 | filtro_01_r5 | filtro_02_r5) & filtro_03_r5, ['Válido', 'Regra 05']] = ['N', 'Corrente A, B e C igual a zero e Potência Ativa diferente de zero']

filtro_00_r6 = df['Potencia Ativa Consistida (kW)'] == 0
filtro_01_r6 = df['Tensão Fase A Consistida (kV)'] != 0
filtro_02_r6 = df['Tensao Fase B Consistida(kV)'] != 0
filtro_03_r6 = df['Tensao Fase C Consistida (kV)'] != 0
filtro_04_r6 = df['Corrente Fase A Consistida (A)'] != 0
filtro_05_r6 = df['Corrente Fase B Consistida (A)'] != 0
filtro_06_r6 = df['Corrente Fase C Consistida (A)'] != 0
# filtro_07_r6 = df['Potencia Ativa Consistida (kW)'] != 0
filtro_08_r6 = df['Potência Reativa Consistida (kVAr)'] != 0
filtro_09_r6 = df['Energia Ativa Consistida (kWh)'] != 0
filtro_10_r6 = df['Energia Reativa Consistida (kVArh)'] != 0
filtro_11_r6 = df['Frequencia Consistida (Hz)'] != 0
filtro_12_r6 = df['Potencia Ativa (kW)'] < 0	
filtro_13_r6 = df['Energia Ativa (kWh)'] < 0	
filtro_14_r6 = df['Potencia Reativa (kVAr)'] < 0	
filtro_15_r6 = df['Energia Reativa (kVArh)'] < 0	    
filtro_16_r6 = df['Tensao Fase A (kV)'] < 0	    
filtro_17_r6 = df['Tensao Fase B (kV)'] < 0	    	
filtro_18_r6 = df['Tensao Fase C (kV)'] < 0	    
filtro_19_r6 = df['Corrente Fase A (A)'] < 0	
filtro_20_r6 = df['Corrente Fase B (A)'] < 0	
filtro_21_r6 = df['Corrente Fase C (A)'] < 0	    	
filtro_22_r6 = df['Frequencia (Hz)'] < 0	




r6 = df.loc[filtro_00_r6 & (filtro_01_r6 | filtro_02_r6 | filtro_03_r6 | filtro_04_r6 | filtro_05_r6 | filtro_06_r6 | filtro_08_r6 | filtro_09_r6 | filtro_10_r6 | filtro_11_r6 | filtro_12_r6 | filtro_13_r6 | filtro_14_r6 | filtro_15_r6 | filtro_16_r6 | filtro_17_r6 | filtro_18_r6 | filtro_19_r6 | filtro_20_r6 | filtro_21_r6 | filtro_22_r6 )]

if len(r6) > 0:
    df.loc[filtro_00_r6 & (filtro_01_r6 | filtro_02_r6 | filtro_03_r6 | filtro_04_r6 | filtro_05_r6 | filtro_06_r6 | filtro_08_r6 | filtro_09_r6 | filtro_10_r6 | filtro_11_r6 | filtro_12_r6 | filtro_13_r6 | filtro_14_r6 | filtro_15_r6 | filtro_16_r6 | filtro_17_r6 | filtro_18_r6 | filtro_19_r6 | filtro_20_r6 | filtro_21_r6 | filtro_22_r6 ), ['Válido', 'Regra 06']] = ['N', 'Potência Ativa igual a zero e todas as outras grandezas deferente de zero']

filtro_00_r7 = df['Frequencia Consistida (Hz)'] == 0
filtro_01_r7 = df['Tensão Fase A Consistida (kV)'] != 0
filtro_02_r7 = df['Tensao Fase B Consistida(kV)'] != 0 
filtro_03_r7 = df['Tensao Fase C Consistida (kV)'] != 0



r6 = df.loc[filtro_00_r7 & (filtro_01_r7 | filtro_02_r7 | filtro_03_r7)]

if len(r6) > 0:
   df.loc[filtro_00_r7 & (filtro_01_r7 | filtro_02_r7 | filtro_03_r7), ['Válido', 'Regra 07']] = ['N', 'Frequência igual a zero e qualquer umas das tensões diferente de zero']

filtro_00_r8 = df['Frequencia Consistida (Hz)'] != 0
filtro_01_r8 = df['Tensão Fase A Consistida (kV)'] == 0
filtro_02_r8 = df['Tensao Fase B Consistida(kV)'] == 0
filtro_03_r8 = df['Tensao Fase C Consistida (kV)'] == 0


r8 = df.loc[filtro_00_r8 & (filtro_01_r8 | filtro_02_r8 | filtro_03_r8)]

if len(r8) > 0:
    df.loc[filtro_00_r8 & (filtro_01_r8 | filtro_02_r8 | filtro_03_r8), ['Válido', 'Regra 08']] = ['N', 'Frequência diferente de zero e qualquer uma das tensões iguais a zero']

filtro_00_r9 = df['Frequencia Consistida (Hz)'] == 0 
filtro_01_r9 = df['Tensão Fase A Consistida (kV)'] == 0
filtro_02_r9 = df['Tensao Fase B Consistida(kV)'] == 0
filtro_03_r9 = df['Tensao Fase C Consistida (kV)'] == 0
filtro_04_r9 = df['Corrente Fase A Consistida (A)'] != 0
filtro_05_r9 = df['Corrente Fase B Consistida (A)'] != 0
filtro_06_r9 = df['Corrente Fase C Consistida (A)'] != 0

r9 = df.loc[(filtro_00_r9 & filtro_01_r9 & filtro_02_r9 & filtro_03_r9) & (filtro_04_r9 & filtro_05_r9 & filtro_06_r9)]

if len(r9) > 0:
    df.loc[(filtro_00_r9 & filtro_01_r9 & filtro_02_r9 & filtro_03_r9) & (filtro_04_r9 & filtro_05_r9 & filtro_06_r9), ['Válido', 'Regra 09']] = ['N', 'Frequência e tensões iguais a zero e todas as correntes diferentes de zero']
 

erros = df[df['Válido'] == 'N']
if len(erros) > 0:
    try:
        df2 = pd.read_excel('Código Medidores.xlsx')
    except:
        espera(3, 'Não foi encotrado o arquivo "Código Medidores"')
        
    erros.loc[:, 'Identificador'] = erros.apply(lambda row: df2.loc[df2['Cod. Medidor'] == row['Identificador'], 'Usina '].values[0] if 'Cod. Medidor' in df2.columns and df2.loc[df2['Cod. Medidor'] == row['Identificador'], 'Usina '].values.size > 0 else row['Identificador'], axis=1)

    erros.to_excel('Memoria-Invalida.xlsx', index='False')
    espera(3, 'CONCLUIDO COM SUCESSO!, as ocorrências inválidas estão em "Mémorias inválidas"')
else: 
    espera(3, 'Não foram encontrados ocorrências inválidas')
