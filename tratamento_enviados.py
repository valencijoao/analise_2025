import pandas as pd
import openpyxl

df = pd.read_excel('tecsystems_2025.xlsm', sheet_name='Enviados')

arquivo_saida = 'enviados_tratados.xlsx'
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
df = df[df['Portal'] == 'ComprasNet'].reset_index(drop=True)

map_clientes = {
    'Humberto':'Veículos',
    'Mercedes':'Veículos',
    'Divena':'Veículos',
    'Toriba':'Veículos',
    'Obras - Pequeno Porte':'Construtoras',
    'Granero':'Transportadoras',

}

df['Cliente'] = df['Cliente'].replace(map_clientes)

partes_processo = df['Arquivo'].str.split('_')

df['Data'] = partes_processo.str[0]
df['UASG'] = partes_processo.str[2].str.replace('.pdf','', regex=False)

df['Data'] = df['Data'] = pd.to_datetime(df['Data'], format='%d%m%Y', errors='coerce').dt.date

df = df.rename(columns={
    'Portal': 'Site',
    'ID':'Licitação'
})

nova_ordem = ['Data','Licitação','Site','UASG','Cliente']
df_tratado = df[nova_ordem]

df_tratado.to_excel(arquivo_saida, index=False, engine='openpyxl')
print(f"Arquivo salvo com sucesso:{arquivo_saida}")