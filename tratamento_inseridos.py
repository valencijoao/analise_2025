import pandas as pd
import openpyxl

df = pd.read_excel('tecsystems_2025.xlsm' , sheet_name= 'Inseridos' )




df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
# 1. Limpeza profunda de espaços (incluindo espaços internos duplos e invisíveis)
df['Cliente'] = df['Cliente'].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()

# 2. Padronização (Usando apenas um pedaço do nome para não errar)
df.loc[df['Cliente'].str.contains('AIR LIQUIDE', case=False, na=False), 'Cliente'] = 'Air Liquide'

map_clientes = {
    'PORTO SEGURO COMPANHIA DE SEGUROS GERAIS':'Porto Seguro',
    'MANUPA COMERCIO EXPORTACAO IMPORTACAO DE EQUIPAMENTOS E VEIC': 'Veículos',
    'EDUCANTES PLATAFORMA ONLINE EDUCACIONAL LTDA SP':'Educantes',
    'HEXIS CIENTIFICA LTDA':'Hexis',
    'COMAZI TRATORES E MAQUINAS LTDA':'Comazi',
    'COVEZI CAMINHOES E ONIBUS LTDA - TOCANTINS':'Covezi',
    'MEGAFER COMERCIO DE FERRO E ACO LTDA - ME':'MG3',
    'LICITEC COMERCIAL LTDA':'Licitec',
    'L.P.M. TELEINFORMATICA LTDA':'LPM',
    'GCT - GERENCIAMENTO E CONTROLE DE TRANSITO S/A':'GCT',
    'WATSON-MARLOW BREDEL INDUSTRIA E COMERCIO DE BOMBAS LTDA':'Watson-Marlow',
    'CABALA SOLUCOES GOVERNAMENTAIS LTDA':'Cabala',
    'DANFOSS DO BRASIL INDUSTRIA E COMERCIO LTDA':'Danfoss',
    'SERVICOS AEREOS INDUSTRIAIS ESPECIALIZADOS SAI LTDA':'Sai Brasil',
    'PHD SISTEMAS DE ENERGIA INDUSTRIA, COMERCIO, IMPORTACAO E EX':'PHD',
    'POTTENCIAL VEICULOS ESPECIAIS LTDA':'Pottencial',
}

df['Cliente'] = df['Cliente'].replace(map_clientes)
print(df.head())