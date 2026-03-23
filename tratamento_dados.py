import pandas as pd
import re
from rapidfuzz import fuzz


def limpar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = re.sub(r'[^0-9A-Z ]', '', valor)
    return valor


def gerar_possiveis_matches(arquivo_entrada, arquivo_saida):

    df_env = pd.read_excel(arquivo_entrada, sheet_name="Enviados")
    df_ins = pd.read_excel(arquivo_entrada, sheet_name="Inseridos")

    # limpar colunas
    df_env.columns = df_env.columns.str.strip().str.replace('\xa0', '')
    df_ins.columns = df_ins.columns.str.strip().str.replace('\xa0', '')

    # padronizar nomes
    df_ins = df_ins.rename(columns={'Licitação': 'ID', 'Site': 'Portal'})

    # limpar campos principais
    df_env['CLIENTE_LIMPO'] = df_env['Cliente'].apply(limpar_texto)
    df_ins['CLIENTE_LIMPO'] = df_ins['Cliente'].apply(limpar_texto)

    df_env['PORTAL_LIMPO'] = df_env['Portal'].apply(limpar_texto)
    df_ins['PORTAL_LIMPO'] = df_ins['Portal'].apply(limpar_texto)

    # resultado
    resultados = []

    # comparação cruzada (controlada)
    for i, env_row in df_env.iterrows():

        for j, ins_row in df_ins.iterrows():

            # similaridade cliente
            score_cliente = fuzz.ratio(
                env_row['CLIENTE_LIMPO'],
                ins_row['CLIENTE_LIMPO']
            )

            # similaridade portal
            score_portal = fuzz.ratio(
                env_row['PORTAL_LIMPO'],
                ins_row['PORTAL_LIMPO']
            )

            # regra mínima (ajustável)
            if score_cliente > 80 and score_portal > 80:

                resultados.append({
                    'ID_ENVIADO': env_row['ID'],
                    'ID_INSERIDO': ins_row['ID'],
                    'CLIENTE_ENV': env_row['Cliente'],
                    'CLIENTE_INS': ins_row['Cliente'],
                    'PORTAL_ENV': env_row['Portal'],
                    'PORTAL_INS': ins_row['Portal'],
                    'SCORE_CLIENTE': score_cliente,
                    'SCORE_PORTAL': score_portal
                })

    df_resultado = pd.DataFrame(resultados)

    df_resultado.to_excel(arquivo_saida, index=False)

gerar_possiveis_matches('tecsystems_2025.xlsm', 'analise_teste_2025.xlsx')
print("Matches gerados com sucesso!")
