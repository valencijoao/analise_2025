import pandas as pd
import re
import openpyxl


def gerar_relatorio(arquivo_entrada, arquivo_saida):

    df_enviados = pd.read_excel(arquivo_entrada, sheet_name="Enviados")
    df_inseridos = pd.read_excel(arquivo_entrada, sheet_name="Inseridos")

    df_enviados.columns = df_enviados.columns.str.strip()
    df_inseridos.columns = df_inseridos.columns.str.strip()

    def limpar_id(valor):
        if pd.isna(valor):
            return ""
        
        valor = str(valor).strip()
        valor = re.sub(r'[^0-9a-zA-Z]', '', valor)
        return valor.upper()

    colunas_chave_enviados = ['Portal', 'Cliente', 'Data', 'ID']
    colunas_chave_inseridos = ['Data Licitação', 'Licitação','Site', 'Cliente']

    for col in colunas_chave_enviados:
        df_enviados[col + '_LIMPO'] = df_enviados[col].apply(limpar_id)

    for col in colunas_chave_inseridos:
        df_enviados[col + '_LIMPO'] = df_enviados[col].apply(limpar_id)


    df_inseridos_unico = df_inseridos[['ID_LIMPO']].drop_duplicates().copy()
    df_inseridos_unico['Participou'] = 'Sim'

    df_final = df_enviados.merge(
        df_inseridos_unico,
        on='ID_LIMPO',
        how='left'
    )

    df_final['Participou'] = df_final['Participou'].fillna('Não')

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="Analise Completa", index=False)



gerar_relatorio('tecsystems_2025.xlsm','analise_2026.xlsx' )
    
print("Arquivo gerado com sucesso!")