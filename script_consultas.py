import pandas as pd
import sqlite3
import openpyxl
import os


arquivo1 = "enviados_tratados.xlsx"
arquivo2 = "inseridos_tratados.xlsx"
saida = "comparacao_final.xlsx"


df1 = pd.read_excel(arquivo1)
df2 = pd.read_excel(arquivo2)


colunas_necessarias = ["Data", "Licitação", "Site", "Cliente", "UASG"]

for col in colunas_necessarias:
    if col not in df1.columns:
        raise Exception(f"Coluna faltando no df1: {col}")
    if col not in df2.columns:
        raise Exception(f"Coluna faltando no df2: {col}")


for df in [df1, df2]:
    df["Licitação"] = df["Licitação"].astype(str).str.strip()
    df["UASG"] = df["UASG"].astype(str).str.strip()
    df["Site"] = df["Site"].astype(str).str.strip()
    df["Cliente"] = df["Cliente"].astype(str).str.strip()

    # 🔥 normaliza data (CRÍTICO)
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.strftime("%Y-%m-%d")


df1 = df1.drop_duplicates(subset=colunas_necessarias)
df2 = df2.drop_duplicates(subset=colunas_necessarias)


conn = sqlite3.connect(":memory:")

df1.to_sql("t1", conn, index=False, if_exists="replace")
df2.to_sql("t2", conn, index=False, if_exists="replace")


query_debug = """
SELECT COUNT(*) as total_match
FROM t1
INNER JOIN t2
ON t1."Licitação" = t2."Licitação"
AND t1."UASG" = t2."UASG"
"""
print("🔍 Matches encontrados:", pd.read_sql(query_debug, conn))


query_classificacao = """
SELECT 
    t1."Licitação",
    t1."UASG",

    t1."Site" AS site_t1,
    t2."Site" AS site_t2,

    t1."Cliente" AS cliente_t1,
    t2."Cliente" AS cliente_t2,

    t1."Data" AS data_t1,
    t2."Data" AS data_t2,

    CASE
        WHEN t2."Licitação" IS NULL THEN 'SOMENTE_T1'

        WHEN t1."Licitação" = t2."Licitação"
         AND t1."UASG" = t2."UASG"
         AND t1."Site" = t2."Site"
         AND t1."Cliente" = t2."Cliente"
         AND t1."Data" = t2."Data"
        THEN 'IDENTICO'

        WHEN t1."Licitação" = t2."Licitação"
         AND t1."UASG" = t2."UASG"
         AND t1."Site" = t2."Site"
         AND t1."Cliente" = t2."Cliente"
         AND t1."Data" != t2."Data"
        THEN 'DATA_DIVERGENTE'

        WHEN t1."Licitação" = t2."Licitação"
         AND t1."UASG" = t2."UASG"
        THEN 'MATCH_FORTE'

        ELSE 'DIVERGENTE_TOTAL'
    END AS STATUS

FROM t1
LEFT JOIN t2
ON t1."Licitação" = t2."Licitação"
AND t1."UASG" = t2."UASG"
"""

df_classificacao = pd.read_sql(query_classificacao, conn)


query_parcial = """
SELECT 
    t1."Licitação",
    t1."UASG",

    t1."Site" AS site_t1,
    t2."Site" AS site_t2,

    t1."Cliente" AS cliente_t1,
    t2."Cliente" AS cliente_t2,

    t1."Data" AS data_t1,
    t2."Data" AS data_t2,

    'MATCH_PARCIAL' AS STATUS

FROM t1
INNER JOIN t2
ON (
    t1."Licitação" = t2."Licitação"
    OR t1."UASG" = t2."UASG"
)
WHERE NOT (
    t1."Licitação" = t2."Licitação"
    AND t1."UASG" = t2."UASG"
)
"""

df_parcial = pd.read_sql(query_parcial, conn)


df_conferencia = pd.concat([
    df_classificacao[
        df_classificacao["STATUS"].isin([
            "MATCH_FORTE",
            "DATA_DIVERGENTE"
        ])
    ],
    df_parcial
])

prioridade = {
    "IDENTICO": 1,
    "DATA_DIVERGENTE": 2,
    "MATCH_FORTE": 3,
    "MATCH_PARCIAL": 4,
    "DIVERGENTE_TOTAL": 5,
    "SOMENTE_T1": 6
}

df_classificacao["PRIORIDADE"] = df_classificacao["STATUS"].map(prioridade)
df_classificacao = df_classificacao.sort_values("PRIORIDADE")

df_melhor_match = df_classificacao.drop_duplicates(
    subset=["Licitação", "UASG"], keep="first"
)

query_t2 = """
SELECT t2.*
FROM t2
LEFT JOIN t1
ON t2."Licitação" = t1."Licitação"
AND t2."UASG" = t1."UASG"
WHERE t1."Licitação" IS NULL
"""

df_somente_t2 = pd.read_sql(query_t2, conn)


if os.path.exists(saida):
    try:
        os.remove(saida)
    except PermissionError:
        raise Exception("❌ Feche o arquivo Excel antes de rodar!")

with pd.ExcelWriter(saida, engine="openpyxl") as writer:
    df_melhor_match.to_excel(writer, sheet_name="RESUMO", index=False)
    df_conferencia.to_excel(writer, sheet_name="CONFERENCIA", index=False)
    df_somente_t2.to_excel(writer, sheet_name="SOMENTE_T2", index=False)

print("✅ Comparação inteligente finalizada!")