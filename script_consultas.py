import pandas as pd
import sqlite3

df_enviados = pd.read_excel('enviados_tratados.xlsx')
df_inseridos = pd.read_excel('inseridos_tratados.xlsx')

conn = sqlite3.connect(':memory:')

df_enviados.to_sql('enviados_tratados', conn, index=False, if_exists='replace')

df_inseridos.to_sql('inseridos_tratados', conn, index=False, if_exists='replace')

clientes_unicos1 = "SELECT DISTINCT Cliente FROM enviados_tratados"
clientes_unicos2 = "SELECT DISTINCT Cliente FROM inseridos_tratados"

df_clientes_unicos1 = pd.read_sql_query(clientes_unicos1, conn)
df_clientes_unicos2 = pd.read_sql_query(clientes_unicos2, conn)

df_consulta_clientes = pd.concat([df_clientes_unicos1.reset_index(drop=True), 
                            df_clientes_unicos2.reset_index(drop=True)], 
                           axis=1)
 
df_consulta_clientes.to_csv('lista_clientes_unicos.csv', index=False, sep=';', encoding='latin1')