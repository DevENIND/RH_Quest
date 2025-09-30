import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import plotly.graph_objects as go
import pymysql
import io
import base64
import flet as ft

def mysql_connection():
    host = 'bdnuvemwa.mysql.dbaas.com.br'
    database = 'bdnuvemwa'
    user = 'bdnuvemwa'
    passwd = 'W102030b!@'

    try:
        connection = pymysql.connect(
            host=host,
            user=user,
            password=passwd,
            database=database,
            charset="utf8mb4"
        )
        return connection
    except Exception:
        print('❌ Erro ao conectar ao banco de dados:')
        return None


import pymysql
import math
import pandas as pd

def calcular_media(notas):
    if not notas:
        return None
    if min(notas) == max(notas):
        return notas[0]
    return math.ceil(sum(notas) / len(notas))

# Conexão
conn =  mysql_connection()

# Consulta básica
df = pd.read_sql("""
    SELECT Participante, id_rel, Resposta
    FROM QuestRH_Respostas
""", conn)

# Agrupamento e aplicação da função customizada
resultado = df.groupby("Participante")["Resposta"].apply(list).apply(calcular_media)
print(resultado)
