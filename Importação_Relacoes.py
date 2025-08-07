import pandas as pd
import pymysql
import traceback

# === Conexão com o banco ===
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
        traceback.print_exc()
        return None

# === Lê o Excel e trata os dados ===
df = pd.read_excel("Relações.xlsx")
df.columns = df.columns.str.strip()

# Substitui NaNs por None
df = df.where(pd.notnull(df), None)

# === Função de inserção ===
def inserir_banco(tabela, dados):
    try:
        conn = mysql_connection()
        if not conn:
            return False

        query = f"""
        INSERT INTO {tabela} 
        (ID, Participante, Avaliador1, Avaliador2, Tipo_Avaliacao)
        VALUES (%s, %s, %s, %s, %s)
        """

        with conn.cursor() as cursor:
            cursor.execute(query, dados)
            conn.commit()

        conn.close()
        return True
    except Exception as e:
        print(f'❌ Erro ao inserir dados no banco:')
        traceback.print_exc()
        return False

# === Loop de inserção ===
tabela_destino = 'QuestRH_Relacoes'

for index, row in df.iterrows():
    dados_linha = (
        row['ID'],
        row['Participante'],
        row['Avaliador1'],
        row['Avaliador2'],
        row['Tipo_Avaliacao']
    )

    print(f"🔄 Inserindo: {dados_linha}")
    sucesso = inserir_banco(tabela_destino, dados_linha)

    if not sucesso:
        print(f"❌ Erro ao inserir dados da pessoa: {row['Participante']}")
