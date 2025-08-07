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
df = pd.read_excel("Pessoas.xlsx")
df.columns = df.columns.str.strip()

# Renomeia coluna com hífen
df.rename(columns={'E-mail': 'Email'}, inplace=True)


# Data no formato mm/dd/yyyy
#df['Data_Nascimento'] = pd.to_datetime(df['Data_Nascimento'], errors='coerce')
#df['Data_Nascimento'] = df['Data_Nascimento'].dt.strftime('%Y/%m/%d')

# Substitui NaNs por None
df = df.where(pd.notnull(df), None)

# === Função de inserção ===
def Atualizar_Cadastro_Log(tabela, Login, Senha, Nome):
    try:
        conn = mysql_connection()
        if not conn:
            return False

        query = f"UPDATE {tabela} SET Login = '{Login}', Senha = '{Senha}' where Nome = '{Nome}'"
        #print(query)

        with conn.cursor() as cursor:
            cursor.execute(query)
            conn.commit()

        conn.close()
        return True
    except Exception as e:
        print(f'❌ Erro ao inserir dados no banco:')
        traceback.print_exc()
        return False


# === Função de inserção ===
def inserir_banco(tabela, dados):
    try:
        conn = mysql_connection()
        if not conn:
            return False

        query = f"""
        INSERT INTO {tabela} 
        (Nome, Email, C_Custo, Cargo, Local, CPF, Data_Nasc)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
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
tabela_destino = 'QuestRH_Pessoas'

for index, row in df.iterrows():

    '''
    dados_linha = (
        row['Nome'],
        row['Email'],
        row['C_Custo'],
        row['Cargo'],
        row['Local'],
        row['CPF'],
        row['Data_Nascimento']
    )
    '''

    
    print(f"🔄 Atulizando login de: {row['Nome']}")
    sucesso = Atualizar_Cadastro_Log(tabela_destino, row['Login'], row['Senha'], row['Nome'])

    if not sucesso:
        print(f"❌ Erro ao inserir dados da pessoa: {row['Nome']}")
