
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import pymysql
import io
import base64
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.stats import norm


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


def gera_ninebox(pilar = '', competencia = '', participante = '', outras_condicoes = ''):
    try:
        sql_condicao =  "id_rel IN (1,2)"

        if pilar != '':
            sql_condicao += f" and Pilar = '{pilar}'"

        if competencia != '':
            sql_condicao += f" and Competencia = '{competencia}'"

        if participante != '':
            sql_condicao += f" and Participante = '{participante}'"

        if outras_condicoes != '' and outras_condicoes != None:
            sql_condicao += f" and {outras_condicoes}"

        scrp_sql = f"""
        Select Participante,
            CASE WHEN média_av1 = média_av2 THEN média_av1 ELSE CEIL( (média_av1 + média_av2) / 2) END AS Media_Avaliadores,
            CASE WHEN média_desemp_av1 = média_desemp_av2 THEN média_desemp_av1 ELSE CEIL( (média_desemp_av1 + média_desemp_av2) / 2) END AS Media_Aval_Desemp
        from 
        (SELECT Participante, id_rel, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
            round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1,
            round(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as média_desemp_av1,
            round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2,
            round(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as média_desemp_av2,
            round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_auto,
            round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_desemp_auto
        FROM QuestRH_Respostas WHERE {sql_condicao}
        group by Participante
        HAVING COUNT(DISTINCT id_rel) = 2
        ) as t
        """
        
        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()
 
        # === Converter para DataFrame ===
        df = pd.DataFrame(resultados)
       
        # Exemplo de dados (simulação)
        '''
        data = {
            "Participante": ["A", "B", "C", "D", "E", "F"],
            "Potencial": [2, 3, 4, 5, 5, 1],
            "Performance": [2, 4, 5, 3, 1, 5]
        }
        df = pd.DataFrame(data)
        '''
        # Função de classificação
        def classificar(row):
            x = row["Media_Avaliadores"]
            y = row["Media_Aval_Desemp"]
            participante = row["Participante"]

            if participante == 'Leandro Cabrini':
                print(x, y)

            if x in (1, 2) and y in (1, 2):
                return 'BAIXA PERFORMANCE'
            elif x in (1, 2) and y in (3, 4):
                return 'INCONSISTENTE'
            elif x in (1, 2) and y == 5:
                return 'ESPECIALISTA'
            elif x in (3, 4) and y in (1, 2):
                return 'DILEMA'
            elif x in (3, 4) and y in (3, 4):
                return 'COMPETENTE'
            elif x in (3, 4) and y == 5:
                return 'FORTE ENTREGA'
            elif x == 5 and y in (1, 2):
                return 'DESAFIO'
            elif x == 5 and y in (3, 4):
                return 'FORTE CULTURA'
            elif x == 5 and y == 5:
                return 'ALTO POTENCIAL'
            else:
                return 'N/A'

        df["Categoria"] = df.apply(classificar, axis=1)

        # Percentual por quadrante
        n_total = len(df)
        categorias = df.groupby("Categoria").size().reset_index(name="Qtd")
        categorias["Perc"] = (categorias["Qtd"] / n_total * 100).round(1)

        # Montar grid como imagem de Nine Box com base nas categorias
        # Mapeia posições dos quadrantes
        posicoes = {
            'BAIXA PERFORMANCE': ('Baixo', 'Baixo'),
            'INCONSISTENTE': ('Médio', 'Baixo'),
            'ESPECIALISTA': ('Alto', 'Baixo'),
            'DILEMA': ('Baixo', 'Médio'),
            'COMPETENTE': ('Médio', 'Médio'),
            'FORTE ENTREGA': ('Alto', 'Médio'),
            'DESAFIO': ('Baixo', 'Alto'),
            'FORTE CULTURA': ('Médio', 'Alto'),
            'ALTO POTENCIAL': ('Alto', 'Alto')
        }

        # Criar nova coluna com Categoria_X e Categoria_Y
        df['Categoria_X'] = df['Categoria'].map(lambda c: posicoes.get(c, ('', ''))[0])
        df['Categoria_Y'] = df['Categoria'].map(lambda c: posicoes.get(c, ('', ''))[1])

        # Títulos das células
        descricoes = {
            ('Alto','Baixo'): "DESAFIO",
            ('Alto','Médio'): "FORTE CULTURA",
            ('Alto','Alto'):  "ALTO POTENCIAL",
            ('Médio','Baixo'): "DILEMA",
            ('Médio','Médio'): "COMPETENTE",
            ('Médio','Alto'):  "FORTE ENTREGA",
            ('Baixo','Baixo'): "BAIXA PERFORMANCE",
            ('Baixo','Médio'): "INCONSISTENTE",
            ('Baixo','Alto'):  "ESPECIALISTA"
        }

        # Percentual por quadrante
        quadros = df.groupby(["Categoria_X", "Categoria_Y"]).size().reset_index(name="Qtd")
        quadros["Perc"] = (quadros["Qtd"] / n_total * 100).round(1)

        # Limites dos quadrantes
        boundaries = {
            "Baixo": (0.5, 2.5),
            "Médio": (2.5, 4.5),
            "Alto":  (4.5, 6.5)
        }

        
        # Cores das células
        cell_colors = {
            ('Baixo','Baixo'): (1.0, 0.3, 0.3, 0.6),
            ('Baixo','Médio'): (1.0, 0.7, 0.4, 0.6),
            ('Baixo','Alto'):  (1.0, 0.75, 0.0, 0.6),
            ('Médio','Baixo'): (1.0, 0.7, 0.4, 0.6),
            ('Médio','Médio'): (1.0, 0.75, 0.0, 0.6),
            ('Médio','Alto'):  (0.8, 1.0, 0.8, 0.6),
            ('Alto','Baixo'):  (1.0, 0.75, 0.0, 0.6),
            ('Alto','Médio'):  (0.8, 1.0, 0.8, 0.6),
            ('Alto','Alto'):   (0.7, 0.8, 0.9, 0.6)
        }


        # Criar figura
        fig, ax = plt.subplots(figsize=(8, 8))

        y_names = ["Alto","Médio","Baixo"]
        x_names = ["Baixo","Médio","Alto"]

        # Desenhar quadrantes e inserir textos
        for y_name in y_names:
            y0, y1 = boundaries[y_name]
            for x_name in x_names:
                x0, x1 = boundaries[x_name]

                # Preencher célula com cor
                ax.fill_between([x0,x1],[y0,y0],[y1,y1], color=cell_colors[(x_name,y_name)], edgecolor="k")

                # Percentual
                perc = quadros.loc[
                    (quadros["Categoria_X"]==x_name) & (quadros["Categoria_Y"]==y_name),
                    "Perc"
                ]
                perc_text = f"{perc.values[0]}%" if len(perc)>0 else "0%"

                # Título da célula
                titulo = descricoes.get((y_name, x_name), "")
                texto = f"{titulo}\n{perc_text}"

                ax.text((x0+x1)/2, (y0+y1)/2, texto,
                        ha="center", va="center", fontsize=10, weight="bold")

        # Ajuste dos eixos
        ax.set_xlim(0.5,6.5)
        ax.set_ylim(0.5,6.5)
        ax.set_xticks([1.5,3.5,5.5])
        ax.set_xticklabels(["Baixo (1-2)", "Médio (3-4)", "Alto (5)"])
        ax.set_yticks([1.5,3.5,5.5])
        ax.set_yticklabels(["Baixo (1-2)", "Médio (3-4)", "Alto (5)"])
        ax.set_xlabel("Potencial")
        ax.set_ylabel("Desempenho")
        ax.set_title("Nine Box", fontsize=16, weight="bold")
        ax.grid(False)

        # Remover margens e fundo branco
        fig.patch.set_alpha(0.0)
        ax.patch.set_alpha(0.0)

        # Salvar em memória como PNG transparente
        buf = io.BytesIO()
        plt.savefig(buf, format="png", transparent=True, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        img_bytes = buf.getvalue()

        # Converter para base64 (pronto para Flet)
        img64 = base64.b64encode(img_bytes).decode()
        #print(img64[:100], "...")  # mostra apenas início do base64

        # img64 agora pode ser usado diretamente em Flet
        return img64 , ''
    except Exception as e:
        print(f'❌ Erro ao gerar gráfico Ninebox: {e}')
        return None,e

def gera_gráfico_pilar(participante = '', pilar= '', competencia='', outras_condicoes = ''):
    try:
        # Dados
        categorias = ["E", "P", "C"]
        apurado = []

        sql_condicao =  "id_rel IN (1,2)"

        if pilar != '':
            sql_condicao += f" and Pilar = '{pilar}'"

        if competencia != '':
            sql_condicao += f" and Competencia = '{competencia}'"

        if participante != '':
            sql_condicao += f" and Participante = '{participante}'"

        if outras_condicoes != '' and outras_condicoes != None:
            sql_condicao += f" and {outras_condicoes}"


        scrp_sql = f"""
            Select Pilar, round(AVG(Media_Geral),0) as Media_Geral from(
            SELECT Pilar, Participante,
            CASE WHEN média_av1 = média_av2 THEN média_av1 ELSE CEIL( (média_av1 + média_av2) / 2) END AS Media_Geral 
            from(SELECT Pilar, Participante, 
                round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1, 
                round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2
                FROM QuestRH_Respostas
                WHERE {sql_condicao}
                GROUP BY Pilar, Participante
            HAVING COUNT(DISTINCT id_rel) = 2)as x 
            )as y group by Pilar
        """
        
        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()

        conn.close()

        if len(resultados) == 0:   
            return None, 'Nenhum dado encontrado.'
        
        if pilar == '':
            apurado.append(float(resultados[1]['Media_Geral'])) #E
            apurado.append(float(resultados[2]['Media_Geral'])) #P
            apurado.append(float(resultados[0]['Media_Geral'])) #C
            desejado = [4.0, 4.0, 5.0]

        elif pilar == 'Excelência com Cuidado ':
            apurado.append(float(resultados[0]['Media_Geral']))
            apurado.append(0.0)
            apurado.append(0.0)
            desejado = [4.0, 0.0, 0.0]

        elif pilar == 'Protagonismo Empreendedor ':
            apurado.append(0.0)
            apurado.append(float(resultados[0]['Media_Geral']))
            apurado.append(0.0)
            desejado = [0.0, 4.0, 0.0]


        elif pilar == 'Criação com Propósito ':
            apurado.append(0.0)
            apurado.append(0.0)
            apurado.append(float(resultados[0]['Media_Geral']))
            desejado = [0.0, 0.0, 5.0]


        gap = [a - b for a, b in zip(apurado, desejado)]

        x = np.arange(len(categorias))
        largura = 0.25

        fig, ax = plt.subplots(figsize=(8, 4))

        # Barras
        b1 = ax.bar(x - largura, desejado, largura, label="Desejado", color="#4472C4")
        b2 = ax.bar(x, apurado, largura, label="Apurado", color="#ED7D31")
        b3 = ax.bar(x + largura, gap, largura, label="GAP", color="gray")

        # Título
        ax.set_title("Comparativo Pilares", fontsize=14, fontweight="bold")

        # Eixo X
        ax.set_xticks(x)
        ax.set_xticklabels(categorias)

        # Eixo Y com margem para não cortar rótulos
        y_min = min(min(desejado), min(apurado), min(gap)) - 1.5
        y_max = max(max(desejado), max(apurado), max(gap)) + 1.5
        ax.set_ylim(y_min, y_max)

        # Linha zero
        ax.axhline(0, color="black", linewidth=0.8)

        # Legenda fora do gráfico (lado direito)
        ax.legend(loc="center left", bbox_to_anchor=(1, 0.5))

        # Função para rótulos
        def autolabel(barras):
            for barra in barras:
                altura = barra.get_height()
                ax.annotate(f'{altura:.0f}',
                            xy=(barra.get_x() + barra.get_width() / 2, altura),
                            xytext=(0, 5 if altura >= 0 else -10),
                            textcoords="offset points",
                            ha='center', va='bottom')

        autolabel(b1)
        autolabel(b2)
        autolabel(b3)

        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", transparent=True)
        plt.close(fig)
        buf.seek(0)

        # Converter para base64
        return base64.b64encode(buf.read()).decode("utf-8"), ''

    except Exception as e:
        return None, e

def gera_gráfico_Competencia(participante = '', pilar= '', competencia='',  outras_condicoes = ''):
    try:
        # Dados
 
        apurado = []
        desejado = []

        desejado_dict = {
            'Equilibrio emocional':  4.0,
            'Comprometimento com resultado':  4.0,
            'Resiliência':  4.0,
            'Faz como se fosse seu':  4.0,
            'Foco em soluções':  4.0,
            'Atualização e inovação constantes':  4.0,
            'Inspiração e mobilização de pessoas':  5.0,
            'Trabalho em equipe':  5.0,
            'Ensinar e compartilhar conhecimento':  5.0,
            'Resolve, não empurra':  4.0,
            'Entrega com precisão':  4.0,
            'Joga junto':  4.0,
            'Melhora sempre':  5.0,
        }

        sql_condicao =  "id_rel IN (1,2)"

        if pilar != '':
            sql_condicao += f" and Pilar = '{pilar}'"

        if competencia != '':
            sql_condicao += f" and Competencia = '{competencia}'"

        if participante != '':
            sql_condicao += f" and Participante = '{participante}'"
        
        if outras_condicoes != '' and outras_condicoes != None:
            sql_condicao += f" and {outras_condicoes}"

        scrp_sql = f"""
                SELECT Competencia, 
                ROUND(AVG(Media_geral),0) as Media_Geral from(
                SELECT Competencia, 
                CASE WHEN média_av1 = média_av2 THEN média_av1 ELSE CEIL( (média_av1 + média_av2) / 2) END AS Media_Geral
                    FROM (
                        SELECT Competencia, Participante, 
                        round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1, 
                        round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2
                        FROM QuestRH_Respostas
                        WHERE {sql_condicao}
                        GROUP BY Competencia, Participante
                        HAVING COUNT(DISTINCT id_rel) = 2
                    ) AS Sub) as x
                    GROUP BY Competencia
                    ORDER BY Competencia;
        """

        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()

        conn.close()

        if len(resultados) == 0:   
            return None, 'Nenhum dado encontrado.'
        
        labels_apurado = []
        for res in resultados:
            competencia = str(res['Competencia']).strip()
            apurado.append(float(res['Media_Geral']))
            desejado.append(float(desejado_dict[competencia]))
            labels_apurado.append(competencia)


        gap = np.array(apurado) - np.array(desejado)

        x = np.arange(len(desejado))
        largura = 0.25

        # Altera o tamanho da figura, tamanho da colunas e texto
        fig, ax = plt.subplots(figsize=(15, 5))

        # Barras
        b1 = ax.bar(x - largura, desejado, largura, label="Desejado", color="#6DDAE9")
        b2 = ax.bar(x, apurado, largura, label="Apurado", color="#0353FF")
        b3 = ax.bar(x + largura, gap, largura, label="GAP", color="gray")

        # Título
        ax.set_title("Comparativo Competências", fontsize=14, fontweight="bold")

        # Eixo X
        ax.set_xticks(x)
        ax.set_xticklabels(labels_apurado, rotation=45, ha="right")

        # Eixo Y com margem para não cortar rótulos
        y_min = min(min(desejado), min(apurado), min(gap)) - 1.5
        y_max = max(max(desejado), max(apurado), max(gap)) + 1.5
        ax.set_ylim(y_min, y_max)

        # Linha zero
        ax.axhline(0, color="black", linewidth=0.8)

        # Legenda fora do gráfico (lado direito)
        ax.legend(loc="center left", bbox_to_anchor=(1, 0.5))

        # Função para rótulos
        def autolabel(barras):
            for barra in barras:
                altura = barra.get_height()
                ax.annotate(f'{altura:.0f}',
                            xy=(barra.get_x() + barra.get_width() / 2, altura),
                            xytext=(0, 5 if altura >= 0 else -10),
                            textcoords="offset points",
                            ha='right')

        
        autolabel(b1)
        autolabel(b2)
        autolabel(b3)

        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", transparent=True)
        plt.close(fig)
        buf.seek(0)

        # Converter para base64
        return base64.b64encode(buf.read()).decode("utf-8"), ''

    except Exception as e:
        print(f'❌ Erro ao gerar gráfico de competências: {e}')
        return None, e

def gera_gráfico_Comparativo(participante = '', pilar= '', competencia='', outras_condicoes = ''):
    try:
        # Dados
        categorias = ['Equilibrio emocional',
            'Comprometimento com resultado',
            'Resiliência',
            'Faz como se fosse seu',
            'Foco em soluções',
            'Atualização e inovação constantes',
            'Inspiração e mobilização de pessoas',
            'Trabalho em equipe',
            'Ensinar e compartilhar conhecimento',
            'Resolve, não empurra',
            'Entrega com precisão',
            'Joga junto',
            'Melhora sempre',
            ]
        apurado = []
        desejado = []

        sql_condicao =  "id_rel IN (0,1,2)"

        if pilar != '':
            sql_condicao += f" and Pilar = '{pilar}'"

        if competencia != '':
            sql_condicao += f" and Competencia = '{competencia}'"

        if participante != '':
            sql_condicao += f" and Participante = '{participante}'"

        if outras_condicoes != '' and outras_condicoes != None:
            sql_condicao += f" and {outras_condicoes}"

        scrp_sql = f"""
                    Select Competencia,
                    ROUND(AVG(Media_Auto),0) as Media_Auto,
                    ROUND(AVG(Media_Avaliadores),0) as Media_Avaliadores from (
                    Select Competencia,
                        Participante,
                        Media_Auto,
                        CASE WHEN média_av1 = média_av2 THEN média_av1 ELSE CEIL( (média_av1 + média_av2) / 2) END AS Media_Avaliadores
                    from (SELECT Competencia, Participante,
                        ROUND(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) AS Media_Auto,
                        round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1,
                        round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2
                    FROM QuestRH_Respostas WHERE {sql_condicao} AND Participante IN (
                            SELECT Participante
                            FROM QuestRH_Respostas
                            WHERE id_rel IN (0,1,2)
                            GROUP BY  Participante
                            HAVING COUNT(DISTINCT id_rel) = 3
                        )group by  Competencia, Participante) as x
                        ) as y group by  Competencia
                    """


        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()
        conn.close()

        if len(resultados) == 0:   
            return None, 'Nenhum dado encontrado.'
        
        labels_apurado = []
        for res in resultados:
            apurado.append(float(res['Media_Avaliadores']))
            desejado.append(float(res['Media_Auto']))
            labels_apurado.append(str(res['Competencia']).strip())

        x = np.arange(len(labels_apurado))
        largura = 0.25

        # Altera o tamanho da figura, tamanho da colunas e texto
        fig, ax = plt.subplots(figsize=(15, 5))

        # Barras
        b1 = ax.bar(x - largura, desejado, largura, label="Auto", color="#0BE1FD")
        b2 = ax.bar(x, apurado, largura, label="Gestores", color="#0642C4")

        # Título
        ax.set_title("Comparativo - Auto Avaliação x Avaliação Gestores", fontsize=14, fontweight="bold")

        # Eixo X
        ax.set_xticks(x)
        ax.set_xticklabels(labels_apurado, rotation=45, ha="right")

        # Eixo Y com margem para não cortar rótulos
        y_min = min(min(desejado), min(apurado)) - 1.5
        y_max = max(max(desejado), max(apurado)) + 1.5
        ax.set_ylim(y_min, y_max)

        # Linha zero
        ax.axhline(0, color="black", linewidth=0.8)

        # Legenda fora do gráfico (lado direito)
        ax.legend(loc="center left", bbox_to_anchor=(1, 0.5))

        # Função para rótulos
        def autolabel(barras):
            for barra in barras:
                altura = barra.get_height()
                ax.annotate(f'{altura:.0f}',
                            xy=(barra.get_x() + barra.get_width() / 2, altura),
                            xytext=(0, 5 if altura >= 0 else -10),
                            textcoords="offset points",
                            ha='right')


        autolabel(b1)
        autolabel(b2)

        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", transparent=True)
        plt.close(fig)
        buf.seek(0)

        # Converter para base64
        return base64.b64encode(buf.read()).decode("utf-8"), ''

    except Exception as e:
        print(f'❌ Erro ao gerar gráfico de comparação: {e}')
        return None, e
    
def gera_bell_curve(estrategico = True, nao_estrategico = True):
    try:
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)

        sql_condicao =  'id_rel in (1,2)'

        if estrategico == False and nao_estrategico == True:
            sql_condicao += f" and Grupo_estrategico = 'Não'"
        elif estrategico == True and nao_estrategico == False:
            sql_condicao += f" and Grupo_estrategico = 'Sim'"


        query_sql = f"""
            Select Participante,
                    CASE WHEN média_av1 = média_av2 THEN média_av1 ELSE CEIL( (média_av1 + média_av2) / 2) END AS Media
            from 
            (SELECT Participante, id_rel, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
                round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1,
                round(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as média_desemp_av1,
                round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2,
                round(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as média_desemp_av2,
                round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_auto,
                round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_desemp_auto
            FROM QuestRH_Respostas WHERE {sql_condicao}
            group by Participante
            HAVING COUNT(DISTINCT id_rel) = 2
            ) as t ORDER BY Participante
        """


        cursor.execute(query_sql)
        pessoas = cursor.fetchall()
        conn.close()

        dados = [float(p['Media']) for p in pessoas if p['Media'] is not None]

        # Filtrar apenas valores numéricos válidos
        dados = [x for x in dados if isinstance(x, (int, float)) and not np.isnan(x)]

        if len(dados) < 2:
            # Não tem dados suficientes
            fig, ax = plt.subplots(figsize=(6,4))
            ax.text(0.5, 0.5, "Sem dados suficientes", ha="center", va="center")
            ax.axis("off")
            plt.tight_layout()
        else:
            mu, sigma = np.mean(dados), np.std(dados)

            if sigma == 0:
                # Todos os valores iguais → histograma simples
                fig, ax = plt.subplots(figsize=(6,4))
                ax.hist(dados, bins=5, color="lightblue", edgecolor="darkblue")
                ax.set_title("Todos os valores iguais", fontsize=14, weight="bold")
                ax.set_xlabel("Notas")
                ax.set_ylabel("Frequência")
                plt.tight_layout()
            else:
                # Criar curva normal
                x = np.linspace(min(dados)-1, max(dados)+1, 200)
                y = norm.pdf(x, mu, sigma) * len(dados)

                fig, ax = plt.subplots(figsize=(8,5))
                ax.plot(x, y, color="darkblue", linewidth=3)
                ax.fill_between(x, y, color="lightblue", alpha=0.5)

                # Percentual de cada ponto
                total = len(dados)
                for nota in dados:
                    y_val = norm.pdf(nota, mu, sigma) * total
                    qtd = dados.count(nota)
                    perc = (dados.count(nota) / total) * 100
                    ax.scatter(nota, y_val, color="darkblue", s=60, zorder=5)
                    ax.annotate(
                        f"nota: {nota}\npercentual:{perc:.1f}%\nQtd pessoas: {qtd}",              # texto do rótulo
                        xy=(nota, y_val),            # ponto a ser ligado
                        xytext=(nota, -0.3),         # posição do texto (abaixo do gráfico)
                        ha="center", va="top",
                        fontsize=9, color="black",
                        arrowprops=dict(
                                arrowstyle="-", 
                                color="gray", 
                                alpha=0.4   # 🔹 transparência da linha
                        ),
                        bbox=dict(
                            facecolor="white", 
                            alpha=0.5, # 🔹 transparência do fundo do texto
                            boxstyle="round,pad=0.2"
                        )
                    )

                    #ax.text(nota, y_val - 0.2, f"nota: {nota}\npercentual:{perc:.1f}%\nQtd pessoas: {qtd}", 
                            #ha="center", va="bottom", fontsize=9, color="black")

                ax.set_title("Distribuição Bell Curve", fontsize=14, weight="bold")
                ax.set_xlabel("Notas")
                ax.set_ylabel("Quantidade de Pessoas")
                ax.grid(False)
                plt.tight_layout()

        # Converter para base64
        buf = io.BytesIO()
        plt.savefig(buf, format="png", transparent=True)
        plt.close(fig)
        buf.seek(0)
        return base64.b64encode(buf.read()).decode("utf-8"), ''
    
    except Exception as e:
        print(f'❌ Erro ao gerar gráfico Bell Curve: {e}')
        return None, e


def gera_grafico_conclusao_feedback(anterior = False):
    try:

        sql_condicao =''

        if anterior:
            tabela = 'QuestRH_Feedbacks'
            texto_grafico = 'Feedabacks Ciclo Anterior'
        else:
            tabela =  'QuestRH_Relacoes'
            texto_grafico = 'Feedabacks Ciclo Atual'
            
        connection = mysql_connection()
        cursor = connection.cursor()
        
        sql_condicao= f"Select count(Participante) as Total from {tabela}"
        cursor.execute(sql_condicao)
        qtd_participantes = cursor.fetchone()[0]
        if anterior:
            sql_condicao= f"""Select count(Participante) as Total from QuestRH_Feedbacks where 
            data_feedback is not null and id_ciclo = (SELECT MAX(id_ciclo) FROM QuestRH_Feedbacks);"""
        else: 
            sql_condicao= f"Select count(Participante) as Total from QuestRH_Relacoes where data_feedback is not null"
        cursor.execute(sql_condicao)
        realizado= cursor.fetchone()[0]
        
        connection.close()

        if realizado is None:
            realizado = 0

        # Dados
        nao_realizados = qtd_participantes - realizado
        perc_realizado = round((realizado / qtd_participantes) * 100,2)
        perc_nao_realizados = round((nao_realizados / qtd_participantes) * 100,2)

        sizes = [perc_realizado, perc_nao_realizados]  # Percentuais ou valores absolutos
        colors = ["#0076c5b2","#E4E4E49E"]  # Cores personalizadas

        # Criando o gráfico de rosca
        fig, ax = plt.subplots(figsize=(4, 4))
        ax.pie(
            sizes, 
            #=labels,
            colors=colors, 
            startangle=90,  # Rotaciona o gráfico para começar de cima
            wedgeprops={'width':0.4},  # largura da "roda", define o buraco do centro
            #autopct='%1.1f%%'  # Mostra os percentuais
        )
        # Mantém o gráfico como círculo
        ax.axis('equal')

        ax.text(
            0, 0, 
            f'{perc_realizado}%',  # texto a ser exibido
            horizontalalignment='center', 
            verticalalignment='center',
            fontsize=30,  # ajuste o tamanho conforme necessário
            fontfamily='serif',
            color="#0076c5ff"
        )

    

        plt.text(
            0, -1.2,  # coordenadas x, y (y negativo abaixo do centro)
            texto_grafico, 
            horizontalalignment='center',
            fontsize=18,
            color="black",
            fontfamily='serif'
        )

        #if finalizados == False:
        #    plt.savefig("gauge.png", bbox_inches="tight")
        #else:
        #    plt.savefig("gauge_finalizado.png", bbox_inches="tight")
        # Exibe o gráfico
        #plt.show()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", transparent=True)
        plt.close(fig)
        buf.seek(0)

        # Converter para base64
        return base64.b64encode(buf.read()).decode("utf-8"), ''

    except Exception as e:
        print(f'❌ Erro ao gerar gráfico de performances:{e}')
        return None, e

def gera_grafico_conclusao(finalizados = False):
    try:

        sql_condicao =''

        if finalizados == False:
            agrupar = 'x.Participante, x.Nome_Avaliador'
            texto_grafico = 'Avaliações Finalizadas'
            qtd_participantes = 386 #Quantidade constante na tabela Lista_Emails
        else:
            sql_condicao = "Where Status_Av = 'Finalizado'"
            agrupar = 'x.Participante'
            texto_grafico = 'Participantes Finalizados'
            scrp_participantes = f"""Select count(Participante) as Total from QuestRH_Relacoes"""
            connection = mysql_connection()
            cursor = connection.cursor()
            cursor.execute(scrp_participantes)
            qtd_participantes = cursor.fetchone()[0]
            connection.close()


        if finalizados == True:
            scrp_sql = f"""
                    Select Sum(Grupo) as Realizado from
                    (SELECT Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo, count(concat(Participante, ' - ', Nome_Avaliador)) as Grupo, 
                            Case WHEN  isnull(Media_Auto) and Avaliacao = 'A3' Then 'N/A' Else Media_Auto End as Media_Auto,
                            Media_Avaliadores, 
                            Case When isnull(Media_Auto_Desemp) and Avaliacao = 'A3' Then 'N/A' Else Media_Auto_Desemp End as Media_Auto_Desemp, 
                            Media_Aval_Desemp,
                            CASE 
                                WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (1, 2) THEN 'Baixa Performance'
                                WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (3, 4) THEN 'Inconsistente'
                                WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp = 5 THEN 'Especialista'

                                WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (1, 2) THEN 'Dilema'
                                WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (3, 4) THEN 'Competente'
                                WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp = 5 THEN 'Forte Entrega'

                                WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (1, 2) THEN 'Desafio'
                                WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (3, 4) THEN 'Forte Cultura'
                                WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp = 5 THEN 'Alto Potencial'

                                ELSE ''
                            END as Performance,
                            Status_Av from (
                        Select Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo, Nome_Avaliador,
                            média_auto as Media_Auto,
                            CASE WHEN média_av1 = média_av2 THEN média_av1 ELSE CEIL( (média_av1 + média_av2) / 2) END AS Media_Avaliadores,
                            média_desemp_auto as Media_Auto_Desemp,
                            CASE WHEN média_desemp_av1 = média_desemp_av2 THEN média_desemp_av1 ELSE CEIL( (média_desemp_av1 + média_desemp_av2) / 2) END AS Media_Aval_Desemp,
                            Status_Av
                        from 
                        (SELECT Participante, id_rel, Avaliacao, Sigla_Emp, C_Custo, Cargo, Nome_Avaliador, 
                            round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1,
                            round(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as média_desemp_av1,
                            round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2,
                            round(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as média_desemp_av2,
                            round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_auto,
                            round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_desemp_auto,
                            -- Status
                        CASE 
                            WHEN (COUNT(DISTINCT CASE When id_rel in (1,2) then id_rel end) = 2)
                            THEN 'Finalizado' 
                            ELSE 'Pendente' 
                        END as Status_Av
                        FROM QuestRH_Respostas
                        group by Participante) as t
                    ) as x {sql_condicao} group by {agrupar}
                    )y;
            """
        else:
            scrp_sql = f"""
                Select Sum(Grupo) as Realizado from
                    (SELECT count(concat(Participante, ' - ', Nome_Avaliador)) as Grupo 
                        From (SELECT Distinct Participante, Nome_Avaliador FROM QuestRH_Respostas) t
                )y;
            """

        connection = mysql_connection()
        cursor = connection.cursor()
        cursor.execute(scrp_sql)
        realizado = cursor.fetchone()[0]
        connection.close()

        if realizado is None:
            realizado = 0

        # Dados
        nao_realizados = qtd_participantes - realizado
        perc_realizado = round((realizado / qtd_participantes) * 100,2)
        perc_nao_realizados = round((nao_realizados / qtd_participantes) * 100,2)

        sizes = [perc_realizado, perc_nao_realizados]  # Percentuais ou valores absolutos
        colors = ["#0076c5b2","#E4E4E49E"]  # Cores personalizadas

        # Criando o gráfico de rosca
        fig, ax = plt.subplots(figsize=(4, 4))
        ax.pie(
            sizes, 
            #=labels,
            colors=colors, 
            startangle=90,  # Rotaciona o gráfico para começar de cima
            wedgeprops={'width':0.4},  # largura da "roda", define o buraco do centro
            #autopct='%1.1f%%'  # Mostra os percentuais
        )
        # Mantém o gráfico como círculo
        ax.axis('equal')

        ax.text(
            0, 0, 
            f'{perc_realizado}%',  # texto a ser exibido
            horizontalalignment='center', 
            verticalalignment='center',
            fontsize=30,  # ajuste o tamanho conforme necessário
            fontfamily='serif',
            color="#0076c5ff"
        )

    

        plt.text(
            0, -1.2,  # coordenadas x, y (y negativo abaixo do centro)
            texto_grafico, 
            horizontalalignment='center',
            fontsize=18,
            color="black",
            fontfamily='serif'
        )

        #if finalizados == False:
        #    plt.savefig("gauge.png", bbox_inches="tight")
        #else:
        #    plt.savefig("gauge_finalizado.png", bbox_inches="tight")
        # Exibe o gráfico
        #plt.show()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", transparent=True)
        plt.close(fig)
        buf.seek(0)

        # Converter para base64
        return base64.b64encode(buf.read()).decode("utf-8"), '', realizado, nao_realizados

    except Exception as e:
        print(f'❌ Erro ao gerar gráfico de performances:{e}')
        return None, e, 0.00
    
def gera_gráfico_potencial(finalizados = False):
    try:
        # Dados
        categorias = []
        apurado = []

        sql_condicao =  "Where id_rel IN (1,2)"

        if finalizados == True:
            titulo = "Distribuição de performances finalizadas"
            sql_condicao_final = " and Status_Av = 'Finalizado'"
        else:
            titulo = "Distribuição de performances não finalizadas"
            sql_condicao_final = ' and Status_Av <> "Finalizado"'


        scrp_sql =f"""
        SELECT Performance, COUNT(Participante) as Contagem From (
            SELECT Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
                    Case WHEN  isnull(Media_Auto) and Avaliacao = 'A3' Then 'N/A' Else Media_Auto End as Media_Auto,
                    Media_Avaliadores, 
                    Case When isnull(Media_Auto_Desemp) and Avaliacao = 'A3' Then 'N/A' Else Media_Auto_Desemp End as Media_Auto_Desemp, 
                    Media_Aval_Desemp,
                    CASE 
                        WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (1, 2) THEN 'Baixa Performance'
                        WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (3, 4) THEN 'Inconsistente'
                        WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp = 5 THEN 'Especialista'

                        WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (1, 2) THEN 'Dilema'
                        WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (3, 4) THEN 'Competente'
                        WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp = 5 THEN 'Forte Entrega'

                        WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (1, 2) THEN 'Desafio'
                        WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (3, 4) THEN 'Forte Cultura'
                        WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp = 5 THEN 'Alto Potencial'

                        ELSE ''
                    END as Performance,
                    Status_Av from (
                Select Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo,
                    média_auto as Media_Auto,
                    CASE WHEN média_av1 = média_av2 THEN média_av1 
                    ELSE CASE WHEN isnull(média_av1) Then média_av2
                    ELSE CASE When isnull(média_av2) then média_av1 
                    ELSE CEIL( (média_av1 + média_av2) / 2) END END END
                    AS Media_Avaliadores,
                    média_desemp_auto as Media_Auto_Desemp,
                    CASE WHEN média_desemp_av1 = média_desemp_av2 THEN média_desemp_av1 
                     ELSE CASE WHEN isnull(média_desemp_av1) Then média_desemp_av2
                    ELSE CASE When isnull(média_desemp_av2) then média_desemp_av1 
                    ELSE CEIL( (média_desemp_av1 + média_desemp_av2) / 2) END END END AS Media_Aval_Desemp,
                    Status_Av
                from 
                (SELECT Participante, id_rel, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
                    round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1,
                    round(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as média_desemp_av1,
                    round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2,
                    round(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as média_desemp_av2,
                    round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_auto,
                    round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_desemp_auto,
                    -- Status
                CASE 
                    WHEN (COUNT(DISTINCT CASE When id_rel in (1,2) then id_rel end) = 2)
                    THEN 'Finalizado' 
                    ELSE 'Pendente' 
                END as Status_AV
                FROM QuestRH_Respostas {sql_condicao}
                group by Participante) as t
            ) as x        
        ) as y where Performance <> '' {sql_condicao_final}
                group by Performance 
                order by Contagem;
        """
        
        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()

        conn.close()

        if len(resultados) == 0:   
            return None, 'Nenhum dado encontrado.'
        else:
            for res in resultados:
                categorias.append(res['Performance'])
                apurado.append(res['Contagem'])
        

        x = np.arange(len(categorias))
        largura = 0.25

        fig, ax = plt.subplots(figsize=(8, 4))

        # Barras)
        b1 = ax.bar(x, apurado, largura, label="Apurado", color="#0076c5b2")

        # Título
        ax.set_title(titulo, fontsize=14, fontfamily='serif')

        # Eixo X
        ax.set_xticks(x)
        ax.set_xticklabels(categorias)

        # Eixo Y com margem para não cortar rótulos
        y_min = min(apurado) - 5
        y_max = max(apurado) + 5
        ax.set_ylim(y_min, y_max)

        # Linha zero
        ax.axhline(0, color="black", linewidth=0.8)

        # Legenda fora do gráfico (lado direito)
        ax.legend(loc="center left", bbox_to_anchor=(1, 0.5))

        # Função para rótulos
        def autolabel(barras):
            for barra in barras:
                altura = barra.get_height()
                ax.annotate(f'{altura:.0f}',
                            xy=(barra.get_x() + barra.get_width() / 2, altura),
                            xytext=(0, 5 if altura >= 0 else -10),
                            textcoords="offset points",
                            ha='center', va='bottom',fontfamily='serif')

        autolabel(b1)
        plt.tight_layout()
        #if finalizados == True:
        #    plt.savefig("potencial_finalizados.png", bbox_inches="tight")
        #else:
        #    plt.savefig("potencial.png", bbox_inches="tight")

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", transparent=True)
        plt.close(fig)
        buf.seek(0)

        # Converter para base64
        return base64.b64encode(buf.read()).decode("utf-8"), ''

    except Exception as e:
        print(f'❌ Erro ao gerar gráfico de potencial: {e}')
        return None, e
   

def gera_grafico_empresas():
    try:
        scrp_sql = f"""
                SELECT Sigla_Emp, COUNT(Participante) as Contagem From (SELECT Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
                Case WHEN  isnull(Media_Auto) and Avaliacao = 'A3' Then 'N/A' Else Media_Auto End as Media_Auto,
                Media_Avaliadores, 
                Case When isnull(Media_Auto_Desemp) and Avaliacao = 'A3' Then 'N/A' Else Media_Auto_Desemp End as Media_Auto_Desemp, 
                Media_Aval_Desemp,
                CASE 
                    WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (1, 2) THEN 'Baixa Performance'
                    WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (3, 4) THEN 'Inconsistente'
                    WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp = 5 THEN 'Especialista'

                    WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (1, 2) THEN 'Dilema'
                    WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (3, 4) THEN 'Competente'
                    WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp = 5 THEN 'Forte Entrega'

                    WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (1, 2) THEN 'Desafio'
                    WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (3, 4) THEN 'Forte Cultura'
                    WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp = 5 THEN 'Alto Potencial'

                    ELSE ''
                END as Performance,
                Status_Av from (
                Select Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo,
                    média_auto as Media_Auto,
                    CASE WHEN média_av1 = média_av2 THEN média_av1 ELSE CEIL( (média_av1 + média_av2) / 2) END AS Media_Avaliadores,
                    média_desemp_auto as Media_Auto_Desemp,
                    CASE WHEN média_desemp_av1 = média_desemp_av2 THEN média_desemp_av1 ELSE CEIL( (média_desemp_av1 + média_desemp_av2) / 2) END AS Media_Aval_Desemp,
                    Status_Av
                from 
                (SELECT Participante, id_rel, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
                    round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1,
                    round(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as média_desemp_av1,
                    round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2,
                    round(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as média_desemp_av2,
                    round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_auto,
                    round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_desemp_auto,
                    -- Status
                CASE 
                    WHEN (COUNT(DISTINCT CASE When id_rel in (1,2) then id_rel end) = 2)
                    THEN 'Finalizado' 
                    ELSE 'Pendente' 
                END as Status_Av
                FROM QuestRH_Respostas
                group by Participante) as t
            ) as x 
        ) as y where Status_Av = 'Finalizado' group by Sigla_Emp order by Sigla_Emp
        """
        Valores = [0,0,0,0]
        Textos=['0 de 96 (0%)','0 de 2 (0%)','0 de 15 (0%)','0 de 30 (0%)']
        Percentual = []

        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()

        conn.close()
        

        Listagem_Avaliacoes ={
            'Construção': 96,
            'Locações': 2,
            'Montagens Industriais': 15,
            'Serviços': 30,
        }

        # Dados
        labels = [
            "Construção",
            "Locações",
            "Montagens Industriais",
            "Serviços",
        ]

        if len(resultados) == 0:   
            return None, 'Nenhum dado encontrado.'
        else:
            for res in resultados:
                sigla_emp = res['Sigla_Emp']
                Percentual =  (float(res['Contagem'])/float(Listagem_Avaliacoes[sigla_emp]))*100
                for i, x in enumerate(labels):
                    if sigla_emp == x:
                        Valores[i] = Percentual
                        Textos[i] = f"{res['Contagem']} de {Listagem_Avaliacoes[sigla_emp]} ({Percentual:.2f}%)"


        #values = [2, 40, 40, 1, 0, None, 1964, 395]  # None usado porque "Normal" não é numérico
        #texts = ["2", "40%", "40%", "1", "0", "Normal", "1964", "395"]

        colors = [
            "blue",         # Construção
            "orange",       # Locações
            "green",        # Montagens Industriais
            "red",          # Serviços
        ]

        # Plot
        fig, ax = plt.subplots(figsize=(8, 5))

        y_pos = range(len(labels))

        # Paleta mais suave (pastéis)
        colors = [
            "#6baed6",  # azul claro
            "#fdae6b",  # laranja suave
            "#74c476",  # verde suave
            "#fc9272",  # vermelho suave
        ]

        for i, (label, val, text, color) in enumerate(zip(labels, Valores, Textos, colors)):
            if val is not None:
                ax.barh(
                    i,
                    val,
                    color=color,
                    height=0.45,
                    edgecolor="black",   # borda preta fina
                    linewidth=0.8
                )
                ax.text(
                    val + max(Valores) * 0.01,
                    i,
                    text,
                    va="center",
                    fontsize=9,
                    fontweight="bold",
                    color="black"
                )
            else:
                ax.barh(
                    i,
                    0.1,
                    color=color,
                    height=0.45,
                    edgecolor="black",
                    linewidth=0.8
                )
                ax.text(0.5, i, text, va="center", ha="center", fontsize=9, fontweight="bold")

        # Estilo do eixo
        ax.set_yticks(y_pos)
        ax.set_yticklabels(labels, fontsize=10, fontweight="bold")
        ax.invert_yaxis()
        ax.set_xlim(0, max([v for v in Valores if v is not None]) * 1.25)

        # Título mais destacado
        ax.set_title("Empresas Concluídas", fontsize=13, fontweight="bold", pad=15)

        # Remover eixos desnecessários
        ax.xaxis.set_visible(False)
        for spine in ["right", "top", "bottom"]:
            ax.spines[spine].set_visible(False)

        plt.tight_layout()

        # Salvar
        #plt.savefig("barra_empresas.png", bbox_inches="tight", dpi=150)

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", transparent=True)
        plt.close(fig)
        buf.seek(0)


        return base64.b64encode(buf.getvalue()).decode(), None
    except Exception as e:
        print(f'❌ Erro ao gerar gráfico de barras de empresas: {e}')
        return None, e
