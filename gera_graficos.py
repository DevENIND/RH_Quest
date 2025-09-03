
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

import matplotlib.pyplot as plt
import pandas as pd
import io
import base64

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

        scrp_sql = f'Select Participante, ROUND(AVG(Resposta)) as Media from QuestRH_Respostas where {sql_condicao} group by Participante HAVING COUNT(DISTINCT id_rel) = 2'
        scrp_desempenho =  f'Select Participante, ROUND(AVG(Desempenho_tecnico)) as Media from QuestRH_Respostas where {sql_condicao} group by Participante HAVING COUNT(DISTINCT id_rel) = 2'

        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()

        cursor.execute(scrp_desempenho)
        resultados_desempenho = cursor.fetchall()
        conn.close()

        if len(resultados) == 0 or len(resultados_desempenho) == 0:
            return None, 'Não há dados para gerar o Nine-Box'

        # === Converter para DataFrame ===
        df_potencial = pd.DataFrame(resultados).rename(columns={"Media": "Potencial"})
        df_desempenho = pd.DataFrame(resultados_desempenho).rename(columns={"Media": "Performance"})

        # Juntar os dois dataframes
        df = pd.merge(df_potencial, df_desempenho, on="Participante", how="inner")

        # Map note to category
        def classificar(n):
            if n <= 2:
                return 'Baixo'
            elif n <= 4:
                return 'Médio'
            else:
                return 'Alto'
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
        def classificar(n):
            if n <= 2:
                return 'Baixo'
            elif n <= 4:
                return 'Médio'
            else:
                return 'Alto'

        df["Categoria_X"] = df["Potencial"].apply(classificar)
        df["Categoria_Y"] = df["Performance"].apply(classificar)

        # Percentual por quadrante
        n_total = len(df)
        quadros = df.groupby(["Categoria_X", "Categoria_Y"]).size().reset_index(name="Qtd")
        quadros["Perc"] = (quadros["Qtd"]/n_total*100).round(1)

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
               SELECT Pilar, ROUND(AVG(Media)) AS Media_Geral
                FROM (
                    SELECT Pilar, Participante, AVG(Resposta) AS Media
                    FROM QuestRH_Respostas
                    WHERE {sql_condicao}
                    GROUP BY Pilar, Participante
                    HAVING COUNT(DISTINCT id_rel) = 2
                ) AS Sub
                GROUP BY Pilar
                ORDER BY Pilar;
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
               SELECT Competencia, ROUND(AVG(Media)) AS Media_Geral
                FROM (
                    SELECT Competencia, Participante, AVG(Resposta) AS Media
                    FROM QuestRH_Respostas
                    WHERE {sql_condicao}
                    GROUP BY Competencia, Participante
                    HAVING COUNT(DISTINCT id_rel) = 2
                ) AS Sub
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
                    SELECT 
                        Competencia,
                        ROUND(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 2) AS Media_Auto,
                        ROUND(AVG(CASE WHEN id_rel IN (1,2) THEN Resposta END), 2) AS Media_Avaliadores
                    FROM QuestRH_Respostas
                    WHERE 
                        ({sql_condicao})
                        AND Participante IN (
                            SELECT Participante
                            FROM QuestRH_Respostas
                            WHERE id_rel IN (0,1,2)
                            GROUP BY Participante
                            HAVING COUNT(DISTINCT id_rel) = 3
                        )
                    GROUP BY Competencia
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
            desejado.append(float(res['Media_auto']))
            labels_apurado.append(str(res['Competencia']).strip())

        x = np.arange(len(labels_apurado))
        largura = 0.25

        # Altera o tamanho da figura, tamanho da colunas e texto
        fig, ax = plt.subplots(figsize=(15, 5))

        # Barras
        b1 = ax.bar(x - largura, desejado, largura, label="Auto", color="#0BE1FD")
        b2 = ax.bar(x, apurado, largura, label="Gestores", color="#0642C4")

        # Título
        ax.set_title("Comparativo - Auto Avaliação x AvaliaçãoGestores", fontsize=14, fontweight="bold")

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
    
def gera_bell_curve():
    try:
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute("""
            SELECT Participante, ROUND(AVG(Resposta),0) as Media
            FROM QuestRH_Respostas
            WHERE id_rel > 0
            GROUP BY Participante
            HAVING COUNT(DISTINCT id_rel) = 2
            ORDER BY Participante
        """)
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