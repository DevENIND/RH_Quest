
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import pymysqlgi
import io
import base64
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt



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



def gera_ninebox(pilar = '', competencia = '', participante = ''):
    try:
        sql_condicao =  "id_rel IN (1,2)"

        if pilar != '':
            sql_condicao += f" and Pilar = '{pilar}'"

        if competencia != '':
            sql_condicao += f" and Competencia = '{competencia}'"

        if participante != '':
            sql_condicao += f" and Participante = '{participante}'"


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

        # Performance (X) e Potencial (Y) — por enquanto usamos a mesma nota para ambas
        df['Categoria_X'] = df["Potencial"].apply(classificar) # Performance
        df['Categoria_Y'] = df["Performance"].apply(classificar)  # Potencial

        # Total de pessoas
        n_total = len(df)

        # Percentual por quadrante — agrupamos por Performance (X) e Potencial (Y)
        quadros = df.groupby(['Categoria_X', 'Categoria_Y']).size().reset_index(name='Qtd')
        quadros['Perc'] = (quadros['Qtd'] / n_total * 100).round(1)

        # Grid boundaries
        boundaries = {
            'Baixo': (0.5, 2.5),
            'Médio': (2.5, 4.5),
            'Alto':  (4.5, 6.5)
        }

        # Colors for the nine cells
        cell_colors = {
            ('Baixo','Baixo'): 'rgba(255,  77,  77,0.6)',
            ('Baixo','Médio'): 'rgba(255, 178, 102,0.6)',
            ('Baixo','Alto'):  'rgba(255, 192, 0,0.6)',
            ('Médio','Baixo'): 'rgba(255, 178, 102,0.6)',
            ('Médio','Médio'): 'rgba(255, 192, 0,0.6)',
            ('Médio','Alto'):  'rgba(204, 255, 204,0.6)',
            ('Alto','Baixo'):  'rgba(255, 192, 0,0.6)',
            ('Alto','Médio'):  'rgba(204, 255, 204,0.6)',
            ('Alto','Alto'):   'rgba(180, 198, 231,0.6)'
        }


        texto_field1=  "Novo na função, em fase de <br> desenvolvimento. Necessita de <br> avaliação para garantir que esteja <br> adequado ao papel."


        # Descrições indexadas por (Potencial, Performance) => (Categoria_Y, Categoria_X)
        descricoes = {
            # Potencial Alto (linha superior)
            ('Alto','Baixo'): ("DESAFIO", "Novo na função, em fase de<br>desenvolvimento. Necessita de<br>avaliação para garantir que esteja<br>adequado ao papel."),
            ('Alto','Médio'): ("FORTE CULTURA", "Profissional de referência, com<br>potencial para desafios<br>maiores e crescimento em nível<br>hierárquico."),
            ('Alto','Alto'):  ("ALTO POTENCIAL", "Pronto para assumir desafios<br>maiores, até dois níveis acima.<br>Futuro líder estratégico da organização."),

            # Potencial Médio (linha do meio)
            ('Médio','Baixo'): ("DILEMA", "Requer revisão de atribuições,<br>feedback estruturado e plano<br>de ação. Monitorar de perto e<br>avaliar movimentação."),
            ('Médio','Médio'): ("COMPETENTE", "Alcança as expectativas e se<br>adapta a novas situações, mas<br>ainda precisa aprimorar-se<br>para subir de posição."),
            ('Médio','Alto'):  ("FORTE ENTREGA", "Profissional valioso, com<br>espaço para mais<br>responsabilidades e potencial<br>para assumir cargos superiores."),

            # Potencial Baixo (linha inferior)
            ('Baixo','Baixo'): ("BAIXA PERFORMANCE","Necessita de feedback<br>constante e revisão de<br>desempenho a curto prazo.<br>Avaliar realocação em outras funções<br>ou considerar o desligamento."),
            ('Baixo','Médio'): ("INCONSISTENTE", "Cumpre as expectativas,<br>ocasionalmente excedendo<br>quando necessário. Mantém o<br>desempenho esperado."),
            ('Baixo','Alto'):  ("ESPECIALISTA", "Apresenta resultados acima da<br>média, com potencial para<br>movimentação lateral e <br>crescimento na posição."),
        }

        fig = go.Figure()

        # Labels for axes (y invertido para que Alto apareça no topo)
        y_names = ['Alto', 'Médio', 'Baixo']
        x_names = ['Baixo', 'Médio', 'Alto']

        # Draw background and add percentual + descrição
        for y_name in y_names:
            y0, y1 = boundaries[y_name]
            for x_name in x_names:
                x0, x1 = boundaries[x_name]
                fig.add_shape(type='rect', x0=x0, y0=y0, x1=x1, y1=y1,
                            fillcolor=cell_colors[(x_name, y_name)], line=dict(width=1), layer='below')

                # Buscar percentual do quadrante (filtrando por Performance=X e Potencial=Y)
                perc = quadros.loc[(quadros['Categoria_X']==x_name) & (quadros['Categoria_Y']==y_name), 'Perc']
                perc_text = f"{perc.values[0]}%" if len(perc) > 0 else "0%"

                # Buscar descrição correta usando (Potencial, Performance)
                titulo, desc = descricoes.get((y_name, x_name), ("", ""))
                desc = ''

                # Texto final: título, percentual, descrição (a descrição fica abaixo do percentual)
                # Usamos HTML (<br>) para quebra de linhas
                texto =  f"{titulo}<br><br>{perc_text}<br><br>{desc}"

                fig.add_annotation(x=(x0+x1)/2, y=(y0+y1)/2,
                                text=texto, showarrow=False,
                                font=dict(size=12, color="black"),
                                align="center",
                                xanchor='center', 
                                yanchor='middle')

        # Layout
        fig.update_xaxes(range=[0.5, 6.5], tickmode='array', tickvals=[0.5, 1, 1.5 , 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6], ticktext=["","","Baixo (1-2) ", "", "","", "Médio (3-4) ","","", "", "Alto (5) ", ""], title_text='Potencial')
        fig.update_yaxes(range=[0.5, 6.5], tickmode='array', tickvals=[0.5, 1, 1.5 , 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6], ticktext=["","","Baixo (1-2) ", "", "","", "Médio (3-4) ","","", "", "Alto (5) ", ""], title_text='Desempenho')
        fig.update_layout(
                        height=1800,
                        width=1800,
                        hovermode='closest',
                        plot_bgcolor='rgba(0,0,0,0)',   # fundo do gráfico transparente
                        paper_bgcolor='rgba(0,0,0,0)')   # fundo da "folha" transparente)

        fig.update_xaxes(showgrid=False, zeroline=False)
        fig.update_yaxes(showgrid=False, zeroline=False)


        img_bytes = fig.to_image(format="png")
        img64 = base64.b64encode(img_bytes).decode()
        
        #fig.write_html('ninebox.html', include_plotlyjs='cdn')
        #print("Gráfico gerado e salvo em: ninebox.html")

        return img64,''

    except Exception as e:
        print(f'❌ Erro ao gerar gráfico Ninebox: {e}')
        return None,e


def gera_gráfico_pilar(participante = '', pilar= '', competencia=''):
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


        scrp_sql = f'Select Pilar, ROUND(AVG(Resposta)) as Media from QuestRH_Respostas where {sql_condicao} group by Pilar HAVING COUNT(DISTINCT id_rel) = 2'

        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()

        conn.close()

        if len(resultados) == 0:   
            return None, 'Nenhum dado encontrado.'
        
        if pilar == '':
            apurado.append(float(resultados[1]['Media'])) #E
            apurado.append(float(resultados[2]['Media'])) #P
            apurado.append(float(resultados[0]['Media'])) #C
            desejado = [4.0, 4.0, 5.0]

        elif pilar == 'Excelência com Cuidado ':
            apurado.append(float(resultados[0]['Media']))
            apurado.append(0.0)
            apurado.append(0.0)
            desejado = [4.0, 0.0, 0.0]

        elif pilar == 'Protagonismo Empreendedor ':
            apurado.append(0.0)
            apurado.append(float(resultados[0]['Media']))
            apurado.append(0.0)
            desejado = [0.0, 4.0, 0.0]


        elif pilar == 'Criação com Propósito ':
            apurado.append(0.0)
            apurado.append(0.0)
            apurado.append(float(resultados[0]['Media']))
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

def gera_gráfico_Competencia(participante = '', pilar= '', competencia=''):
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


        scrp_sql = f'Select Competencia, ROUND(AVG(Resposta)) as Media from QuestRH_Respostas where {sql_condicao} group by Competencia HAVING COUNT(DISTINCT id_rel) = 2'

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
            apurado.append(float(res['Media']))
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

def gera_gráfico_Comparativo(participante = '', pilar= '', competencia=''):
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
