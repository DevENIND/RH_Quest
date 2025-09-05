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


def gera_grafico_conclusao(finalizados = False):
    try:

        sql_condicao =''

        if finalizados == False:
            agrupar = 'x.Participante, x.Nome_Avaliador'
            texto_grafico = 'Avaliações Finalizadas'
            qtd_participantes = 391 #Quantidade constante na tabela Lista_Emails
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


        scrp_sql = f"""
                Select Sum(Grupo) as Realizado from
                (SELECT count(concat(Participante, ' - ', Nome_Avaliador)) as Grupo 
                    From (SELECT Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo, Nome_Avaliador,
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
                                    Status_Av
                                FROM (
                                    SELECT 
                                        Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo, Nome_Avaliador,
                                        ROUND(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) AS Media_Auto,
                                        ROUND(AVG(CASE WHEN id_rel IN (1,2) THEN Resposta END), 0) AS Media_Avaliadores,
                                        ROUND(AVG(CASE WHEN id_rel = 0 THEN Desempenho_Tecnico END), 0) AS Media_Auto_Desemp,
                                        ROUND(AVG(CASE WHEN id_rel IN (1,2) THEN Desempenho_Tecnico END), 0) AS Media_Aval_Desemp,
                                        CASE 
                                            WHEN (COUNT(DISTINCT id_rel) = 2 AND Avaliacao = 'A3') 
                                            OR (COUNT(DISTINCT id_rel) = 3 AND Avaliacao <> 'A3') 
                                            THEN 'Finalizado' 
                                            ELSE 'Pendente' 
                                        END as Status_Av
                                    FROM QuestRH_Respostas where id_rel in (1,2)
                                    GROUP BY Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo
                                ) t
                            ) x {sql_condicao} group by {agrupar}
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

        if finalizados == False:
            plt.savefig("gauge.png", bbox_inches="tight")
        else:
            plt.savefig("gauge_finalizado.png", bbox_inches="tight")
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
        SELECT Performance, COUNT(Participante) as Contagem From (SELECT Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
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
                    Status_Av
                FROM (
                    SELECT 
                        Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo,
                        ROUND(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) AS Media_Auto,
                        ROUND(AVG(CASE WHEN id_rel IN (1,2) THEN Resposta END), 0) AS Media_Avaliadores,
                        ROUND(AVG(CASE WHEN id_rel = 0 THEN Desempenho_Tecnico END), 0) AS Media_Auto_Desemp,
                        ROUND(AVG(CASE WHEN id_rel IN (1,2) THEN Desempenho_Tecnico END), 0) AS Media_Aval_Desemp,
                        CASE 
                            WHEN (COUNT(DISTINCT id_rel) = 2 AND Avaliacao = 'A3') 
                            OR (COUNT(DISTINCT id_rel) = 3 AND Avaliacao <> 'A3') 
                            THEN 'Finalizado' 
                            ELSE 'Pendente' 
                        END as Status_Av
                    FROM QuestRH_Respostas {sql_condicao}
                    GROUP BY Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo
                ) t
            ) x where x.Performance <> '' {sql_condicao_final} group by x.Performance order by Contagem
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
        if finalizados == True:
            plt.savefig("potencial_finalizados.png", bbox_inches="tight")
        else:
            plt.savefig("potencial.png", bbox_inches="tight")

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
                    Status_Av
                FROM (
                    SELECT 
                        Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo,
                        ROUND(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) AS Media_Auto,
                        ROUND(AVG(CASE WHEN id_rel IN (1,2) THEN Resposta END), 0) AS Media_Avaliadores,
                        ROUND(AVG(CASE WHEN id_rel = 0 THEN Desempenho_Tecnico END), 0) AS Media_Auto_Desemp,
                        ROUND(AVG(CASE WHEN id_rel IN (1,2) THEN Desempenho_Tecnico END), 0) AS Media_Aval_Desemp,
                        CASE 
                            WHEN (COUNT(DISTINCT id_rel) = 2 AND Avaliacao = 'A3') 
                            OR (COUNT(DISTINCT id_rel) = 3 AND Avaliacao <> 'A3') 
                            THEN 'Finalizado' 
                            ELSE 'Pendente' 
                        END as Status_Av
                    FROM QuestRH_Respostas 
                    GROUP BY Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo
                ) t
            ) x group by Sigla_Emp order by Sigla_Emp
        """
        Valores = [0,0,0,0,0,0]
        Textos=['0 de 250','0 de 3','0 de 40','0 de 73','0 de 23','0 de 1']
        Percentual = []

        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()

        conn.close()
        

        Listagem_Avaliacoes ={
            'Construção': 250,
            'Locações': 3,
            'Montagens Industriais': 40,
            'Serviços': 73,
            "Participações": 23,
            'Construtivo': 1
        }

        # Dados
        labels = [
            "Construção",
            "Construtivo",
            "Locações",
            "Montagens Industriais",
            "Participações",
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
            "blue",      # Construção
            "yellow",     # Construtivo
            "orange",        # Locações
            "green",       # Montagens Industriais
            "purple",        # Participações
            "red",        # Serviços
        ]

        # Plot
        fig, ax = plt.subplots(figsize=(8, 5))

        y_pos = range(len(labels))

        # Paleta mais suave (pastéis)
        colors = [
            "#6baed6",  # azul claro
            "#ffd966",  # amarelo suave
            "#fdae6b",  # laranja suave
            "#74c476",  # verde suave
            "#bcbddc",  # lilás
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
        plt.savefig("barra_empresas.png", bbox_inches="tight", dpi=150)

        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", transparent=True)
        plt.close(fig)
        buf.seek(0)


        return base64.b64encode(buf.getvalue()).decode(), None
    except Exception as e:
        print(f'❌ Erro ao gerar gráfico de barras de empresas: {e}')
        return None, e


def main(page: ft.Page):
    def atualizar_graficos():
        grafico64_potencial, erro = gera_gráfico_potencial(finalizados=True)
        grafico64_potencial_nao_finalizados, erro = gera_gráfico_potencial()
        grafico64_concluidos, erro, realizado, nao_concluidos = gera_grafico_conclusao()
        grafico64_finalizados, erro, finalizados, nao_finalizados = gera_grafico_conclusao(finalizados=True)
        grafico64_empresas, erro = gera_grafico_empresas()

        if grafico64_potencial:
            img_Potencial.src_base64 = grafico64_potencial
        if grafico64_potencial_nao_finalizados:
            img_Potencial_nao_finalizados.src_base64 = grafico64_potencial_nao_finalizados
        if grafico64_concluidos:
            img_Concluidos.src_base64 = grafico64_concluidos
        if grafico64_finalizados:
            img_Finalizados.src_base64 = grafico64_finalizados
        if grafico64_empresas:
            img_Barra_Empresa.src_base64 = grafico64_empresas

        txt_Finalizados.value = str(finalizados or 0)
        txt_Realizados.value = str(realizado or 0)
        txt_Pendentes.value = str(nao_concluidos or 0)
        txt_Pendentes_participantes.value = str(nao_finalizados or 0)

        page.update()
    
    # imagens
    img_Potencial = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_Potencial_nao_finalizados = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_Concluidos = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_Finalizados = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_Barra_Empresa = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    
    # textos -> Referencia de tamanhos de fontes =https://flet.dev/docs/controls/text/#font_family
    txt_Realizados = ft.Text('', size=40, weight='w200', font_family='roboto',text_align="center", color=ft.Colors.GREEN)
    txt_Finalizados = ft.Text('', size=40, weight='w200', font_family='roboto',text_align="center",color=ft.Colors.BLUE)
    txt_Pendentes = ft.Text('', size=40, weight='w200', font_family='roboto',text_align="center",color=ft.Colors.RED)
    txt_Pendentes_participantes = ft.Text('', size=40, weight='w200', font_family='roboto',text_align="center",color=ft.Colors.AMBER)

    # containers
    container_realizados = ft.Container(
        content=ft.Column([
            txt_Realizados,
            ft.Text('Avaliações Realizadas',weight='w200', size=15, font_family='roboto', color=ft.Colors.GREEN)
        ],alignment="center", horizontal_alignment="center", spacing=5),
        expand=3,
        padding=20,
        border_radius=12,
        border=ft.border.all(1, "green"),
        shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.GREEN_100, offset=ft.Offset(0,2)),
        width=200
    )

    container_pendentes = ft.Container(
        content=ft.Column([
            txt_Pendentes,
            ft.Text('Avaliações Pendentes',weight='w200', size=15, font_family='roboto', color=ft.Colors.RED)
        ], alignment="center", horizontal_alignment="center", spacing=5),
        expand=3,
        padding=20,
        border_radius=12,
        border=ft.border.all(1, "red"),
        shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.RED_100, offset=ft.Offset(0,2)),
        width=200
    )
    
    container_finalizados = ft.Container(
        content=ft.Column([
            txt_Finalizados,
            ft.Text('Participantes Finalizados',weight='w200', size=15, font_family='roboto', color=ft.Colors.BLUE)
        ], alignment="center", horizontal_alignment="center", spacing=5),
        expand=3,
        padding=20,
        border_radius=12,
        border=ft.border.all(1, "blue"),
        shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.BLUE_100, offset=ft.Offset(0,2)),
        width=200
    )
    
    container_pendentes_participantes = ft.Container(
        content=ft.Column([
            txt_Pendentes_participantes,
            ft.Text('Participantes Pendentes',weight='w200', size=15, font_family='roboto', color=ft.Colors.AMBER),
        ], alignment="center", horizontal_alignment="center", spacing=5),
        expand=3,
        padding=20,
        border_radius=12,
        border=ft.border.all(1, "amber"),
        shadow=ft.BoxShadow(blur_radius=8, color=ft.Colors.AMBER_100, offset=ft.Offset(0,2)),
        width=200
    )

    # agrupando os cards
    container_texto_grafico = ft.Row([
        container_realizados,
        container_pendentes,
        container_finalizados,
        container_pendentes_participantes,
        
    ])

    # painel principal
    container_painel_grafico = ft.Column([
        ft.Row([ft.Text("📊 Painel Controle Gráfico", size=20, weight="bold")]),
        ft.Divider(),
        container_texto_grafico,
        ft.Row([
            # coluna da esquerda (2 linhas)
            ft.Column([
                ft.Row([
                    ft.Container(img_Concluidos, expand=1),
                    ft.Container(img_Potencial_nao_finalizados, expand=2)
                ]),
                ft.Row([
                    ft.Container(img_Finalizados, expand=1),
                    ft.Container(img_Potencial, expand=2),
                ])
            ],expand=2),

             # coluna da direita (um único gráfico maior)
            ft.Column([
                ft.Row([
                    ft.Container(img_Barra_Empresa, expand=2)
                ])
            ],expand=2)
        ]),
    ])
    
    atualizar_graficos()
    page.scroll = ft.ScrollMode.AUTO
    page.theme_mode = ft.ThemeMode.LIGHT
    page.add(container_painel_grafico)

ft.app(target=main, view=ft.WEB_BROWSER)
