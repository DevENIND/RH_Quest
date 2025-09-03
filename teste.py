import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.stats import norm
import io
import base64
import flet as ft
import pymysql

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

# ---------- FUNÇÕES ----------
def gera_bell_curve(dados):

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

            fig, ax = plt.subplots(figsize=(7,5))
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
    return base64.b64encode(buf.read()).decode("utf-8")




# ---------- FUNÇÃO PRINCIPAL FLET ----------

def main(page: ft.Page):
    page.title = "Bell Curve"
    page.window.center()

    # Controles globais para atualização
    tabela_bell = ft.ListView(expand=True, auto_scroll=False, spacing=5, padding=10)

    grafico = ft.Image(expand=True, fit=ft.ImageFit.CONTAIN)

    def busca_dados_bell():
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
        return pessoas

    # Função para atualizar tabela e gráfico
    def atualiza_dados(e=None):
        pessoas = busca_dados_bell()
        notas = [float(p['Media']) for p in pessoas if p['Media'] is not None]
             
        # Atualiza gráfico
        img_b64 = gera_bell_curve(notas)
        grafico.src_base64 = img_b64
        grafico.update()

    def atualiza_visualização_bell(e=None):
        if grafico.visible:
            grafico.visible = False
            container_tabela_bell.visible = True
            btn_ver_grafico.icon = ft.Icons.BAR_CHART_ROUNDED
            btn_ver_grafico.text = "Ver Gráfico"
        else:
            grafico.visible = True
            container_tabela_bell.visible = False
            btn_ver_grafico.icon = ft.Icons.DATA_ARRAY_OUTLINED
            btn_ver_grafico.text = "Ver Tabela"


        grafico.update()
        tabela_bell.update()
        page.update()

    # Botão de atualização
    btn_atualizar = ft.ElevatedButton("Atualizar", on_click=atualiza_dados)
    btn_ver_grafico = ft.ElevatedButton("Ver Gráfico", icon=ft.Icons.DATA_ARRAY_OUTLINED, on_click=lambda e: atualiza_visualização_bell(e))


        # Cabeçalho fixo
    cabecalho_tabela = ft.Container(
        content= ft.Row([
            ft.Container(ft.Text("Participante", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Notas", weight=ft.FontWeight.BOLD), expand=3)
        ],alignment=ft.MainAxisAlignment.CENTER),
        padding=5,
        border_radius=ft.border_radius.only(top_left=10, top_right=10),
    )
    
    container_tabela_bell = ft.Column(
        controls=[
            cabecalho_tabela,
            tabela_bell
         ], 
        visible = False)

    container_bell = ft.Column(
        controls=[
            btn_ver_grafico,
            ft.Column([container_tabela_bell, grafico]),
        ])
    
    # Layout
    container_geral = ft.Column(
        controls=[
            ft.Text("Bell Curve", weight="bold", size=18),
            ft.Divider(),
            btn_atualizar,
            container_bell
        ]
    )

    page.add(container_geral)

    # Inicializa com dados
    atualiza_dados()

ft.app(target=main)
