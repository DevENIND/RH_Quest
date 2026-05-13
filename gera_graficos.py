
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


def gera_ninebox(pilar='', competencia='', participante='', outras_condicoes='', performance=''):
    try:
        # 1. Construção da Condição SQL (Otimizada)
        filtros = ["Status_Av = 'Finalizado'"]
        if pilar: filtros.append(f"Pilar = '{pilar}'")
        if competencia: filtros.append(f"Competencia = '{competencia}'")
        if participante: filtros.append(f"Participante = '{participante}'")
        if outras_condicoes: filtros.append(outras_condicoes)
        
        sql_condicao = " WHERE " + " AND ".join(filtros)

        # SQL unificado (já traz a Performance classificada)
        scrp_sql = f"""
        SELECT * FROM (
            SELECT 
                Participante, Media_Avaliadores, Media_Aval_Desemp,
                CASE 
                    WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (1, 2) THEN 'BAIXA PERFORMANCE'
                    WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (3, 4) THEN 'INCONSISTENTE'
                    WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp = 5 THEN 'ESPECIALISTA'
                    WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (1, 2) THEN 'DILEMA'
                    WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (3, 4) THEN 'COMPETENTE'
                    WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp = 5 THEN 'FORTE ENTREGA'
                    WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (1, 2) THEN 'DESAFIO'
                    WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (3, 4) THEN 'FORTE CULTURA'
                    WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp = 5 THEN 'ALTO POTENCIAL'
                    ELSE 'N/A'
                END as Performance
            FROM (
                SELECT 
                    Participante,
                    CASE WHEN m_av1 = m_av2 THEN m_av1 ELSE CEIL((m_av1 + m_av2) / 2) END AS Media_Avaliadores,
                    CASE WHEN m_d_av1 = m_d_av2 THEN m_d_av1 ELSE CEIL((m_d_av1 + m_d_av2) / 2) END AS Media_Aval_Desemp,
                    Status_Av, C_Custo, Cargo, Grupo_estrategico, Sigla_Emp
                    
                FROM (
                    SELECT 
                        Participante, 
                        ROUND(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as m_av1,
                        ROUND(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as m_d_av1,
                        ROUND(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as m_av2,
                        ROUND(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as m_d_av2,
                        CASE WHEN (COUNT(DISTINCT CASE WHEN id_rel IN (1,2) THEN id_rel END) = 2) THEN 'Finalizado' ELSE 'Pendente' END as Status_Av,
                        Sigla_Emp, C_Custo, Cargo, Grupo_estrategico
                    FROM QuestRH_Respostas 
                    GROUP BY Participante
                ) as t
            ) as x {sql_condicao}
        ) as final_query
        """
        
        if performance:
            scrp_sql += f" WHERE Performance = '{performance}'"

        # 2. Execução
        conn = mysql_connection()
        df = pd.read_sql(scrp_sql, conn)
        conn.close()

        if df.empty:
            return None, "Nenhum dado encontrado"

        # 3. Mapeamentos Estáticos (Otimizado)
        mapeamento = {
            'BAIXA PERFORMANCE': ('Baixo', 'Baixo', (1.0, 0.3, 0.3, 0.6)),
            'INCONSISTENTE':     ('Médio', 'Baixo', (1.0, 0.7, 0.4, 0.6)),
            'ESPECIALISTA':      ('Alto',  'Baixo', (1.0, 0.75, 0.0, 0.6)),
            'DILEMA':            ('Baixo', 'Médio', (1.0, 0.7, 0.4, 0.6)),
            'COMPETENTE':        ('Médio', 'Médio', (1.0, 0.75, 0.0, 0.6)),
            'FORTE ENTREGA':     ('Alto',  'Médio', (0.8, 1.0, 0.8, 0.6)),
            'DESAFIO':           ('Baixo', 'Alto',  (1.0, 0.75, 0.0, 0.6)),
            'FORTE CULTURA':     ('Médio', 'Alto',  (0.8, 1.0, 0.8, 0.6)),
            'ALTO POTENCIAL':    ('Alto',  'Alto',  (0.7, 0.8, 0.9, 0.6))
        }

        # Cálculo de percentuais direto no DF
        n_total = len(df)
        contagem = df['Performance'].value_counts(normalize=True) * 100

        # 4. Geração do Gráfico
        fig, ax = plt.subplots(figsize=(8, 8))
        boundaries = {"Baixo": (0.5, 2.5), "Médio": (2.5, 4.5), "Alto": (4.5, 6.5)}

        for label, (x_cat, y_cat, color) in mapeamento.items():
            x0, x1 = boundaries[x_cat]
            y0, y1 = boundaries[y_cat]
            
            # Desenha o quadrante
            ax.fill_between([x0, x1], [y0, y0], [y1, y1], color=color, edgecolor="black", linewidth=0.5)
            
            # Texto (Título + %)
            perc = contagem.get(label, 0.0)
            ax.text((x0+x1)/2, (y0+y1)/2, f"{label}\n{perc:.1f}%", 
                    ha="center", va="center", fontsize=9, weight="bold")

        # Estilização final
        ax.set_xlim(0.5, 6.5); ax.set_ylim(0.5, 6.5)
        ax.set_xticks([1.5, 3.5, 5.5]); ax.set_xticklabels(["Baixo", "Médio", "Alto"])
        ax.set_yticks([1.5, 3.5, 5.5]); ax.set_yticklabels(["Baixo", "Médio", "Alto"])
        ax.set_title("Nine Box Performance", fontsize=14, pad=20, weight="bold")
        
        plt.tight_layout()
        
        # Base64
        buf = io.BytesIO()
        fig.savefig(buf, format="png", transparent=True)
        plt.close(fig)
        return base64.b64encode(buf.getvalue()).decode(), None

    except Exception as e:
        return None, str(e)

def gera_gráfico_pilar(participante='', pilar='', competencia='', outras_condicoes='', performance=''):
    try:
        # 1. Construção da Condição SQL
        filtros = ["Status_Av_Interno = 'Finalizado'"] # Alterado para o nome real do cálculo
        if pilar: filtros.append(f"Pilar = '{pilar}'")
        if competencia: filtros.append(f"Competencia = '{competencia}'")
        if participante: filtros.append(f"Participante = '{participante}'")
        if outras_condicoes: filtros.append(outras_condicoes)
        
        sql_condicao = " WHERE " + " AND ".join(filtros)

        scrp_sql = f"""
        SELECT Pilar, ROUND(AVG(Media_Geral), 0) as Media_Geral 
        FROM (
            SELECT * FROM (
                SELECT 
                    Pilar,
                    Participante, 
                    CASE WHEN m_av1 = m_av2 THEN m_av1 ELSE CEIL((m_av1 + m_av2) / 2) END AS Media_Geral,
                    CASE 
                        WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (1, 2) THEN 'BAIXA PERFORMANCE'
                        WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (3, 4) THEN 'INCONSISTENTE'
                        WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp = 5 THEN 'ESPECIALISTA'
                        WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (1, 2) THEN 'DILEMA'
                        WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (3, 4) THEN 'COMPETENTE'
                        WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp = 5 THEN 'FORTE ENTREGA'
                        WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (1, 2) THEN 'DESAFIO'
                        WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (3, 4) THEN 'FORTE CULTURA'
                        WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp = 5 THEN 'ALTO POTENCIAL'
                        ELSE 'N/A'
                    END as Performance,
                    Status_Av_Interno as Status_Av
                FROM (
                    SELECT 
                        Pilar, Participante,
                        CASE WHEN m_av1 = m_av2 THEN m_av1 ELSE CEIL((m_av1 + m_av2) / 2) END AS Media_Avaliadores,
                        CASE WHEN m_d_av1 = m_d_av2 THEN m_d_av1 ELSE CEIL((m_d_av1 + m_d_av2) / 2) END AS Media_Aval_Desemp,
                        m_av1, m_av2, Status_Av_Interno, Sigla_Emp, C_Custo, Cargo, Grupo_estrategico
                    FROM (
                        SELECT 
                            Pilar, Participante, 
                            ROUND(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as m_av1,
                            ROUND(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as m_d_av1,
                            ROUND(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as m_av2,
                            ROUND(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as m_d_av2,
                            CASE 
                                WHEN (COUNT(DISTINCT CASE WHEN id_rel IN (1,2) THEN id_rel END) = 2) 
                                THEN 'Finalizado' ELSE 'Pendente' 
                            END as Status_Av_Interno, Sigla_Emp, C_Custo, Cargo, Grupo_estrategico
                        FROM QuestRH_Respostas 
                        GROUP BY Pilar, Participante
                    ) as t
                    {sql_condicao} -- Filtro aplicado onde a coluna 'Status_Av_Interno' já existe
                ) as x
            ) as final_query
            {f"WHERE Performance = '{performance}'" if performance else ""}
        ) as base_calculo
        GROUP BY Pilar
        """
        
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        res = cursor.fetchall()
        conn.close()

        if not res:
            return None, 'Nenhum dado encontrado.'

        # 2. Mapeamento Inteligente (Evita erro de índice [1], [2])
        # Criamos um dicionário onde a chave é o nome do pilar
        mapa_resultados = {r['Pilar'].strip(): float(r['Media_Geral']) for r in res}

        # Nomes exatos dos pilares conforme seu banco (ajuste espaços se necessário)
        p_excelencia = "Excelência com Cuidado"
        p_protagonismo = "Protagonismo Empreendedor"
        p_criacao = "Criação com Propósito"

        # Buscar valores do mapa ou usar 0.0 se não existir
        v_e = mapa_resultados.get(p_excelencia, 0.0)
        v_p = mapa_resultados.get(p_protagonismo, 0.0)
        v_c = mapa_resultados.get(p_criacao, 0.0)

        # Definir Apurado e Desejado
        apurado = [v_e, v_p, v_c]
        
        if not pilar:
            desejado = [4.0, 4.0, 5.0]
        else:
            # Se um pilar específico foi filtrado, zeramos os desejados dos outros
            desejado = [
                4.0 if pilar.strip() == p_excelencia else 0.0,
                4.0 if pilar.strip() == p_protagonismo else 0.0,
                5.0 if pilar.strip() == p_criacao else 0.0
            ]

        # 3. Gerar Gráfico (mesma lógica sua, mas com proteção de dados)
        categorias = ["Excelência", "Protagonismo", "Criação"]
        gap = [a - b for a, b in zip(apurado, desejado)]

        x_pos = np.arange(len(categorias))
        largura = 0.25

        fig, ax = plt.subplots(figsize=(10, 5))
        # 1. Gerar as barras
        b1 = ax.bar(x_pos - largura, desejado, largura, label="Desejado", color="#4472C4")
        b2 = ax.bar(x_pos, apurado, largura, label="Apurado", color="#ED7D31")
        b3 = ax.bar(x_pos + largura, gap, largura, label="GAP", color="gray")

        # 2. Adicionar os rótulos automaticamente
        # fmt='%.0f' define que não terá casas decimais. Use '%.1f' se quiser uma casa.
        ax.bar_label(b1, padding=3, fmt='%.0f', fontsize=9)
        ax.bar_label(b2, padding=3, fmt='%.0f', fontsize=9)
        ax.bar_label(b3, padding=3, fmt='%.0f', fontsize=9)

        # 3. Ajustes de layout
        ax.set_title("Comparativo Pilares", fontsize=14, fontweight="bold", pad=20)
        ax.set_xticks(x_pos)
        ax.set_xticklabels(categorias)
        
        # Aumentar um pouco o limite superior para o rótulo não encostar no topo
        ax.set_ylim(min(gap + [0]) - 1, max(desejado + apurado) + 2)
        
        ax.legend(loc="upper right", bbox_to_anchor=(1.2, 1))
        ax.axhline(0, color="black", linewidth=0.8)

        plt.tight_layout()
        buf = io.BytesIO()
        plt.savefig(buf, format="png", transparent=True)
        plt.close(fig)
        
        return base64.b64encode(buf.getvalue()).decode("utf-8"), ''

    except Exception as e:
        print(f"Erro: {e}") # Log para você ver no console
        return None, str(e)

def gera_gráfico_Competencia(participante='', pilar='', competencia='', outras_condicoes='', performance=''):
    try:
        # 1. Mapeamento de Metas (Desejado)
        desejado_dict = {
            'Equilibrio emocional': 4.0, 'Comprometimento com resultado': 4.0,
            'Resiliência': 4.0, 'Faz como se fosse seu': 4.0,
            'Foco em soluções': 4.0, 'Atualização e inovação constantes': 4.0,
            'Inspiração e mobilização de pessoas': 5.0, 'Trabalho em equipe': 5.0,
            'Ensinar e compartilhar conhecimento': 5.0, 'Resolve, não empurra': 4.0,
            'Entrega com precisão': 4.0, 'Joga junto': 4.0, 'Melhora sempre': 5.0,
        }

        # 2. Construção dos Filtros Dinâmicos
        filtros = ["Status_Av_Interno = 'Finalizado'"]
        if pilar: filtros.append(f"Pilar = '{pilar}'")
        if competencia: filtros.append(f"Competencia = '{competencia}'")
        if participante: filtros.append(f"Participante = '{participante}'")
        if outras_condicoes: filtros.append(outras_condicoes)
        
        sql_condicao = " WHERE " + " AND ".join(filtros)

        # 3. SQL com Suporte a Filtro de Performance
        scrp_sql = f"""
        SELECT Competencia, ROUND(AVG(Media_Geral), 0) as Media_Geral 
        FROM (
            SELECT * FROM (
                SELECT 
                    Competencia, Participante, 
                    CASE WHEN m_av1 = m_av2 THEN m_av1 ELSE CEIL((m_av1 + m_av2) / 2) END AS Media_Geral,
                    CASE 
                        WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (1, 2) THEN 'BAIXA PERFORMANCE'
                        WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp IN (3, 4) THEN 'INCONSISTENTE'
                        WHEN Media_Avaliadores IN (1, 2) AND Media_Aval_Desemp = 5 THEN 'ESPECIALISTA'
                        WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (1, 2) THEN 'DILEMA'
                        WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp IN (3, 4) THEN 'COMPETENTE'
                        WHEN Media_Avaliadores IN (3, 4) AND Media_Aval_Desemp = 5 THEN 'FORTE ENTREGA'
                        WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (1, 2) THEN 'DESAFIO'
                        WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp IN (3, 4) THEN 'FORTE CULTURA'
                        WHEN Media_Avaliadores = 5 AND Media_Aval_Desemp = 5 THEN 'ALTO POTENCIAL'
                        ELSE 'N/A'
                    END as Performance,
                    Status_Av_Interno
                FROM (
                    SELECT 
                        Competencia, Pilar, Participante,
                        CASE WHEN m_av1 = m_av2 THEN m_av1 ELSE CEIL((m_av1 + m_av2) / 2) END AS Media_Avaliadores,
                        CASE WHEN m_d_av1 = m_d_av2 THEN m_d_av1 ELSE CEIL((m_d_av1 + m_d_av2) / 2) END AS Media_Aval_Desemp,
                        m_av1, m_av2, Status_Av_Interno, Sigla_Emp, C_Custo, Cargo, Grupo_estrategico
                    FROM (
                        SELECT 
                            Competencia, Pilar, Participante, 
                            ROUND(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as m_av1,
                            ROUND(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as m_d_av1,
                            ROUND(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as m_av2,
                            ROUND(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as m_d_av2,
                            CASE 
                                WHEN (COUNT(DISTINCT CASE WHEN id_rel IN (1,2) THEN id_rel END) = 2) 
                                THEN 'Finalizado' ELSE 'Pendente' 
                            END as Status_Av_Interno,
                            Sigla_Emp, C_Custo, Cargo, Grupo_estrategico
                        FROM QuestRH_Respostas 
                        GROUP BY Competencia, Pilar, Participante
                    ) as t
                    {sql_condicao}
                ) as x
            ) as final_query
            {f"WHERE Performance = '{performance}'" if performance else ""}
        ) as base_calculo
        GROUP BY Competencia
        ORDER BY Competencia
        """

        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()
        conn.close()

        if not resultados:
            return None, 'Nenhum dado encontrado.'

        # 4. Processamento dos Dados para o Gráfico
        labels_apurado = []
        apurado = []
        desejado = []

        for res in resultados:
            comp_nome = str(res['Competencia']).strip()
            val_apurado = float(res['Media_Geral'])
            val_desejado = desejado_dict.get(comp_nome, 0.0) # 0.0 caso não ache no dict
            
            labels_apurado.append(comp_nome)
            apurado.append(val_apurado)
            desejado.append(val_desejado)

        gap = np.array(apurado) - np.array(desejado)
        x_pos = np.arange(len(labels_apurado))
        largura = 0.25

        # 5. Geração do Gráfico
        fig, ax = plt.subplots(figsize=(15, 6))

        b1 = ax.bar(x_pos - largura, desejado, largura, label="Desejado", color="#6DDAE9")
        b2 = ax.bar(x_pos, apurado, largura, label="Apurado", color="#0353FF")
        b3 = ax.bar(x_pos + largura, gap, largura, label="GAP", color="gray")

        # Rótulos das barras (Aparecer valores)
        ax.bar_label(b1, padding=3, fmt='%.0f', fontsize=8)
        ax.bar_label(b2, padding=3, fmt='%.0f', fontsize=8)
        ax.bar_label(b3, padding=3, fmt='%.0f', fontsize=8)

        ax.set_title("Comparativo por Competência", fontsize=14, fontweight="bold", pad=20)
        ax.set_xticks(x_pos)
        ax.set_xticklabels(labels_apurado, rotation=45, ha="right", fontsize=9)
        
        # Ajuste de escala Y
        y_vals = list(desejado) + list(apurado) + list(gap)
        ax.set_ylim(min(y_vals) - 1, max(y_vals) + 1.5)
        
        ax.axhline(0, color="black", linewidth=0.8)
        ax.legend(loc="upper left", bbox_to_anchor=(1, 1))

        plt.tight_layout()

        # Converter para Base64
        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=120, bbox_inches="tight", transparent=True)
        plt.close(fig)
        return base64.b64encode(buf.getvalue()).decode("utf-8"), ''

    except Exception as e:
        print(f"❌ Erro Competências: {e}")
        return None, str(e)

def gera_gráfico_Comparativo(participante='', pilar='', competencia='', outras_condicoes='', performance=''):
    try:
        # 1. Construção dos Filtros Dinâmicos
        # Nota: Aqui usamos id_rel IN (0,1,2) pois precisamos da Auto Avaliação (0)
        filtros = ["Status_Av_Interno = 'Finalizado'"]
        if pilar: filtros.append(f"Pilar = '{pilar}'")
        if competencia: filtros.append(f"Competencia = '{competencia}'")
        if participante: filtros.append(f"Participante = '{participante}'")
        if outras_condicoes: filtros.append(outras_condicoes)
        
        sql_condicao = " WHERE " + " AND ".join(filtros)

        # 2. SQL Otimizado com Auto Avaliação e Filtro de Performance
        scrp_sql = f"""
SELECT Competencia, 
       ROUND(AVG(Media_Auto), 0) as Media_Auto, 
       ROUND(AVG(Media_Gestores), 0) as Media_Gestores
FROM (
    SELECT 
        Competencia, 
        Participante, 
        Media_Auto,
        Media_Avaliadores_Final as Media_Gestores,
        -- Classificação de Performance baseada nas médias calculadas
        CASE 
            WHEN Media_Avaliadores_Final IN (1, 2) AND Media_Aval_Desemp IN (1, 2) THEN 'BAIXA PERFORMANCE'
            WHEN Media_Avaliadores_Final IN (1, 2) AND Media_Aval_Desemp IN (3, 4) THEN 'INCONSISTENTE'
            WHEN Media_Avaliadores_Final IN (1, 2) AND Media_Aval_Desemp = 5 THEN 'ESPECIALISTA'
            WHEN Media_Avaliadores_Final IN (3, 4) AND Media_Aval_Desemp IN (1, 2) THEN 'DILEMA'
            WHEN Media_Avaliadores_Final IN (3, 4) AND Media_Aval_Desemp IN (3, 4) THEN 'COMPETENTE'
            WHEN Media_Avaliadores_Final IN (3, 4) AND Media_Aval_Desemp = 5 THEN 'FORTE ENTREGA'
            WHEN Media_Avaliadores_Final = 5 AND Media_Aval_Desemp IN (1, 2) THEN 'DESAFIO'
            WHEN Media_Avaliadores_Final = 5 AND Media_Aval_Desemp IN (3, 4) THEN 'FORTE CULTURA'
            WHEN Media_Avaliadores_Final = 5 AND Media_Aval_Desemp = 5 THEN 'ALTO POTENCIAL'
            ELSE 'N/A'
        END as Performance
    FROM (
        SELECT 
            Competencia, 
            Participante,
            ROUND(m_auto_raw, 0) AS Media_Auto,
            CASE WHEN m_av1 = m_av2 THEN m_av1 ELSE CEIL((m_av1 + m_av2) / 2) END AS Media_Avaliadores_Final,
            CASE WHEN m_d_av1 = m_d_av2 THEN m_d_av1 ELSE CEIL((m_d_av1 + m_d_av2) / 2) END AS Media_Aval_Desemp,
            Status_Av_Interno,
            Sigla_Emp, C_Custo, Cargo, Grupo_estrategico
        FROM (
            SELECT 
                Competencia, Pilar, Participante, 
                AVG(CASE WHEN id_rel = 0 THEN Resposta END) as m_auto_raw,
                ROUND(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as m_av1,
                ROUND(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as m_d_av1,
                ROUND(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as m_av2,
                ROUND(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as m_d_av2,
                -- A coluna id_rel é usada aqui para o COUNT, mas não precisa estar no SELECT externo
                CASE 
                    WHEN (COUNT(DISTINCT id_rel) = 3) THEN 'Finalizado' 
                    ELSE 'Pendente' 
                END as Status_Av_Interno,
                Sigla_Emp, C_Custo, Cargo, Grupo_estrategico
            FROM QuestRH_Respostas 
            GROUP BY Competencia, Pilar, Participante
        ) as t
        {sql_condicao}
    ) as x
    {f"WHERE Performance = '{performance}'" if performance else ""}
) as base_calculo
GROUP BY Competencia
ORDER BY Competencia
"""

        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()
        conn.close()

        if not resultados:
            return None, 'Nenhum dado encontrado.'

        # 3. Preparação dos Dados
        labels = [str(r['Competencia']).strip() for r in resultados]
        auto_vals = [float(r['Media_Auto']) for r in resultados]
        gestor_vals = [float(r['Media_Gestores']) for r in resultados]

        x_pos = np.arange(len(labels))
        largura = 0.35

        # 4. Geração do Gráfico
        fig, ax = plt.subplots(figsize=(15, 6))
        
        b1 = ax.bar(x_pos - largura/2, auto_vals, largura, label="Auto Avaliação", color="#0BE1FD")
        b2 = ax.bar(x_pos + largura/2, gestor_vals, largura, label="Avaliação Gestores", color="#0642C4")

        # Rótulos automáticos
        ax.bar_label(b1, padding=3, fmt='%.0f', fontsize=9)
        ax.bar_label(b2, padding=3, fmt='%.0f', fontsize=9)

        # Estilização
        ax.set_title("Comparativo: Auto Avaliação vs. Gestores", fontsize=14, fontweight="bold", pad=20)
        ax.set_xticks(x_pos)
        ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=10)
        
        # Ajuste do eixo Y para os números não sumirem
        max_val = max(max(auto_vals), max(gestor_vals))
        ax.set_ylim(0, max_val + 1.5)

        ax.legend(loc="upper left", bbox_to_anchor=(1, 1))
        ax.axhline(0, color="black", linewidth=0.8)

        plt.tight_layout()

        # Converter para Base64
        buf = io.BytesIO()
        plt.savefig(buf, format="png", dpi=120, bbox_inches="tight", transparent=True)
        plt.close(fig)
        return base64.b64encode(buf.getvalue()).decode("utf-8"), ''

    except Exception as e:
        print(f"❌ Erro Comparativo: {e}")
        return None, str(e)
    
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
            scrp_participantes = f"Select count(Participante) from QuestRH_Relacoes Where Tipo_Avaliacao <>'A3';"
            scrp_avaliadores = f"Select count(Participante) as Total from QuestRH_Relacoes"
            connection = mysql_connection()
            cursor = connection.cursor()
            
            cursor.execute(scrp_participantes)
            qtd_participantes = cursor.fetchone()[0]
            
            cursor.execute(scrp_avaliadores)
            qtd_avaliadores = cursor.fetchone()[0]
            
            connection.close()
            
            #qtd participantes = todos participantes + avaliadores A1 e A2 por isso consta qtd_avaliadores*2
            qtd_participantes += qtd_avaliadores*2  
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
        
        scrp_sql2 = f"""SELECT 
    y.Sigla_Emp, 
    count(x.Participante) as Cont
FROM QuestRH_Relacoes AS x
INNER JOIN QuestRH_Pessoas AS y ON y.Nome = x.Participante
Group by y.Sigla_Emp;"""        
        
        
      

        conn = mysql_connection()

        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()

        total_empresas = cursor.execute(scrp_sql2)
        total_empresas = cursor.fetchall()
        Listagem_Avaliacoes = {}
        labels = []
        Textos = []
        Percentual = []
        
        for row in total_empresas:
            Listagem_Avaliacoes[row['Sigla_Emp']] = row['Cont']
            labels.append(row['Sigla_Emp'])
            Textos.append(f"0 de {row['Cont']} (0%)")
            
        conn.close()
        

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
