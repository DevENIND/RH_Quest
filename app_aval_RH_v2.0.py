import flet as ft
import pymysql
from datetime import datetime
from pytz import timezone
import time
import threading
import os
import platform
import pandas as pd
import asyncio
import base64
from pathlib import Path
import io
import gera_graficos
import pytz
from xhtml2pdf import pisa

import os
import json
import sys

#Instalando aplicativo para local Windows
#pyinstaller --noconfirm --onefile --add-data "Imagem_Quest.png;." --icon=enind.ico  app_aval_RH.py


'''
Comandos de atulização:

Atualização do git hub
ssh root@191.252.219.242

ativar o ambiente virtual
source venv/bin/activate

cd RH_Quest
git pull origin main
cd

#################################################################

criação de um ambiente flet:
sudo nano /etc/systemd/system/fletapp.service

colocar no script:

[Unit]
Description=Flet App RH Quest
After=network.target

[Service]
User=root
Group=root
WorkingDirectory=/root/RH_Quest
ExecStart=/root/venv/bin/python /root/RH_Quest/app_aval_RH_v1.0.py
Restart=always

[Install]
WantedBy=multi-user.target

##############################################################

Atualização do deploy:
sudo systemctl stop fletapp.service
sudo systemctl daemon-reload
sudo systemctl enable fletapp.service
sudo systemctl start fletapp.service

teste da aplicação:
sudo systemctl status fletapp.service
'''

# Descobre o diretório onde o script está
BASE_DIR = Path(__file__).parent
TOKEN_FILE = BASE_DIR / "session_token.json"


# Caminho absoluto para a imagem
#image_path = BASE_DIR / "Imagem_Quest.svg"
#gif_path = BASE_DIR / "evolucao.gif"
#image_log_path = BASE_DIR / "Enind Grupo - Vetor.svg"

def caminho_recurso(rel_path):
    """Retorna o caminho absoluto de arquivos mesmo no executável PyInstaller"""
    try:
        base_path = sys._MEIPASS  # quando está empacotado
    except AttributeError:
        base_path = os.path.abspath(".")  # quando está em desenvolvimento
    return os.path.join(base_path, rel_path)


################################################################################################################################################################################
################################################################################## PREPARANDO EMAILS ###########################################################################
################################################################################################################################################################################

import mimetypes

import smtplib
import getpass



from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import random
import datetime

def prepara_corpo_email_Codigo(NumCod):
    agora = datetime.datetime.now()
    fuso_horario =  timezone('America/Sao_Paulo') #Horario de São Paulo
    agora = agora.astimezone(fuso_horario)
    hora_agora = agora.time().hour

    
    corpo = 'Prezado(a),'

    if hora_agora >= 0 and hora_agora <= 12:
        corpo = 'Bom Dia!'
    elif hora_agora >= 13 and hora_agora <= 18:
        corpo = 'Boa Tarde!'
    elif hora_agora > 19:
        corpo = 'Boa Noite'
    

    corpo += f"<br><br>Segue o numero para acessar a página de Questionário da ENIND: <br><br><b>{NumCod}</b><br><br>"
    corpo += f"<b>E-mail automatico, utilizado apenas para envio.</b>"
    corpo += f"<br> Atenciosamente,"

    return corpo


def enviar_email(para, assunto, corpo):
    try:
        #sender = 'NF@enind.com.br'
        #password = 'Enind@2020'

        sender = 'nao-responda@enind.com.br'
        password = 'N102030r!@'
        
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = para
        msg['Subject'] = assunto
        
        CaminhoGIF = "https://enind.com.br/wp-content/uploads/2024/03/Automacao-ENIND-4-1-1.gif"
        CorpoEmail = corpo +  f"<br><img src={chr(34)}{CaminhoGIF}{chr(34)}>" 
        
        # Corpo da mensagem
        msg.attach(MIMEText(CorpoEmail, 'html', 'utf-8'))

        raw = msg.as_string()


        with smtplib.SMTP('smtp-mail.outlook.com', 587) as smtp_server:
            smtp_server.ehlo()  # Pode ser omitido
            smtp_server.starttls()  # Protege a conexao
            smtp_server.ehlo()  # Pode ser omitido
            smtp_server.login(sender, password)
            smtp_server.sendmail(sender, para, raw)
            smtp_server.quit()


        return True, ""
    except Exception as inst:
       return False, inst
        



###########################################################################################################################################      
############################################################### Banco de Dados ############################################################
########################################################################################################################################### 

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

def inserir_banco(tabela, dados, campos):
    try:
        conn = mysql_connection()
        if not conn: return False

        query = f"""
        INSERT INTO {tabela} {campos}
        """

        with conn.cursor() as cursor:
            cursor.execute(query, dados)
            conn.commit()

        conn.close()
        return True
    except Exception as e:
        print(f'❌ Erro ao inserir dados no banco: {e}')
        return False

def registra_login(Pessoa):
    try:
    # Horário Brasil
        fuso = pytz.timezone("America/Sao_Paulo")
        data_login = datetime.datetime.now(fuso)
        data_login_formatado = data_login.strftime("%Y-%m-%d %H:%M:%S")

        conn = mysql_connection()
        if not conn:
            return False

        query = """
            INSERT INTO QuestRH_Acessos (Nome, Data_Acesso, Usuario_Acesso, Maquina_Acesso)
            VALUES (%s, %s, %s, %s)
        """

        cursor = conn.cursor()
        cursor.execute(query, (Pessoa, data_login_formatado, '', ''))
        conn.commit()

        cursor.close()
        conn.close()
        return True
    except Exception as e:
        return False
  
        
def captura_valor_nota(Participante, Avaliador):
    try:
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(f"SELECT ROUND(AVG(Resposta)) as Media FROM QuestRH_Respostas where Participante = '{Participante}' and Nome_Avaliador = '{Avaliador}'")
        resultados = cursor.fetchone()
        conn.close()

        if resultados['Media'] == 0:
            return ''
        else:
            return resultados['Media']
    except Exception as e:
        print("Erro ao conectar ao banco de dados:", e)
        return ''


def obter_perguntas(Pessoa):
    try:
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(f"SELECT Tipo_Avaliacao FROM QuestRH_Relacoes where Participante = '{Pessoa}'")
        resultados = cursor.fetchone()

        tipo_aval = resultados['Tipo_Avaliacao']
        cursor.execute(f"SELECT * FROM QuestRH_Perguntas where Tipo_Avaliacao = '{tipo_aval}'")
        resultados = cursor.fetchall()
        conn.close()

        perguntas = []
        for row in resultados:
            perguntas.append({
                "ID": row['ID'],
                "Pilar": row['Pilar'],
                "Competencia": row['Competencia'],
                "Pergunta": row['Pergunta']
            })

        return perguntas
    except Exception as e:
        print("Erro ao conectar ao banco de dados:", e)
        return []

def  define_status(Participante, Pessoa):
    conn = mysql_connection()
    cursor = conn.cursor(pymysql.cursors.DictCursor)
    scrp_sql = f"SELECT * FROM QuestRH_Respostas Where Participante = '{Participante}' and Nome_Avaliador = '{Pessoa}'"
    cursor.execute(scrp_sql)
    resultados = cursor.fetchone()
    conn.close()
    if resultados:
        return 'Realizado'
    else:
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        scrp_sql = f"SELECT * FROM QuestRH_Relacoes Where Participante = '{Participante}'"
        cursor.execute(scrp_sql)
        resultados = cursor.fetchone()
        conn.close()

        if resultados['Tipo_Avaliacao']=='A3' and Pessoa == Participante:
            return 'Realizado'
        else:
            return 'Pendente'

def valida_texto(texto):
    NaoPermitidos = f"SELECT,DELETE,INSERT,',%,{chr(34)},TRUNCATE,DROP,JOIN"
    palavras = NaoPermitidos.split(",")

    for palavra in palavras:
            if palavra in texto.upper():
                return False
    
    return True


def lista_pendencias(estrategico = True, nao_estrategico = True):
    conn = mysql_connection()
    cursor = conn.cursor(pymysql.cursors.DictCursor)

    sql_condicao = ''

    if estrategico ==  True and nao_estrategico == False:
        sql_condicao = " Where Grupo_estrategico = 'Sim'"
    elif estrategico ==  False and nao_estrategico == True:
        sql_condicao = " Where Grupo_estrategico = 'Não'"

    # Carrega tudo de uma vez
    cursor.execute(f"SELECT * FROM QuestRH_Relacoes {sql_condicao}")
    relacoes = cursor.fetchall()

    cursor.execute(f'SELECT Participante, Nome_Avaliador, ROUND(AVG(Resposta)) as Media FROM QuestRH_Respostas {sql_condicao} GROUP BY Participante, Nome_Avaliador order by Participante')
    respostas_agrupadas = cursor.fetchall()

    conn.close()    

    # Indexa dados em dicionários para acesso rápido
    respostas_dict = {
        (r['Participante'], r['Nome_Avaliador']): r['Media']
        for r in respostas_agrupadas
    }

    tipo_avaliacao_dict = {
        (r['Participante']): r['Tipo_Avaliacao']
        for r in relacoes
    }

    lista = []

    for row in relacoes:
        participante = row['Participante']
        avaliador1 = row['Avaliador1']
        avaliador2 = row['Avaliador2']
        tipo_avaliacao = tipo_avaliacao_dict.get(participante, '')
        if row['data_feedback']:
            data_feedback = row['data_feedback'].strftime('%d/%m/%Y %H:%M:%S')
        else:
            data_feedback = ''

        def obter_status(p, a):
            if (p, a) in respostas_dict:
                return 'Realizado'
            elif tipo_avaliacao == 'A3' and p == a:
                return 'Realizado'
            else:
                return 'Pendente'

        def obter_media(p, a):
            valor = respostas_dict.get((p, a), None)
            return '' if valor == 0 or valor is None else valor

        lista.append({
            "Participante": participante,
            "Status": obter_status(participante, participante),
            "Avaliacao": obter_media(participante, participante),
            "Avaliador1": avaliador1,
            "Status1": obter_status(participante, avaliador1),
            "Avaliacao1": obter_media(participante, avaliador1),
            "Avaliador2": avaliador2,
            "Status2": obter_status(participante, avaliador2),
            "Avaliacao2": obter_media(participante, avaliador2),
            "Questionário": tipo_avaliacao,
            'data_feedback': data_feedback
        })

        lista.sort(key=lambda x: x["Participante"])

    return lista

def define_avaliacao_final(Pessoa):
    scrp_sql = f"""
        SELECT Case When média_av1 = média_av2 then média_av1 else CEIL(( média_av1 + média_av2) / 2) end as Media FROM 
        (Select round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1,
                round(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as média_desemp_av1,
                round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2,
                round(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as média_desemp_av2 from
        QuestRH_Respostas Where Participante = '{Pessoa}' and ID_Rel > 0
        having COUNT(DISTINCT ID_Rel) = 2) as x
    """
    conn = mysql_connection()
    cursor = conn.cursor(pymysql.cursors.DictCursor)
    cursor.execute(scrp_sql)
    resultados = cursor.fetchone()
    conn.close()
    if resultados:
        return resultados['Media']
    else:
        return 0
  
    
def obter_questionarios(Pessoa):
    try:
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        if Pessoa != 'Administrador':
            scrp_sql = f"SELECT * FROM QuestRH_Relacoes Where Participante = '{Pessoa}' or Avaliador1 = '{Pessoa}' or Avaliador2 = '{Pessoa}' order by Participante"
        else:
            scrp_sql = f"SELECT * FROM QuestRH_Relacoes order by Participante"

        #print(scrp_sql)
        lista = []
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()

        # Carrega tudo de uma vez
        cursor.execute("SELECT * FROM QuestRH_Relacoes")
        relacoes = cursor.fetchall()

        cursor.execute("SELECT Participante, Nome_Avaliador, ROUND(AVG(Resposta)) as Media FROM QuestRH_Respostas GROUP BY Participante, Nome_Avaliador")
        respostas_agrupadas = cursor.fetchall()

    
        conn.close()

        # Indexa dados em dicionários para acesso rápido
        respostas_dict = {
            (r['Participante'], r['Nome_Avaliador']): r['Media']
            for r in respostas_agrupadas
        }

        tipo_avaliacao_dict = {
            (r['Participante']): r['Tipo_Avaliacao']
            for r in relacoes
        }


        for row in resultados:
            def obter_status(p, a):
                participante = row['Participante']
                tipo_avaliacao = tipo_avaliacao_dict.get(participante, '')

                if (p, a) in respostas_dict:
                    return 'Realizado'
                elif tipo_avaliacao == 'A3' and p == a:
                    return 'Realizado'
                else:
                    return 'Pendente'

            def obter_media(p, a):
                valor = respostas_dict.get((p, a), None)
                return '' if valor == 0 or valor is None else valor

            if row['Participante'] == Pessoa:
                Avaliador = 0
            elif row['Avaliador1'] == Pessoa:
                Avaliador = 1
            elif row['Avaliador2'] == Pessoa:
                Avaliador = 2
            else:
                Avaliador = ''
            
            #Disponibiliza a visualização da nota apenas para os avaliadores 1 e 2
            if Avaliador == 0:
                if obter_status(row['Participante'],row['Participante']) == 'Realizado':
                    Resp_auto = "Sim"
                else:
                    Resp_auto = "Não"

                if obter_status(row['Participante'],row['Avaliador2']) == 'Realizado':
                    Resp_aval2 = "N/I"
                else:
                    Resp_aval2 = "N/I"

                if obter_status(row['Participante'],row['Avaliador1']) == 'Realizado':
                    Resp_aval1 = "N/I"
                else:
                    Resp_aval1 = "N/I"
            else:
                Resp_auto = obter_media(row['Participante'],row['Participante'])
                Resp_aval1 = obter_media(row['Participante'],row['Avaliador1']) 
                Resp_aval2 = obter_media(row['Participante'],row['Avaliador2']) 

            if row['data_feedback']:
                data_feedback = row['data_feedback'].strftime('%d/%m/%Y %H:%M:%S')
            else:
                data_feedback = ''
            
            if Pessoa == 'Administrador':
                if Resp_auto == '' or Resp_aval1 == '' or Resp_aval2 == '':
                    Status = 'Pendente'
                else:
                    Status = 'Realizado'

            else:
                Status = obter_status(row['Participante'], Pessoa)

            
            
            
            lista.append(
                {"nome": row['Participante'], 
                 "Avaliador": Avaliador, 
                 "Questionário": row['Tipo_Avaliacao'],
                 "status": Status,
                 "auto_aval": Resp_auto,
                 "primaria": Resp_aval1, 
                 "secundaria": Resp_aval2,
                 "data_feedback": data_feedback}
            )

        return lista
    except Exception as e:
        print("Erro ao conectar ao banco de dados:", e)
        return []

##########################################################################################################################################
############################################################ SEGURANÇA DA INFORMAÇÃO #####################################################
##########################################################################################################################################
import string


def gera_token(usuario, size=30, chars=string.ascii_uppercase + string.digits + string.ascii_lowercase,):
    strtoken = ''.join(random.choice(chars) for _ in range(size))
    dados = {"token": strtoken, "usuario": usuario}
    with open(TOKEN_FILE, "w") as f:
        json.dump(dados, f)
    return strtoken


def validar_token(token, pessoa):
    scrp_sql = f"SELECT * FROM QuestRH_Pessoas Where Pessoa = '{pessoa}'"
    conn = mysql_connection()
    cursor = conn.cursor(pymysql.cursors.DictCursor)
    cursor.execute(scrp_sql)
    resultados = cursor.fetchone()
    conn.close()

    if token != resultados['Token']:
        return False
    else:
        return True
    




###########################################################################################################################################      
############################################################ Inicio da Aplicação ##########################################################
########################################################################################################################################### 
IDLE_TIMEOUT = 300  # segundos -> 5*60 - 5 Minutos
        
def main(page: ft.Page):
    
    #print(f'🚀 Iniciando aplicação... a imagem está no diretório:{image_path}')
    codigo_enviado = ""
    nome_logado = ""
    page.title = "Sistema de Avaliação ENIND"
    #page.scroll = ft.ScrollMode.AUTO
    page.window.maximized = True
    data_limite = '2026/06/01 23:59:59'
    #data_limite = '2025/07/21 23:59:59'
    file_picker = ft.FilePicker()
    page.overlay.append(file_picker)
    linhas = []

    page.padding = 0
    page.bgcolor = ft.Colors.WHITE
    page.theme_mode = ft.ThemeMode.LIGHT
    

    
    page.meta_tags = [
        {
            "name": "viewport", 
            "content": "width=device-width, initial-scale=1.0, user-scalable=yes"
        }
    ]

    
    # Configuração de Fonte Moderna
    page.fonts = {
        "Roboto-Medium": "https://github.com/google/fonts/raw/main/apache/roboto/static/Roboto-Medium.ttf",
        "Inter": "https://github.com/google/fonts/raw/main/apache/roboto/static/Roboto-Light.ttf"
    }
    page.theme = ft.Theme(font_family="Inter")
    page.update()
    
    chk_estrategico = ft.Checkbox(
        label="Estratégico",
        value=True,
        on_change=lambda e: atualiza_rel(e),
    )

    chk_nao_estrategico = ft.Checkbox(
        label="Não Estrategico",
        value=True,
        on_change=lambda e: atualiza_rel(e),
    )
    # Container de alerta
    alerta_container = ft.Container(
        content=ft.Text("", color=ft.Colors.WHITE),
        bgcolor=ft.Colors.RED_400,
        height=40,
        padding=10,
        alignment=ft.alignment.center,
        visible=False
    )
    
    def fechar_alerta(e= None):
        alerta_container_form.visible = False
        page.update()

    btn_fechar_alerta = ft.TextButton(on_click=lambda _:fechar_alerta, icon=ft.Icons.CLOSE)
    txt_msg_alerta = ft.Text("", color=ft.Colors.WHITE, expand=True) 

    # Container de alerta
    alerta_container_form = ft.Container(
        content=ft.Row([txt_msg_alerta,btn_fechar_alerta], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        bgcolor=ft.Colors.RED_400,
        height=40,
        padding=10,
        alignment=ft.alignment.center,
        visible=False
    )

    # Função para exibir alerta temporariamente
    def mostrar_alerta_temporario(mensagem, coloração, segundos=3):
        def tarefa():
            alerta_container.content.value = mensagem
            alerta_container.visible = True
            alerta_container.bgcolor = coloração
            page.update()
            time.sleep(segundos)
            alerta_container.visible = False
            page.update()
        threading.Thread(target=tarefa).start()
    
    #Criando Objetos

    senha_txt =  ft.TextField(label="Senha", width=380, border_radius=12, password=True, can_reveal_password=True, prefix_icon=ft.Icons.LOCK)
    expiracao_txt = ft.Text(f'Data para envio de formulário expirado no dia {data_limite}.',
                            expand= True, 
                            visible= True, 
                            size=15,
                            weight=ft.FontWeight.BOLD, 
                            color= ft.Colors.RED_700,
                            text_align=ft.alignment.center)
    
    container_expiração = ft.Container(expiracao_txt, 
                                       alignment=ft.alignment.center, 
                                       visible=False, 
                                       bgcolor=ft.Colors.RED_100, 
                                       expand=True, 
                                       border_radius=10, 
                                       padding=10, 
                                       height=50)

    mensagem_aguarde = ft.Text(
                    "Aguarde, atualizando o relatório...",
                    size=18,
                    weight=ft.FontWeight.BOLD,
                    text_align=ft.TextAlign.CENTER,
                    color=ft.Colors.WHITE
                )


    erro_login = ft.Text("", color=ft.Colors.RED)
    txt_observacoes = ft.TextField(label="", expand=True, multiline=True,max_length=200000, 
                                    on_change=lambda e: reset_idle_time(e), 
                                    on_click=lambda e: reset_idle_time(e),
                                    on_blur=lambda e: reset_idle_time(e),
                                    on_focus=lambda e: reset_idle_time(e),
                                    on_submit=lambda e: reset_idle_time(e),
                                    on_animation_end=lambda e: reset_idle_time(e),
                                    visible=True,
                                    border=ft.InputBorder.NONE,
                                    filled=False,
                                    shift_enter=True,)
    

    
    txt_observacoes_e = ft.TextField(label="Pilar E*", expand=True, multiline=False,max_length=2000, 
                                    on_change=lambda e: reset_idle_time(e), 
                                    on_click=lambda e: reset_idle_time(e),
                                    on_blur=lambda e: reset_idle_time(e),
                                    on_focus=lambda e: reset_idle_time(e),
                                    on_submit=lambda e: reset_idle_time(e),
                                    on_animation_end=lambda e: reset_idle_time(e),
                                    visible=True,
                                    filled=False,
                                    shift_enter=True,)

    txt_observacoes_p = ft.TextField(label="Pilar P*", expand=True, multiline=False,max_length=2000, 
                                    on_change=lambda e: reset_idle_time(e), 
                                    on_click=lambda e: reset_idle_time(e),
                                    on_blur=lambda e: reset_idle_time(e),
                                    on_focus=lambda e: reset_idle_time(e),
                                    on_submit=lambda e: reset_idle_time(e),
                                    on_animation_end=lambda e: reset_idle_time(e),
                                    visible=True,
                                    filled=False,
                                    shift_enter=True,)
    
    txt_observacoes_c = ft.TextField(label="Pilar C*", expand=True, multiline=False,max_length=2000, 
                                    on_change=lambda e: reset_idle_time(e), 
                                    on_click=lambda e: reset_idle_time(e),
                                    on_blur=lambda e: reset_idle_time(e),
                                    on_focus=lambda e: reset_idle_time(e),
                                    on_submit=lambda e: reset_idle_time(e),
                                    on_animation_end=lambda e: reset_idle_time(e),
                                    visible=True,
                                    filled=False,
                                    shift_enter=True,)
    
    txt_calibracao = ft.Text("Calibração",weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900)
    divisor_observacao = ft.Divider(height=1, thickness=1, color=ft.Colors.BLUE_800, visible=True)
    txt_campos_obrigatórios = ft.Text("campos com * são obrigatórios o preenchimento",size=10,weight=ft.FontWeight.BOLD,color=ft.Colors.GREY_700)
    txt_titulo_obs = ft.Text("Observações*",weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900)
    linha_calibracao = ft.Column([txt_calibracao, 
                        txt_observacoes_e, 
                        txt_observacoes_p, 
                        txt_observacoes_c,
                        
                        divisor_observacao], spacing=10, visible = False)
    
    container_observacoes = ft.Container(
        content=ft.Column([linha_calibracao,
                            txt_titulo_obs,
                            txt_observacoes,
                            txt_campos_obrigatórios]),
        expand=1,
        bgcolor=ft.Colors.GREY_50,   
        height = page.height * 0.7,
        border_radius=10,
        padding=10,
        border=ft.border.all(1, ft.Colors.BLUE_800),
        shadow=ft.BoxShadow(
                blur_radius=10,
                color=ft.Colors.with_opacity(0.05, ft.Colors.BLACK),
                offset=ft.Offset(0, 4)
            ),
    )
    
    
    
    nome_cb = ft.TextField(label="Login", width=380, border_radius=12, prefix_icon=ft.Icons.PERSON)
    form_inputs = []

    txt_observacoes_auto = ft.Text("", size=15)

    
    container_obs_auto = ft.Container(
                            content=ft.Column([                
                                ft.Text("Observações:", size=20, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                txt_observacoes_auto], scroll=ft.ScrollMode.ADAPTIVE,),
                            bgcolor=ft.Colors.GREY_50,   
                            height= page.height * 0.7,  
                            padding=15, 
                            border_radius=10,
                            expand=True
                            )

    
    txt_observacoes_av1 =ft.Text("", size=15)
    txt_observacoes_av1_e = ft.Text("", size=15)
    txt_observacoes_av1_p = ft.Text("", size=15)
    txt_observacoes_av1_c = ft.Text("", size=15)
    container_obs_av1 = ft.Container(
                            content=ft.Column([ 
                                ft.Text("Calibração", size=20, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                ft.Row([ft.Text("Pilar E:", size=9, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                txt_observacoes_av1_e]), 
                                ft.Row([ft.Text("Pilar P:", size=9, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                txt_observacoes_av1_p]), 
                                ft.Row([ft.Text("Pilar C:", size=9, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                txt_observacoes_av1_c]),   
                                ft.Divider(height=1, thickness=1, color=ft.Colors.BLUE_800),                
                                ft.Text("Observações:", size=20, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                txt_observacoes_av1], scroll=ft.ScrollMode.ADAPTIVE,),
                            bgcolor=ft.Colors.GREY_50,
                            height= page.height * 0.7,       
                            padding=15, 
                            border_radius=10,
                            expand=True
                            )

    
    txt_observacoes_av2 = ft.Text("", size=15, no_wrap=False, expand=True)
    txt_observacoes_av2_e = ft.Text("", size=15)
    txt_observacoes_av2_p = ft.Text("", size=15)
    txt_observacoes_av2_c = ft.Text("", size=15)
    container_obs_av2 = ft.Container(
                            content=ft.Column([ 
                                ft.Text("Calibração", size=20, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                ft.Text("E", size=15, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                txt_observacoes_av2_e, 
                                ft.Text("P", size=15, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                txt_observacoes_av2_p, 
                                ft.Text("C", size=15, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                txt_observacoes_av2_c,   
                                ft.Divider(height=1, thickness=1, color=ft.Colors.BLUE_800),                 
                                ft.Text("Observações:", size=20, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900),
                                txt_observacoes_av2], scroll=ft.ScrollMode.ADAPTIVE,), 
                            bgcolor=ft.Colors.GREY_50,
                            height= page.height * 0.7,       
                            padding=15, 
                            border_radius=10,
                            expand=True
                            )
    

    txt_desempenho_auto = ft.Text("", expand=True, size=15)
    container_desemp_auto = ft.Container(
                content=txt_desempenho_auto,
                border_radius=10,
                padding=10,
            )
    container_desempenho_auto = ft.Container(
                            content=ft.Column([
                                ft.Text('Desempenho - Execução e entrega de atividades de forma geral:', size=15, weight=ft.FontWeight.BOLD),
                                container_desemp_auto],spacing=5), 
                            bgcolor=ft.Colors.GREY_50,      
                            padding=15, 
                            border_radius=10
                            )

    
    txt_desempenho_av1 =ft.Text("", expand=True, size=15)

    container_desemp_av1 = ft.Container(
                content=txt_desempenho_av1, 
                border_radius=10,
                padding=10,
            )
    container_desempenho_av1 = ft.Container(
                            content=ft.Column([
                                ft.Text('Desempenho - Execução e entrega de atividades de forma geral', size=15, weight=ft.FontWeight.BOLD),
                                container_desemp_av1],spacing=5), 
                            bgcolor=ft.Colors.GREY_50,      
                            padding=15, 
                            border_radius=10
                            )

    
    txt_desempenho_av2 = ft.Text("", expand=True, size=15)
    container_desemp_av2 = ft.Container(
                content=txt_desempenho_av2, 
                border_radius=10,
                padding=10,
            )
    container_desempenho_av2 = ft.Container(
                            content=ft.Column([
                                ft.Text('Desempenho - Execução e entrega de atividades de forma geral', size=15, weight=ft.FontWeight.BOLD),
                                container_desemp_av2],spacing=5), 
                            bgcolor=ft.Colors.GREY_50,      
                            padding=15, 
                            border_radius=10
                            )


    # Elemento de saudação
    texto_ola = ft.Text("", size=15, weight=ft.FontWeight.BOLD)
    texto_ola1 = ft.Text("", size=15, weight=ft.FontWeight.BOLD)
    texto_ola2 = ft.Text("", size=15, weight=ft.FontWeight.BOLD)
    texto_ola3 = ft.Text("", size=15, weight=ft.FontWeight.BOLD)
    texto_ola4 = ft.Text("", size=15, weight=ft.FontWeight.BOLD)
    texto_ola5 = ft.Text("", size=15, weight=ft.FontWeight.BOLD)
    texto_ola6 = ft.Text("", size=15, weight=ft.FontWeight.BOLD)

    lista_view = ft.ListView(expand=True, auto_scroll=False,  spacing=10, padding=10)
    lista_pend_view = ft.ListView(expand=True, auto_scroll=False, spacing=10, padding=10,visible=False)

    lista_reultados_auto_view = ft.ListView(expand=True, auto_scroll=False, spacing=10, padding=10)
    lista_reultados_av1_view = ft.ListView(expand=True, auto_scroll=False, spacing=10, padding=10)
    lista_reultados_av2_view = ft.ListView(expand=True, auto_scroll=False, spacing=10, padding=10)

    txt_auto = ft.Text("", size=15, color=ft.Colors.BLUE_900, weight=ft.FontWeight.BOLD)
    txt_av1 = ft.Text("", size=15, color=ft.Colors.BLUE_900, weight=ft.FontWeight.BOLD)
    txt_av2 = ft.Text("", size=15, color=ft.Colors.BLUE_900, weight=ft.FontWeight.BOLD)
    
    txt_status_auto = ft.Text("Não Realizado", size=15, expand=True, text_align='center')
    txt_status_av1 = ft.Text("Não Realizado", size=15, expand=True, text_align='center')
    txt_status_av2 = ft.Text("Não Realizado", size=15, expand=True, text_align='center')
    
    container_status_auto = ft.Container(
        content=txt_status_auto,
        padding=10,
        bgcolor=ft.Colors.WHITE,
        alignment=ft.alignment.center,
        border_radius=10,
        expand=1
    )
    container_status_av1 = ft.Container(
        content=txt_status_av1,
        padding=10,
        bgcolor=ft.Colors.WHITE,
        alignment=ft.alignment.center,
        border_radius=10,
        expand=1
    )
    container_status_av2 = ft.Container(
        content=txt_status_av2,
        alignment=ft.alignment.center,
        padding=10,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        expand=1
    )
    
    
    leading_auto = ft.CircleAvatar(
                            foreground_image_src='', #imagem rosto,
                            radius=22, # Tamanho razoável
                            content=ft.Text('ABC')) # Letra inicial se a imagem falhar
                        
    
    leading_av1 = ft.CircleAvatar(
                        foreground_image_src='', #imagem rosto,
                        radius=22, # Tamanho razoável
                        content=ft.Text('ABC'),
                        ) # Letra inicial se a imagem falhar
                    

    leading_av2 = ft.CircleAvatar(
                    foreground_image_src='', #imagem rosto,
                    radius=22, # Tamanho razoável
                    content=ft.Text('ABC')) # Letra inicial se a imagem falhar
                
 
    list_tile_auto = ft.ListTile(leading=leading_auto, bgcolor=ft.Colors.WHITE)
    container_txt_auto = ft.Container(
                content=ft.Row([
                    ft.Container(list_tile_auto, expand=1),
                    ft.Container(ft.Text('Auto Avaliação', size = 12, color=ft.Colors.BLUE_900), expand=1),
                    ft.Container(txt_auto, expand=4),
                    container_status_auto]),
                border_radius=12,
                bgcolor= ft.Colors.WHITE,
                margin=ft.margin.only(bottom=10),
                animate_opacity=300,
                shadow=ft.BoxShadow(
                    blur_radius=10,
                    color=ft.Colors.with_opacity(0.05, ft.Colors.BLACK),
                    offset=ft.Offset(0, 4)
                ),
                on_click=lambda _: mostrar_formulario(av='auto')
            )
    
    list_tile_av1 = ft.ListTile(leading=leading_av1, bgcolor=ft.Colors.WHITE)
    container_txt_av1 = ft.Container(
                content=ft.Row([
                    ft.Container(list_tile_av1, expand=1),
                    ft.Container(ft.Text('Avaliador 1', size = 12, color=ft.Colors.BLUE_900), expand=1),
                    ft.Container(txt_av1, expand=4),
                    container_status_av1]),
                border_radius=12,
                bgcolor= ft.Colors.WHITE,
                margin=ft.margin.only(bottom=10),
                animate_opacity=300,
                shadow=ft.BoxShadow(
                    blur_radius=10,
                    color=ft.Colors.with_opacity(0.05, ft.Colors.BLACK),
                    offset=ft.Offset(0, 4)
                ),
                on_click=lambda _: mostrar_formulario(av='av1')
    )
    
    list_tile_av2 = ft.ListTile(leading=leading_av2, bgcolor=ft.Colors.WHITE)
    container_txt_av2 = ft.Container(
                content=ft.Row([
                    
                    ft.Container(list_tile_av2, expand=1),
                    ft.Container(ft.Text('Avaliador 2', size = 12, color=ft.Colors.BLUE_900), expand=1),
                    ft.Container(txt_av2, expand=4),
                    container_status_av2]),
                border_radius=12,
                bgcolor= ft.Colors.WHITE,
                margin=ft.margin.only(bottom=10),
                animate_opacity=300,
                shadow=ft.BoxShadow(
                    blur_radius=10,
                    color=ft.Colors.with_opacity(0.05, ft.Colors.BLACK),
                    offset=ft.Offset(0, 4)
                ),
                on_click=lambda _: mostrar_formulario(av='av2')
            )
    
    def mostrar_formulario(av=''):
        if av == '': return
        if av == 'auto':
            if container_respostas_auto.visible == True:
                container_respostas_auto.visible = False
            else:
                container_respostas_auto.visible = True
            container_respostas_av1.visible = False
            container_respostas_av2.visible = False
            
        elif av == 'av1':
            if container_respostas_av1.visible == True:
                container_respostas_av1.visible = False
            else:
                container_respostas_av1.visible = True
            
            container_respostas_auto.visible = False
            container_respostas_av2.visible = False

        elif av == 'av2':
            
            if container_respostas_av2.visible == True:
                container_respostas_av2.visible = False
            else:
                container_respostas_av2.visible = True
                
            container_respostas_auto.visible = False
            container_respostas_av1.visible = False
            
        page.update()
        return
            
    btn_ver_manual= ft.ElevatedButton(
            "Manual de apoio",
            icon= ft.Icons.LIBRARY_BOOKS,
            style=ft.ButtonStyle(
                    bgcolor=ft.Colors.BLUE,
                    color=ft.Colors.WHITE,
                ),
             on_click=lambda e: abrir_manual(e)
    )
    

    msg_deslog = ft.Text(
                size=18,
                weight=ft.FontWeight.BOLD,
                text_align=ft.TextAlign.CENTER,
                color=ft.Colors.WHITE
            )
   

    dropdown_desempenho = ft.Dropdown(
                    options='',
                    hint_text="Avaliação",
                    expand= True
                )
    
    dropdown_Performance = ft.Dropdown(
                    label='Performance',
                    options=[
                        ft.dropdown.Option("Baixa Performance"),
                        ft.dropdown.Option("Inconsistente"),
                        ft.dropdown.Option("Especialista"),
                        ft.dropdown.Option("Dilema"),
                        ft.dropdown.Option("Competente"),
                        ft.dropdown.Option("Forte Entrega"),
                        ft.dropdown.Option("Desafio"),
                        ft.dropdown.Option("Forte Cultura"),
                        ft.dropdown.Option("Alto Potencial"),
                    ],
                    hint_text="Avaliação",
                    expand= True
                )

    txt_media_final = ft.Text("", size=32, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900, text_align=ft.TextAlign.CENTER)

    container_media_final = ft.Container(
                            content=ft.Column([
                                ft.Text(
                                    "MÉDIA FINAL", 
                                    size=11, 
                                    weight=ft.FontWeight.W_500, 
                                    color=ft.Colors.BLUE_GREY_700
                                ),
                                txt_media_final],
                                alignment=ft.MainAxisAlignment.CENTER, # Centraliza verticalmente na coluna
                                horizontal_alignment=ft.CrossAxisAlignment.CENTER, # Centraliza horizontalmente na coluna
                                spacing=0 # Diminui o espaço entre o rótulo e o número
                            ), 
                            bgcolor=ft.Colors.WHITE, # Fundo branco com borda colorida fica mais elegante
                            padding=10,
                            width=110, # Um pouco mais largo para não apertar o texto
                            height=110,
                            border_radius=15,
                            border=ft.border.all(2, ft.Colors.BLUE_100), # Borda suave em vez de fundo sólido forte
                            shadow=ft.BoxShadow(
                                blur_radius=15,
                                color=ft.Colors.with_opacity(0.1, ft.Colors.BLACK),
                                offset=ft.Offset(0, 5)
                            ),
                            alignment=ft.alignment.center, # Alinhamento do conteúdo dentro do container
                            visible=True
                        )

    participante_realizado = ft.Text("", size=15, weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER, color=ft.Colors.GREY_800)
    

    cboxPessoa = ft.Dropdown(label="Pessoa", expand=3)
    cboxPilar = ft.Dropdown(label="Pilar", expand=2)
    cboxCompetencia = ft.Dropdown(label="Competência", expand=2)

    btn_limpar_campos = ft.ElevatedButton("Limpar", on_click=lambda e: limpar_campos(e), width=100 )
  
    img_Ninebox = ft.Image(src_base64= '', fit=ft.ImageFit.CONTAIN, expand=True)
    img_Pilar= ft.Image(src_base64= '', fit=ft.ImageFit.CONTAIN, expand=True)
    img_Comp= ft.Image(src_base64= '', fit=ft.ImageFit.CONTAIN, expand=True)
    img_Compar= ft.Image(src_base64= '', fit=ft.ImageFit.CONTAIN, expand=True)
    img_BellCurve = ft.Image(src_base64= '', fit=ft.ImageFit.CONTAIN, expand=True)

     # Containers com ListView rolável
    dropdown_empresa = ft.Column([ft.ListView(expand=True, spacing=5, auto_scroll=False)])
    dropdown_c_custo = ft.Column([ft.ListView(expand=True, spacing=5, auto_scroll=False)])
    dropdown_cargo = ft.Column([ft.ListView(expand=True, spacing=5, auto_scroll=False)])

    # Botões para abrir os dialogs
    botoes = ft.Row([
        ft.ElevatedButton("Selecionar Empresas", on_click=lambda e: abrir_dialogo("Empresas", dropdown_empresa)),
        ft.ElevatedButton("Selecionar Centro de Custo", on_click=lambda e: abrir_dialogo("C_Custo", dropdown_c_custo)),
        ft.ElevatedButton("Selecionar Cargos", on_click=lambda e: abrir_dialogo("Cargos", dropdown_cargo)),
    ])

    txt_outro_query = ft.Text("", size=15, weight=ft.FontWeight.BOLD,visible=False)


     # imagens
    img_Potencial = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_Potencial_nao_finalizados = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_Concluidos = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_Finalizados = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_Barra_Empresa = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_feedback_antes = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    img_feedback_atual = ft.Image(fit=ft.ImageFit.COVER, expand=True,gapless_playback=True)
    
    # textos -> Referencia de tamanhos de fontes =https://flet.dev/docs/controls/text/#font_family
    txt_Realizados = ft.Text('', size=40, weight='w200',text_align="center", color=ft.Colors.GREEN)
    txt_Finalizados = ft.Text('', size=40, weight='w200',text_align="center",color=ft.Colors.BLUE)
    txt_Pendentes = ft.Text('', size=40, weight='w200',text_align="center",color=ft.Colors.RED)
    txt_Pendentes_participantes = ft.Text('', size=40, weight='w200',text_align="center",color=ft.Colors.AMBER)

   

    ####################################################################################################################################### 
    ######################################################## Funções dos Objetos ##########################################################
    #######################################################################################################################################
    
    #-------------------------------------------------------------------------------------------------------------------------------------------
    #---------------------------------------------------------------- Gráficos -----------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------------------------------------------------
    def atualiza_tabela_bellcurve(e=None):
        Media = dropdown_bellcurve_nota.value
        
        sql_condicao = 'Where ID_REL > 0 '
        sql_codicao_media = ''

        
        if chk_estrategico_grafico.value == True and chk_nao_estrategico_grafico.value == False:
            sql_condicao += ' and Grupo_estrategico = "Sim"'
        elif chk_estrategico_grafico.value == False and chk_nao_estrategico_grafico.value == True:
            sql_condicao += ' and Grupo_estrategico = "Não"'

        if Media:
            sql_codicao_media = f'Where x.Media = {float(Media)}'

        try:
            query_sql = f"""
                Select Participante, Media from (Select Participante,
                    CASE WHEN média_av1 = média_av2 THEN média_av1 ELSE CEIL( (média_av1 + média_av2) / 2) END AS Media
                from 
                (SELECT Participante, id_rel, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
                    round(AVG(CASE WHEN id_rel = 1 THEN Resposta END), 0) as média_av1,
                    round(AVG(CASE WHEN id_rel = 1 THEN Desempenho_Tecnico END), 0) as média_desemp_av1,
                    round(AVG(CASE WHEN id_rel = 2 THEN Resposta END), 0) as média_av2,
                    round(AVG(CASE WHEN id_rel = 2 THEN Desempenho_Tecnico END), 0) as média_desemp_av2,
                    round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_auto,
                    round(AVG(CASE WHEN id_rel = 0 THEN Resposta END), 0) as média_desemp_auto
                FROM QuestRH_Respostas {sql_condicao}
                group by Participante
                HAVING COUNT(DISTINCT id_rel) = 2) as t) as x {sql_codicao_media} order by Participante
            """
        

            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            cursor.execute(query_sql)
            pessoas = cursor.fetchall()
            conn.close()

            lista_grafico_bellcurve.controls.clear()

            for pessoa in pessoas:
                lista_grafico_bellcurve.controls.append(
                    ft.Row([
                        ft.Container(ft.Text(pessoa['Participante'], size=15, weight='w200'),expand=3),
                        ft.Container(ft.Text(str(pessoa['Media']), size=15, weight='w200'), expand=1)
                    ])
                )
                
            lista_grafico_bellcurve.update()
            page.update()
            return
        except Exception as e:
            print(f'Erro ao atualizar tabela: {e}')
            return
    
    def limpar_filtros_bellcurve(e=None):
        dropdown_bellcurve_nota.value = None
        atualiza_tabela_bellcurve()
        return

    def atualizar_painel_status(e=None):
        aguarde_overlay.visible = True
        mensagem_aguarde = 'Aguarde, realizando montagem dos status...'
        page.update()
        grafico64_potencial, erro = gera_graficos.gera_gráfico_potencial(finalizados=True)
        grafico64_potencial_nao_finalizados, erro = gera_graficos.gera_gráfico_potencial()
        grafico64_concluidos, erro, realizado, nao_concluidos = gera_graficos.gera_grafico_conclusao()
        grafico64_finalizados, erro, finalizados, nao_finalizados = gera_graficos.gera_grafico_conclusao(finalizados=True)
        grafico64_empresas, erro = gera_graficos.gera_grafico_empresas()
        grafico64_feedback_antes, erro = gera_graficos.gera_grafico_conclusao_feedback(anterior=True)
        grafico64_feedback_atual, erro = gera_graficos.gera_grafico_conclusao_feedback()
        
        
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
        if grafico64_feedback_antes:
            img_feedback_antes.src_base64 = grafico64_feedback_antes
        if grafico64_feedback_atual:
            img_feedback_atual.src_base64 = grafico64_feedback_atual

        txt_Finalizados.value = str(finalizados or 0)
        txt_Realizados.value = str(realizado or 0)
        txt_Pendentes.value = str(nao_concluidos or 0)
        txt_Pendentes_participantes.value = str(nao_finalizados or 0)


        container_painel_grafico.visible = True
        painel_pend_view.visible = False
        imagem_fundo.visible = False

        aguarde_overlay.visible = False
        page.update()
    



    def alimenta_pilares():
        scpr_pilar = "Select Distinct(Pilar) from QuestRH_Perguntas"
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scpr_pilar)
        resultados = cursor.fetchall()

        cboxPilar.options.clear()

        for pilar in resultados:
            cboxPilar.options.append(ft.dropdown.Option(pilar['Pilar']))
        conn.close()

    def alimenta_competencias(Pilar = ""):

        if Pilar == "": 
            scpr_pilar = "Select Distinct(Competencia) from QuestRH_Perguntas"
        else:
            scpr_pilar = "Select Distinct(Competencia) from QuestRH_Perguntas where Pilar = '" + Pilar + "'"

        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scpr_pilar)
        resultados = cursor.fetchall() 

        cboxCompetencia.options.clear()

        for competencia in resultados:
            cboxCompetencia.options.append(ft.dropdown.Option(competencia['Competencia']))

        conn.close()

    def alimenta_pessoas():
        scpr_pilar = "Select Distinct(Participante) from QuestRH_Respostas where id_rel > 0 group by Participante HAVING COUNT(DISTINCT id_rel) = 2"
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scpr_pilar)
        resultados = cursor.fetchall() 

        cboxPessoa.options.clear()
        
        for competencia in resultados:
            cboxPessoa.options.append(ft.dropdown.Option(competencia['Participante']))

        conn.close()
    
    def limpar_campos(e=None):
        # 1. Resetando ComboBoxes e Dropdowns simples
        cboxPessoa.value = None
        cboxPilar.value = None
        cboxCompetencia.value = None
        dropdown_Performance.value = None
        
        # Dica: Se o componente tiver a propriedade 'placeholder', você pode reforçá-la aqui
        cboxPessoa.placeholder = "Pessoa"
        cboxPilar.placeholder = "Pilar"
        cboxCompetencia.placeholder = "Competência"
        dropdown_Performance.placeholder = "Performance"

        # 2. Resetando Controles de Múltipla Escolha (Checkboxes/Custom)
        # Certifique-se de dar update no controle pai após o loop
        for chk in dropdown_c_custo.controls[0].controls:
            chk.selected = False
            if hasattr(chk, "value"): chk.value = False # Garante se for um Checkbox real

        for chk in dropdown_empresa.controls[0].controls: 
            chk.selected = False
            if hasattr(chk, "value"): chk.value = False

        for chk in dropdown_cargo.controls[0].controls:
            chk.selected = False
            if hasattr(chk, "value"): chk.value = False

        # 3. Campo de texto livre
        txt_outro_query.value = ""

        # 4. Atualização visual em lote
        # Em vez de vários .update(), atualizar a página ou os containers principais é mais eficiente
        cboxPessoa.update()
        cboxPilar.update()
        cboxCompetencia.update()
        dropdown_Performance.update()
        
        # MUITO IMPORTANTE: Atualizar os dropdowns de lista para refletir que os itens não estão mais "selected
        #dropdown_empresa.update()
        #dropdown_cargo.update()
        
        # 5. Sincronização de dados
        # Se essas funções dependem dos valores acima, elas agora lerão tudo vazio/None
        alimenta_competencias()
        alimenta_pessoas()
        
        page.update() # Garante a atualização de toda a árvore de componentes
        
        # Por fim, dispara a atualização dos gráficos com os filtros limpos
        atualiza_dados()
    
    def inicializa_grafico():
        alimenta_pessoas()
        alimenta_pilares()
        alimenta_competencias()
        alimenta_chkbox()
        atualiza_dados()
        atualiza_tabela_bellcurve()

        container_grafico.visible = True
        painel_pend_view.visible = False
        page.update()


    def atualiza_dados(e=None):
        mensagem_aguarde.value = 'Aguarde, realizando montagem dos gráficos...'
        imagem_fundo.visible = False
        aguarde_overlay.visible = True
        page.update()

        participante = cboxPessoa.value
        alimenta_pessoas()

        if participante != None:
            participante = cboxPessoa.value
        else:
            participante = ''
        if cboxPilar.value != None :
            pilar = cboxPilar.value
            alimenta_competencias(pilar)
        else:
            pilar = ''

        if cboxCompetencia.value != None :
            competencia = cboxCompetencia.value
        else:
            competencia = ''

        if dropdown_Performance.value != None:
            performance = dropdown_Performance.value
        else:
            performance = ''

        grafico64, msgerro = gera_graficos.gera_ninebox(pilar=pilar, competencia=competencia, participante=participante,outras_condicoes=txt_outro_query.value, performance=performance)
        if grafico64 == None:
            img_Ninebox.visible = False
        else:
            img_Ninebox.visible = True


        grafico_pilar64, msgerro = gera_graficos.gera_gráfico_pilar(pilar=pilar, competencia=competencia, participante=participante,outras_condicoes=txt_outro_query.value)
        if grafico_pilar64 == None:
            img_Pilar.visible = False
        else:
            img_Pilar.visible = True    


        grafico_comp64, msgerro = gera_graficos.gera_gráfico_Competencia(pilar=pilar, competencia=competencia, participante=participante,outras_condicoes=txt_outro_query.value)
        if grafico_comp64 == None:
            img_Comp.visible = False
        else:
            img_Comp.visible = True

        grafico_compar64, msgerro_compar = gera_graficos.gera_gráfico_Comparativo(pilar=pilar, competencia=competencia, participante=participante,outras_condicoes=txt_outro_query.value)
        if grafico_compar64 == None:
            img_Compar.visible = False
        else:  
            img_Compar.visible = True
        
        grafico_bell64, msgerro_compar = gera_graficos.gera_bell_curve(chk_estrategico_grafico.value, chk_nao_estrategico_grafico.value)
        if grafico_bell64 == None:
            img_BellCurve.visible = False
        else:  
            img_BellCurve.visible = True 

        img_Ninebox.src_base64 = grafico64
        img_Pilar.src_base64 = grafico_pilar64
        img_Comp.src_base64 = grafico_comp64
        img_Compar.src_base64 = grafico_compar64
        img_BellCurve.src_base64 = grafico_bell64

        atualiza_tabela_bellcurve()

        alimenta_tabela_graficos(Participante=participante, outras_condicoes=txt_outro_query.value, performance=performance)
        
        aguarde_overlay.visible = False
        page.update()
    
    def exportar_grafico_excel(e=None):

        sql_condicao = ''

        if chk_estrategico_grafico.value == True and chk_nao_estrategico_grafico.value == False:
            sql_condicao = ' and Grupo_estrategico = "Sim"'
        elif chk_estrategico_grafico.value == False and chk_nao_estrategico_grafico.value == True:
            sql_condicao = ' and Grupo_estrategico = "Não"'


        try:
            # Conexão e consulta
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            scrp_sql =f"""
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
                END as Status_Av
                FROM QuestRH_Respostas where id_rel in (0, 1,2) {sql_condicao}
                group by Participante) as t
            ) as x
            """
            cursor.execute(scrp_sql)
            consulta = cursor.fetchall()
            cursor.close()
            conn.close()

            # Criar DataFrame
            df = pd.DataFrame(consulta)

            # Criar arquivo em memória
            output = io.BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            # Nome do arquivo
            nome_arquivo = f"exportacao_grafico_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
             
            # Codificar para base64
            b64 = base64.b64encode(output.read()).decode()

            # Criar link de download
            link_download = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"

            # Abrir o link no navegador (força o download)
            page.launch_url(link_download, web_window_name=nome_arquivo)

            mostrar_alerta_temporario('Exportação realizada com sucesso', ft.Colors.GREEN_400)
        except Exception as ex:
            mostrar_alerta_temporario(f'Erro ao exportar: {ex}', ft.Colors.RED_400)
            

    def exportar_dados_painel_adm(e=None):
        nonlocal lista_dados
        nonlocal lista_pend
        try:
            df = pd.DataFrame(lista_dados)
            df_pend = pd.DataFrame(lista_pend)
            
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            cursor.execute(f"SELECT * FROM QuestRH_Feedbacks where id_ciclo = (SELECT MAX(id_ciclo) FROM QuestRH_Feedbacks);")
            consulta = cursor.fetchall()
            cursor.close()
            conn.close()

            df_feedback_antes = pd.DataFrame(consulta)
            

            # 1. Criar o objeto BytesIO para salvar na memória
            output = io.BytesIO()

            # 2. Usar o ExcelWriter com o engine 'openpyxl'
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Grava a primeira aba
                df.to_excel(writer, sheet_name='Dados Gerais', index=False)
                
                # Grava a segunda aba
                df_pend.to_excel(writer, sheet_name='Pendentes', index=False)
                
                #Grava a Terceira aba
                df_feedback_antes.to_excel(writer, sheet_name= 'Feedback Anterior', index=False)  

            # 3. Importante: O arquivo é finalizado ao sair do bloco 'with'. 
            # Agora voltamos o ponteiro para o início para leitura.
            output.seek(0)

            # 4. Processo de conversão para download no Flet
            nome_arquivo = f"exportacao_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            b64 = base64.b64encode(output.read()).decode()
            link_download = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"

            # 5. Lançar o download
            page.launch_url(link_download, web_window_name=nome_arquivo)
    

            mostrar_alerta_temporario('Exportação realizada com sucesso', ft.Colors.GREEN_400)
        except Exception as ex:
            mostrar_alerta_temporario(f'Erro ao exportar: {ex}', ft.Colors.RED_400)

    def alimenta_tabela_graficos(Participante = '', outras_condicoes = '', performance = ''):
        mensagem_aguarde.value = 'Aguarde, realizando montagem da tabela de dados...'
        aguarde_overlay.visible = True
        page.update()
        
        lista_grafico.controls.clear()

        if outras_condicoes != '' :
            sql_condicao = f" where {outras_condicoes}"
            if Participante != '':
                sql_condicao += f" and Participante = '{Participante}'"
        elif Participante != '':
            sql_condicao = f" where Participante = '{Participante}'"
        else:
            sql_condicao = ""


        scrp_sql =f"""
            SELECT Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo,  Media_Auto, Media_Avaliadores, Media_Aval_Desemp, Media_Auto_Desemp, Performance, Status_Av from 
            (SELECT Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
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
            FROM QuestRH_Respostas {sql_condicao}
            group by Participante) as t
        ) as x 
        ) as y where Status_Av = 'Finalizado'
        """

        if performance != '':
            scrp_sql += f" and Performance = '{performance}'"
            
        scrp_sql += f" order by Participante"

        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall() 
        conn.close

        for r in resultados:
            if r["Status_Av"] == 'Pendente':
                bg_cor = ft.Colors.RED
            else:
                bg_cor = ft.Colors.GREEN
            
            if not r['Media_Auto'] and r['Status_Av'] == 'Pendente':
                r['Media_Auto'] = ''
            elif r['Status_Av'] == 'Finalizado' and not r['Media_Auto']:
                r['Media_Auto'] = 'N/A'

            if not r['Media_Avaliadores'] and r['Status_Av'] == 'Pendente':
                r['Media_Avaliadores'] = ''
            elif r['Status_Av'] == 'Finalizado' and not r['Media_Avaliadores']:
                r['Media_Avaliadores'] = 'N/A'

            if not r['Media_Auto_Desemp'] and r['Status_Av'] == 'Pendente':
                r['Media_Auto_Desemp'] = ''
            elif r['Status_Av'] == 'Finalizado' and not r['Media_Auto_Desemp']:
                r['Media_Auto_Desemp'] = 'N/A'

            if not r['Media_Aval_Desemp'] and r['Status_Av'] == 'Pendente':
                r['Media_Aval_Desemp'] = ''
            elif r['Status_Av'] == 'Finalizado' and not r['Media_Aval_Desemp']:
                r['Media_Aval_Desemp'] = 'N/A'

            linha = ft.Container(
                content=ft.Row([
                        ft.Container(ft.Text(r["Participante"]), expand=3),
                        ft.Container(ft.Text(r["Avaliacao"]), expand=1),
                        ft.Container(ft.Text(r["Sigla_Emp"]), expand=1),
                        ft.Container(ft.Text(r["C_Custo"]), expand=1),
                        ft.Container(ft.Text(r["Cargo"]), expand=1),
                        ft.Container(ft.Text(r["Avaliacao"]), expand=1),
                        ft.Container(ft.Text(r["Media_Auto"]), expand=1),
                        ft.Container(ft.Text(r["Media_Avaliadores"]), expand=1),
                        ft.Container(ft.Text(r["Media_Auto_Desemp"]), expand=1),
                        ft.Container(ft.Text(r["Media_Aval_Desemp"]), expand=1),
                        ft.Container(ft.Text(r["Performance"]), expand=1),
                        ft.Container(
                            content=ft.Text(r["Status_Av"]),
                            bgcolor=bg_cor,
                            padding=ft.padding.symmetric(horizontal=6, vertical=2),
                            border_radius=5,
                            alignment=ft.alignment.center,
                            expand=1
                        ),
                    ], spacing=10),
                padding=10,
                bgcolor=ft.Colors.TRANSPARENT,
                border_radius=8,
                margin=ft.margin.only(bottom=4)
            )
            
            lista_grafico.controls.append(linha)

        lista_grafico.update()
        aguarde_overlay.visible = False
        page.update()
    
    empresas_selecionadas = []
    c_custos_selecionados = []
    cargos_selecionados = []

   

    # Função para atualizar query
    def atualizar_query(e=None):
        query_parts = []

        empresas_selecionadas.clear()
        c_custos_selecionados.clear()
        cargos_selecionados.clear()

        query_cargo = ''
        query_emp = ''
        query_ccusto = ''

        for chk in dropdown_c_custo.controls[0].controls:
            if chk.value:
                c_custos_selecionados.append(chk.label)

        for chk in dropdown_empresa.controls[0].controls:  # ListView
            if chk.value:
                empresas_selecionadas.append(chk.label)
         
        for chk in dropdown_cargo.controls[0].controls:
            if chk.value:
                cargos_selecionados.append(chk.label)
        
        if empresas_selecionadas:
            query_emp = "Sigla_Emp in ('" + "', '".join(empresas_selecionadas) + "')"
            query_parts.append(query_emp)

        if c_custos_selecionados:
            query_ccusto = "C_Custo in ('" + "', '".join(c_custos_selecionados) + "')"
            query_parts.append( query_ccusto)

        if cargos_selecionados:
            query_cargo = "Cargo in ('" + "', '".join(cargos_selecionados) + "')"
            query_parts.append(query_cargo)


        
        #Atualiza o filtro de centro de custos
        if query_emp != '':
            conn = mysql_connection()
            cursor = conn.cursor()
            query_cons = query_emp
            cursor.execute(f"SELECT DISTINCT C_Custo FROM QuestRH_Pessoas where not isnull(C_Custo) and {query_cons} and C_Custo <> '' order by C_Custo")
            resultados = cursor.fetchall()
            conn.close()

            dropdown_c_custo.controls[0].controls.clear()

            for res in resultados:
                 dropdown_c_custo.controls[0].controls.append(ft.Checkbox(label=res[0]))
                 if res[0] in query_ccusto:
                     dropdown_c_custo.controls[0].controls[-1].value = True
        else:
            conn = mysql_connection()
            cursor = conn.cursor()
            cursor.execute(f"SELECT DISTINCT C_Custo FROM QuestRH_Pessoas where not isnull(C_Custo) and C_Custo <> ''  order by C_Custo")
            resultados = cursor.fetchall()
            conn.close()

            dropdown_c_custo.controls[0].controls.clear()

            for res in resultados:
                 dropdown_c_custo.controls[0].controls.append(ft.Checkbox(label=res[0]))
                 if res[0] in query_ccusto:
                     dropdown_c_custo.controls[0].controls[-1].value = True
        
        #Atualiza o filtro de cargos
        if query_ccusto != '':
            conn = mysql_connection()
            cursor = conn.cursor()
            query_cons = query_ccusto
            cursor.execute(f"SELECT DISTINCT Cargo FROM QuestRH_Pessoas where not isnull(Cargo) and {query_cons} and Cargo <> '' order by Cargo")
            resultados = cursor.fetchall()
            conn.close()

            dropdown_cargo.controls[0].controls.clear()

            for res in resultados:
                 dropdown_cargo.controls[0].controls.append(ft.Checkbox(label=res[0]))
                 if res[0] in query_cargo:
                     dropdown_cargo.controls[0].controls[-1].value = True
        else:
            conn = mysql_connection()
            cursor = conn.cursor()
            cursor.execute(f"SELECT DISTINCT Cargo FROM QuestRH_Pessoas where not isnull(Cargo) and Cargo <> '' order by Cargo")
            resultados = cursor.fetchall()
            conn.close()

            dropdown_cargo.controls[0].controls.clear()

            for res in resultados:
                 dropdown_cargo.controls[0].controls.append(ft.Checkbox(label=res[0]))
                 if res[0] in query_cargo:
                     dropdown_cargo.controls[0].controls[-1].value = True

        if chk_estrategico_grafico.value == False and chk_nao_estrategico_grafico.value == True:
            query_parts.append("Grupo_estrategico = 'Não'")
        elif chk_estrategico_grafico.value == True and chk_nao_estrategico_grafico.value == False:
            query_parts.append("Grupo_estrategico = 'Sim'")

        txt_outro_query.value = " and ".join(query_parts) if query_parts else ''
        
        atualiza_dados()
        #return

    #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------- Inicio Função De Diálogo --------------------------------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    txt_mensagem_snack = ft.Text('')
    # 1. Criamos o objeto SnackBar
    snack = ft.SnackBar(
        content=txt_mensagem_snack,
        action="Fechar"
    )
    
    
    def enviar_msg(mensagem, cor=ft.Colors.BLACK):
        # 2. Adicionamos ao overlay da página (se já não estiver lá)
        if snack not in page.overlay:
            page.overlay.append(snack)
        txt_mensagem_snack.value = mensagem
        # 3. Ativamos a abertura
        snack.open = True
        snack.bgcolor = cor
        page.update()
        

    def restaurar_backup_dinamico(e):
        try:
            conn = mysql_connection()
            cursor = conn.cursor()

            tabela_original = "QuestRH_Respostas"
            tabela_bkp = "QuestRH_Resp_BKP"
            
            # 1. Buscar os nomes das colunas da tabela ORIGINAL
            # Isso evita que a gente tente inserir a coluna 'data_bkp'
            cursor.execute(f"SHOW COLUMNS FROM {tabela_original}")
            colunas = [row[0] for row in cursor.fetchall()]
            
            # Transforma a lista ['id', 'nome', 'email'] em uma string "id, nome, email"
            nomes_colunas = ", ".join(colunas)

            # 2. Limpar a tabela original antes de restaurar
            cursor.execute("SET FOREIGN_KEY_CHECKS = 0;")
            cursor.execute(f"TRUNCATE TABLE {tabela_original}")

            # 3. Montar o SQL dinâmico
            # Selecionamos apenas as colunas que existem na original, ignorando a data_bkp
            sql = f"""
                INSERT INTO {tabela_original} ({nomes_colunas}) 
                SELECT {nomes_colunas} FROM {tabela_bkp}
                WHERE data_bkp = (SELECT MAX(data_bkp) FROM {tabela_bkp})
            """
            
            cursor.execute(sql)
            cursor.execute("SET FOREIGN_KEY_CHECKS = 1;")
            
            conn.commit()

            # Feedback para o usuário
            snack = ft.SnackBar(ft.Text("Restauração concluída com sucesso!"))
            page.overlay.append(snack)
            snack.open = True

        except Exception as err:
            enviar_msg(f"Erro: {err}", ft.Colors.RED_600)
        finally:
                conn.close()
                page.update()

   
    def reiniciar_ciclo(e=None):
            #Função enviar os dados para o back-up
            if texto_ola.value != 'Olá, Administrador':
                enviar_msg ('Seu usuário não possui a permissão de reiniciar ciclo, contate ao programador.',ft.Colors.RED_600)
                return
            else:
                conn = mysql_connection()
                cursor = conn.cursor(pymysql.cursors.DictCursor)
                sql = "SELECT Senha FROM QuestRH_Pessoas where Nome = 'Administrador'"
                cursor.execute(sql)
                resultados = cursor.fetchall()
                conn.commit()
                conn.close()
                
                if senha_txt.value != resultados[0]['Senha']:
                    enviar_msg ('Seu usuário não possui a permissão de reiniciar ciclo, contate ao programador.',ft.Colors.RED_600)
                    return
            
            if container_expiração.visible != True:
                enviar_msg ('Para reiniciar o ciclo o mesmo deve ter se encerrado.',ft.Colors.RED_600)
                return
            
            try:
                conn = mysql_connection()
                cursor = conn.cursor(pymysql.cursors.DictCursor)

                sql_delete = "DELETE from QuestRH_Resp_BKP where DATE(data_bkp) = CURDATE()"
                cursor.execute(sql_delete)
                
                # O comando transfere tudo e carimba a data atual em cada linha
                sql = "INSERT INTO QuestRH_Resp_BKP SELECT *, NOW() FROM QuestRH_Respostas"
                
                cursor.execute(sql)
                conn.commit() # Importante para salvar as alterações!
                
                sql_delete = "DELETE from QuestRH_Respostas"
                cursor.execute(sql_delete)
                conn.commit() # Importante para salvar as alterações!

                enviar_msg(f"{cursor.rowcount} registros copiados para o backup!", ft.Colors.GREEN_600)
            except Exception as e:
                enviar_msg(f"Houve um erro ao enviar ao banco de dados: {e}", ft.Colors.RED_600 )
            finally:
                cursor.close()
                conn.close()
                page.update()
   
   
   
    def alimenta_chkbox(e=None):
            conn = mysql_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT Sigla_Emp FROM QuestRH_Pessoas where not isnull(Sigla_Emp) order by Sigla_Emp")
            empresas = cursor.fetchall()
            conn.close()

            dropdown_empresa.controls[0].controls.clear()
            dropdown_c_custo.controls[0].controls.clear()
            dropdown_cargo.controls[0].controls.clear()

            for emp in empresas:
                dropdown_empresa.controls[0].controls.append(ft.Checkbox(label=emp[0]))
                

            conn = mysql_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT C_Custo FROM QuestRH_Pessoas where not isnull(C_Custo) and C_Custo <> ''  order by C_Custo")
            custos = cursor.fetchall()
            conn.close()

            for c in custos:
                dropdown_c_custo.controls[0].controls.append(ft.Checkbox(label=c[0]))
               

            conn = mysql_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT Cargo FROM QuestRH_Pessoas where not isnull(Cargo) and Cargo <> '' order by Cargo")
            cargos = cursor.fetchall()
            conn.close()
            for cg in cargos:
                dropdown_cargo.controls[0].controls.append(ft.Checkbox(label=cg[0]))
                

            page.update()



    def abrir_dialogo(titulo, container_dropdown):
            def confirmar(e):
                dlg.open = False
                atualizar_query()
                page.update()
            
            def cancelar(e):
                dlg.open = False
                page.update()

            def selecionar_todos(e):
                for chk in container_dropdown.controls[0].controls:
                    chk.value = True
                page.update()

            def limpar_todos(e):
                for chk in container_dropdown.controls[0].controls:
                    chk.value = False
                page.update()

            dlg = ft.AlertDialog(
                modal=True,
                title=ft.Text(f"Selecione {titulo}"),
                content=ft.Container(
                    content=container_dropdown,
                    width=400,
                    height=300,  # altura fixa com rolagem
                    padding=10,
                    bgcolor=ft.Colors.WHITE,
                ),
                actions=[
                    ft.TextButton("Cancelar", on_click=cancelar),
                    ft.TextButton("Selecionar todos", on_click=selecionar_todos),
                    ft.TextButton("Limpar todos", on_click=limpar_todos),
                    ft.TextButton("Confirmar", on_click=confirmar),
                ],
                actions_alignment="spaceBetween",
            )

            page.open(dlg)



    #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------- Inicio Demais Funcoes --------------------------------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    last_interaction = {"time": time.monotonic()}
    tempo_ociosidade = {"segundos": 0}
 
    def reset_idle_time(e=None):
        #print('timer resetado')
        last_interaction["time"] = time.monotonic()
        txt_ociosidade.value = 'Tempo de inatividade: 0s (00:00:00) - tempo limite: 00:05:00 (5 minutos)'
        tempo_ociosidade["segundos"] = 0
        txt_ociosidade.update()

        if deslog_overlay.visible == True:
            deslog_overlay.visible = False
            page.update()

    async def check_idle():
        while True:
            if Gestor_tempo_formulario.visible == True:
                await asyncio.sleep(1)
                
                now = time.monotonic()
                elapsed = now - last_interaction["time"]
                remaining = int(IDLE_TIMEOUT - elapsed)
                tempo_ociosidade["segundos"] += 1 
                tempo_delta = datetime.timedelta(seconds=tempo_ociosidade["segundos"])
                tempo_formatado = str(tempo_delta)

                txt_ociosidade.value = f'Tempo de inatividade: {tempo_ociosidade["segundos"]}s ({tempo_formatado}) - tempo limite: 00:05:00 (5 minutos)'
                txt_ociosidade.update()
             


                if remaining > 0:
                    if remaining <= 40:
                        deslog_overlay.visible = True

                    msg_deslog.value = f"Você está muito tempo inativo, está por aí? o questionário será fechado em {remaining} segundos"
                else:
                    deslog_overlay.visible = False
                    enviar_formulario(None,True)
                    painel_view.visible =True
                    Gestor_tempo_formulario.visible =False
                    page.update()
                    return
            else:
                reset_idle_time()
                return
            page.update()

    def abrir_manual(e):
        nome = texto_ola1.value.replace('Olá, ','')
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        scrp_sql = f"SELECT * FROM QuestRH_Relacoes Where Participante = '{nome}'"
        cursor.execute(scrp_sql)
        resultado = cursor.fetchone()
        conn.close()

        if resultado:
            if resultado['Tipo_Avaliacao'] == 'A1':
                page.launch_url("https://enindengenharia-my.sharepoint.com/:b:/g/personal/wagner_barreiro_enind_com_br/IQAZJERZ2OkaTq7A4sLZc08iAck-56yCulYBHoqq5Qhbfmw?e=W9t1ry")
            elif resultado['Tipo_Avaliacao'] == 'A2':
                page.launch_url("https://enindengenharia-my.sharepoint.com/:b:/g/personal/wagner_barreiro_enind_com_br/IQDhtSFov2WOTIzGnynPQO7uAW6dtdVv6c_KEaD1UyHLdhU?e=iGlIdf")
            else:
                page.launch_url("https://enindengenharia-my.sharepoint.com/:b:/g/personal/wagner_barreiro_enind_com_br/IQDr5yV6IQO5Qol2vGD_5bv9AUblf_GH6YwVH9893JRByx0?e=FE7dCb")
        else:
            page.launch_url("https://enindengenharia-my.sharepoint.com/:b:/g/personal/wagner_barreiro_enind_com_br/IQAZJERZ2OkaTq7A4sLZc08iAck-56yCulYBHoqq5Qhbfmw?e=W9t1ry")
    
    def atualizar_altura_container(e):
        altura_tela = e.height
    
        container_obs_auto.height = altura_tela * 0.5
        container_obs_av1.height = altura_tela * 0.5
        container_obs_av2.height = altura_tela * 0.5

        corpo_tabela_pend.height = altura_tela * 0.7
        corpo_tabela.height = altura_tela * 0.7
        container_perguntas.height = altura_tela * 0.7
        
        container_observacoes.height = altura_tela * 0.7
        txt_observacoes.height = altura_tela * 0.65
        container_dados_bellcurve.height = altura_tela * 0.4
        page.update()
    
    
    def voltar_login(e):
        nome_cb.value =''
        senha_txt.value =''
        
        page.controls.clear()
        overlay = ft.Column([
            alerta_container,
            login_view,
            ft.Row([txt_assinatura], alignment=ft.MainAxisAlignment.CENTER),
            ], 
        expand=True,
        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER
        )

        page.scroll = False
        
        page.add(overlay)
        login_view.visible =True
        page.update()
        animate_entrance()

    def voltar_painel(e=None, atualizar = False):
        mensagem_aguarde.value = 'Aguarde, atualizando o relatório...'
        aguarde_overlay.visible = True
        imagem_fundo.visible = True
        container_painel_grafico.visible = False
        page.update()
        nome_pessoa = texto_ola.value.replace("Olá, ","")

        if nome_pessoa == 'Administrador':
            painel_pend_view.visible = True
            painel_view.visible = False
            container_grafico.visible = False
            if atualizar == True:
                questionarios = lista_pendencias(True, True)
                montar_tabela_pendencias(questionarios)
        else:
            if Gestor_tempo_formulario.visible == True:
                enviar_formulario(None, True)
            
            painel_view.visible = True
            painel_pend_view.visible = False
            if atualizar == True:
                questionarios = obter_questionarios(nome_pessoa)
                montar_tabela(questionarios)

        Gestor_tempo_formulario.visible = False
        formulario_Envio.visible = False
        painel_resposta_view.visible= False

    
        aguarde_overlay.visible = False
        page.update()

    def atualiza_rel(e):
        nonlocal lista_dados
        nonlocal lista_pend
        nome_pessoa = texto_ola.value.replace("Olá, ",'')
        aguarde_overlay.visible = True
        lista_dados = []
        lista_pend = []
        page.update()

        chk_estrategico.value

        if nome_pessoa == 'Administrador':
            painel_pend_view.visible = True
            painel_view.visible = False
            questionarios = lista_pendencias(chk_estrategico.value, chk_nao_estrategico.value)
            montar_tabela_pendencias(questionarios)
        else:
            painel_view.visible = True
            painel_pend_view.visible = False
            questionarios = obter_questionarios(nome_pessoa)
            montar_tabela(questionarios)

        painel_resposta_view.visible= False

        aguarde_overlay.visible = False
        page.update()

    # Função chamada ao clicar "Sim" ou "Não"
    def confirmar_continuacao(e, resposta):
        confirmacao_overlay.visible = False
        reset_idle_time()
        if resposta:
            enviar_formulario(e)
        else:
            mostrar_alerta_temporario('Cancelamento de resposta realizado com sucesso!', ft.Colors.GREEN_400)
        page.update()

    # Botão principal para exibir o overlay
    def mostrar_confirmacao(e):
        confirmacao_overlay.visible = True
        reset_idle_time()
        page.update()


    
    # Elemento de avaliação
    nome_em_avaliacao = ft.Text("", size=25, weight=ft.FontWeight.BOLD,font_family = 'Century Gothic')
    nome_avaliado = ft.Text("", size=20, weight=ft.FontWeight.BOLD, font_family = 'Century Gothic')

    #Função para validar o Login
    def validar_login(e):

        if valida_texto(nome_cb.value) == False:
            mostrar_alerta_temporario('login possui caracteres e palavras inválidas', ft.Colors.RED_400)
            return
        

        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        scrp_sql = f"SELECT * FROM QuestRH_Pessoas Where Login = '{nome_cb.value}'"
        cursor.execute(scrp_sql)
        resultados = cursor.fetchone()
        conn.close()

        if not resultados:
            mostrar_alerta_temporario('login não encontrado na base', ft.Colors.RED_400)
        else:
            if senha_txt.value == resultados['Senha']:
                page.scroll = ft.ScrollMode.AUTO
                #token = gera_token(resultados['Nome'])
                texto_ola.value = f"Olá, {resultados['Nome']}"
                texto_ola1.value = f"Olá, {resultados['Nome']}"
                texto_ola2.value = f"Olá, {resultados['Nome']}"
                texto_ola3.value = f"Olá, {resultados['Nome']}"
                texto_ola4.value = f"Olá, {resultados['Nome']}"
                texto_ola5.value = f"Olá, {resultados['Nome']}"
                texto_ola6.value = f"Olá, {resultados['Nome']}"
                mensagem_aguarde.value = 'Aguarde, verificando seu perfil...'
                aguarde_overlay.visible = True
                page.update()

                conn = mysql_connection()
                cursor = conn.cursor(pymysql.cursors.DictCursor)
                scrp_sql = f"Select * from QuestRH_Relacoes Where Avaliador1 = '{resultados['Nome']}'"
                cursor.execute(scrp_sql)
                avaliador1 = cursor.fetchall()
                conn.close()
            
                registra_login(nome_cb.value)

                if texto_ola.value == 'Olá, Administrador':
                    page.controls.clear()

                    overlay = ft.Column([
                        alerta_container,
                        painel_admin
                        ], 
                    expand=True,
                    alignment=ft.alignment.bottom_center 
                    )
                    stack = ft.Stack([
                        imagem_fundo,
                        overlay
                    ])
                    page.add(stack)

                    page.update()

                    questionarios = lista_pendencias(True, True)
                    montar_tabela_pendencias(questionarios)
                    painel_pend_view.visible = True
                    painel_resposta_view.visible= False 
                    
                elif avaliador1:
                    page.controls.clear()
                    btn_excel_av1.visible = True
                    btn_excel_av1.on_click = lambda e:preparar_relatório_final(e, av1=True, avaliador=resultados['Nome'])
                    overlay = ft.Column([
                        alerta_container,
                        painel_av1,
                        ft.Row([txt_assinatura], alignment=ft.MainAxisAlignment.CENTER),
                        ], 
                    expand=True,
                    alignment=ft.alignment.bottom_center 
                    )
                    stack = ft.Stack([
                        imagem_fundo,
                        overlay
                    ])
                    page.add(stack)

                    page.update()
                   

                    questionarios = obter_questionarios(resultados['Nome'])
                    montar_tabela(questionarios)
                    painel_view.visible = True
                    painel_resposta_view.visible= False 
                else:
                    page.clean()
                    btn_excel_av1.visible = False
                    overlay = ft.Column([
                        alerta_container,
                        painel_comum,
                        ft.Row([txt_assinatura], alignment=ft.MainAxisAlignment.CENTER),
                        ], 
                    expand=True,
                    alignment=ft.alignment.bottom_center 
                    )
                    stack = ft.Stack([
                        imagem_fundo,
                        overlay
                    ])
                    page.add(stack)

                    page.update()
                    

                    questionarios = obter_questionarios(resultados['Nome'])
                    montar_tabela(questionarios)
                    login_view.visible = False
                    painel_view.visible = True
                
                   
                login_view.visible = False
                aguarde_overlay.visible = False
                page.update()
                return
            else:
                mostrar_alerta_temporario('Senha incorreta',ft.Colors.RED_400)
        page.update()
        
       
    def realiza_feedback(e, nome):
        data_atual = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        scrp_sql = f"UPDATE QuestRH_Relacoes SET data_feedback = '{data_atual}' Where Participante = '{nome}'"
        cursor.execute(scrp_sql)
        conn.commit()
        conn.close()

        nome_avaliador = texto_ola.value.replace("Olá, ",'')
        questionario = obter_questionarios(nome_avaliador)
        montar_tabela(questionarios=questionario)
        page.update()

    #Função para criar a tabela de visualização - Painel de controle
    def montar_tabela(questionarios):
        nonlocal data_limite
        nonlocal linhas
        data_atual = datetime.datetime.now()
        data_fechamento = datetime.datetime.strptime(data_limite,"%Y/%m/%d %H:%M:%S") 
        lista_view.controls.clear()



        for q in questionarios:
            if q["status"] == "Pendente":
                Status = 'Pendente'
                bg_cor = ft.Colors.YELLOW_100
            elif q["status"] == "Realizado":
                if q['Questionário'] == 'A3':
                    Status = 'N/A'
                    bg_cor = ft.Colors.GREY_200
                else:
                    Status = 'Realizado'
                    bg_cor = ft.Colors.GREEN_100
            else:
                bg_cor = None

            btn_avaliar = (
                ft.TextButton('Avaliar',icon=ft.Icons.EDIT, on_click=lambda e, nome=q["nome"]: abrir_formulario(nome))
                if (texto_ola.value != "Olá, Administrador") and (q["status"] == "Pendente") and (data_atual <= data_fechamento) and (q['Questionário'] != 'A3' or (q['Questionário'] == 'A3' and q['Avaliador'] != 0)) else ft.Text("")
            )


            #btn visualizar
            btn_ver_resp = (
                ft.TextButton('', icon=ft.Icons.VISIBILITY, on_click=lambda e, nome=q["nome"]: abrir_formulario_respostas(nome))
                if (q['Avaliador'] >= 1 ) else ft.Text("")
            )
            
            
            #btn visualizar
            if q['data_feedback'] == '':
                contole_feedback = (
                    ft.TextButton('Realizar',icon=ft.Icons.FEEDBACK, on_click=lambda e, nome=q["nome"]: realiza_feedback(e, nome), disabled=(data_atual <= data_fechamento))
                    if (q['Avaliador'] == 1) else ft.Text("")
                )
            else:
                contole_feedback = ft.Text(f'{q["data_feedback"]}')
            
            nome_pessoa = q['nome']
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            scrp_sql = f"Select Nome, link_foto from QuestRH_Pessoas Where Nome ='{nome_pessoa}'"
            cursor.execute(scrp_sql)
            link_foto_resultado = cursor.fetchone()
            cursor.close()
            conn.close()
            
            if link_foto_resultado:
                if os.path.exists(os.path.join(BASE_DIR, "assets", "fotos", f"{link_foto_resultado['Nome']}.jpg")):
                    # Note que REMOVEMOS o "assets" da string que vai para o componente
                    link_foto = f"/fotos/{link_foto_resultado['Nome'].strip()}.jpg"
                else:
                    link_foto = link_foto_resultado['link_foto']

                
           
            
            # Dados que vão rolar
            linha = ft.Container(
                content=ft.Row([
                        ft.Container(ft.ListTile(leading=ft.CircleAvatar(
                            foreground_image_src=f'{link_foto}', #imagem rosto,
                            radius=22, # Tamanho razoável
                            content=ft.Text(str(q["nome"][0]).upper()),# Letra inicial se a imagem falhar
                            
                        ), bgcolor=ft.Colors.WHITE, on_click=lambda e, f=link_foto, n=q["nome"]: abrir_foto_popup(f,n)), expand=1),
                        ft.Container(ft.Text(q["nome"]), expand=3),
                        ft.Container(ft.Text(q["auto_aval"]), expand=1),
                        ft.Container(ft.Text(q["primaria"]), expand=1),
                        ft.Container(ft.Text(q["secundaria"]), expand=1),
                        ft.Container(
                            content=ft.Text(Status),
                            bgcolor=bg_cor,
                            padding=ft.padding.symmetric(horizontal=6, vertical=2),
                            border_radius=5,
                            alignment=ft.alignment.center,
                            expand=1
                        ),
                        ft.Container(btn_avaliar, expand=1),
                        ft.Container(btn_ver_resp, expand=1),
                        ft.Container(contole_feedback, expand=1)
                    ], spacing=10),
                padding=10,
                bgcolor=ft.Colors.WHITE,
                border_radius=12,
                margin=ft.margin.only(bottom=10),
                animate_opacity=300,
                shadow=ft.BoxShadow(
                    blur_radius=10,
                    color=ft.Colors.with_opacity(0.05, ft.Colors.BLACK),
                    offset=ft.Offset(0, 4)
                )
            )
            
            lista_view.controls.append(linha)

            
        if data_atual >=data_fechamento:
            container_expiração.visible = True

        lista_pend_view.visible = True
        lista_view.update()
        page.update()
        
    lista_dados = []
    lista_pend = []
    
    #Função para criar a tabela de visualização - Painel de controle - PENDENCIAS
    def montar_tabela_pendencias(questionarios):
        nonlocal data_limite
        nonlocal linhas
        nonlocal lista_dados
        nonlocal lista_pend
        data_atual = datetime.datetime.now()
        data_fechamento = datetime.datetime.strptime(data_limite,"%Y/%m/%d %H:%M:%S") 

        lista_pend_view.controls.clear()

        for q in questionarios:
            if q["Status"] == "Pendente":
                Status = 'Pendente'
                bg_cor = ft.Colors.YELLOW_100
            elif q["Status"] == "Realizado":
                if q["Questionário"] == 'A3':
                    Status = 'Não se Aplica'
                    bg_cor = ft.Colors.GREY_200
                else:
                    Status = 'Realizado'
                    bg_cor = ft.Colors.GREEN_100
            else:
                bg_cor = None

            if q["Status1"] == "Pendente":
                bg_cor1 = ft.Colors.YELLOW_100
            elif q["Status1"] == "Realizado":
                bg_cor1 = ft.Colors.GREEN_100
            else:
                bg_cor1 = None

            if q["Status2"] == "Pendente":
                bg_cor2 = ft.Colors.YELLOW_100
            elif q["Status2"] == "Realizado":
                bg_cor2 = ft.Colors.GREEN_100
            else:
                bg_cor2 = None

            btn_ver_resp = (
                ft.TextButton('', icon=ft.Icons.VISIBILITY, on_click=lambda e, nome=q["Participante"]: abrir_formulario_respostas(nome))
            )
            
            nome_pessoa = q['Participante']
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            scrp_sql = f"Select Nome, link_foto from QuestRH_Pessoas Where Nome ='{nome_pessoa}'"
            cursor.execute(scrp_sql)
            link_foto_resultado = cursor.fetchone()
            cursor.close()
            conn.close()
            
            if link_foto_resultado:
                if os.path.exists(os.path.join(BASE_DIR, "assets", "fotos", f"{link_foto_resultado['Nome']}.jpg")):
                    # Note que REMOVEMOS o "assets" da string que vai para o componente
                    link_foto = f"/fotos/{link_foto_resultado['Nome'].strip()}.jpg"
                else:
                    link_foto = link_foto_resultado['link_foto']

            qtd_pend = 0
            if q["Status"] == 'Pendente': 
                qtd_pend += 1
                if not q['Participante'] in lista_pend:
                    lista_pend.append(q['Participante'])
            if q["Status1"] == 'Pendente': 
                qtd_pend += 1
                if not q['Avaliador1'] in lista_pend:
                    lista_pend.append(q['Avaliador1'])
                
            if q["Status2"] == 'Pendente': 
                qtd_pend += 1
                if not q['Avaliador2'] in lista_pend:
                    lista_pend.append(q['Avaliador2'])
           
            if q["Status1"] =='Realizado' and q["Status2"] == 'Realizado':
                status_avaliadores = 'Realizado'
            else:
                status_avaliadores = 'Pendente'
            
            if qtd_pend>0:
                status = 'Pendente'
            else:
                status = 'Realizado'
            
            
            
            lista_dados.append(
                {
                    "Participante": q["Participante"],
                    "Status": q["Status"],
                    "Avaliacao": q["Avaliacao"],
                    "Avaliador1": q["Avaliador1"],
                    "Status1": q["Status1"],
                    "Avaliacao1": q["Avaliacao1"],
                    "Avaliador2": q["Avaliador2"],
                    "Status2": q["Status2"],
                    "Avaliacao2": q["Avaliacao2"],
                    "Qtd Pendente": qtd_pend,
                    "Status Global": status,
                    "Status Avaliadores": status_avaliadores,
                    'Feedback': q["data_feedback"]
                }
            )
            
             # Dados que vão rolar
            linha = ft.Container(
                content=ft.Row([
                         ft.Container(ft.ListTile(leading=ft.CircleAvatar(
                            foreground_image_src=f'{link_foto}', #imagem rosto,
                            radius=22, # Tamanho razoável
                            content=ft.Text(str(q["Participante"][0]).upper()),# Letra inicial se a imagem falhar
                        ), bgcolor=ft.Colors.WHITE
                        , on_click=lambda e, f=link_foto, n=q["Participante"]: abrir_foto_popup(f,n)), expand=1),
                        ft.Container(ft.Text(q["Participante"], size=10), expand=2),
                        ft.Container(
                            content=ft.Text(Status, size=10),
                            bgcolor=bg_cor,
                            padding=ft.padding.symmetric(horizontal=6, vertical=2),
                            border_radius=5,
                            alignment=ft.alignment.center,
                            expand=1
                        ),
                        ft.Container(ft.Text(q["Avaliacao"]), expand=1, alignment=ft.alignment.center),
                        ft.Container(ft.Text(q["Avaliador1"], size=10), expand=2),
                        ft.Container(
                            content=ft.Text(q["Status1"], size=10),
                            bgcolor=bg_cor1,
                            padding=ft.padding.symmetric(horizontal=6, vertical=2),
                            border_radius=5,
                            alignment=ft.alignment.center,
                            expand=1
                        ),
                        ft.Container(ft.Text(q["Avaliacao1"]), expand=1, alignment=ft.alignment.center),
                        ft.Container(ft.Text(q["Avaliador2"], size=10), expand=2),
                         ft.Container(
                            content=ft.Text(q["Status2"], size=10),
                            bgcolor=bg_cor2,
                            padding=ft.padding.symmetric(horizontal=6, vertical=2),
                            border_radius=5,
                            alignment=ft.alignment.center,
                            expand=1
                        ),
                        ft.Container(ft.Text(q["Avaliacao2"]), expand=1, alignment=ft.alignment.center),
                        ft.Container(ft.Text(q["data_feedback"]), expand=1, alignment=ft.alignment.center),
                        ft.Container(btn_ver_resp, expand=1)
                    ], spacing=10),
                padding=10,
                bgcolor=ft.Colors.WHITE,
                shadow=ft.BoxShadow(
                    blur_radius=10,
                    color=ft.Colors.with_opacity(0.05, ft.Colors.BLACK),
                    offset=ft.Offset(0, 4)
                ),
                border_radius=8,
                margin=ft.margin.only(bottom=4)
            )
            
            lista_pend_view.controls.append(linha)
        
        
        baixar_rel_final.visible = True
        baixar_rel_final_av1.visible = True
            
        if data_atual >=data_fechamento:
            container_expiração.visible = True
            btn_restaurar.visible = True
            btn_reiniciar_ciclo.visible = True
            
            

        lista_pend_view.visible = True
        lista_pend_view.update()
        page.update()
    
    def abrir_foto_popup(foto_caminho: str, nome: str):
        """
        Abre um pop-up profissional e clean para visualizar a foto.
        """
        
        # Referência para o diálogo para que as ações de fechar funcionem
        dlg_ref = ft.Ref[ft.AlertDialog]()

        # Função auxiliar para fechar
        def fechar(e):
            dlg_ref.current.open = False
            page.update()

        # O Diálogo em si
        dlg = ft.AlertDialog(
            ref=dlg_ref,
            content_padding=0,  # Removemos o padding padrão para controlar nós mesmos
            bgcolor=ft.Colors.TRANSPARENT, # Fundo transparente para o AlertDialog, o Card será o fundo
            content=ft.Card(
                elevation=10, # Sombra projetada
                color=ft.Colors.SURFACE, # Cor de fundo do Card
                content=ft.Container(
                    padding=20, # Padding interno do Card
                    width=450,  # Largura fixa razoável
                    content=ft.Column(
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            # --- CABEÇALHO ---
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Column(
                                        controls=[
                                            ft.Text(nome, weight=ft.FontWeight.BOLD, size=18),
                                            ft.Text("Visualização de Foto", size=12, color=ft.Colors.ON_SURFACE_VARIANT),
                                        ],
                                        spacing=0,
                                    ),
                                    ft.IconButton(
                                        icon=ft.Icons.CLOSE_ROUNDED,
                                        icon_color=ft.Colors.ON_SURFACE_VARIANT,
                                        on_click=fechar,
                                        tooltip="Fechar",
                                    )
                                ]
                            ),
                            ft.Divider(height=1, color=ft.Colors.OUTLINE_VARIANT), # Linha sutil separadora
                            
                            # --- CONTEÚDO DA FOTO ---
                            ft.Container(
                                margin=ft.margin.only(top=20, bottom=10),
                                content=ft.Image(
                                    src=foto_caminho,
                                    fit=ft.ImageFit.CONTAIN,
                                    border_radius=12, # Cantos arredondados na imagem
                                ),
                                # Limitamos a altura máxima da imagem para não estourar telas menores
                                height=page.height * 0.6, 
                            ),
                        ],
                    ),
                ),
            ),
        )

        # Adicionamos ao overlay e abrimos
        page.overlay.append(dlg)
        dlg.open = True
        page.update()

     #Função para construir os formulários de respostas
    def abrir_formulario_respostas(Pessoa):
        container_respostas_auto.visible = False
        container_respostas_av1.visible = False
        container_respostas_av2.visible = False
      
        txt_av1.value = ''
        txt_av2.value  = ''
        txt_auto.value  = ''
        
        txt_status_auto.value = 'Não Realizado'
        container_status_auto.bgcolor = ft.Colors.GREY_100
        
        txt_status_av1.value = 'Não Realizado'
        container_status_av1.bgcolor = ft.Colors.GREY_100
        
        txt_status_av2.value = 'Não Realizado'
        container_status_av2.bgcolor = ft.Colors.GREY_100

        txt_observacoes_auto.value = ''
        txt_observacoes_av1.value = ''
        txt_observacoes_av2.value = ''


        lista_reultados_auto_view.controls.clear()
        lista_reultados_av1_view.controls.clear()
        lista_reultados_av2_view.controls.clear()
       
        nome_avaliado.value = f'Você está analisando: {Pessoa}'

        lista_avaliação =[
            {'id':1, 'Resp':"1 - Insatisfatório - Não atende ou atende minimamente aos padrões", 'cor':ft.Colors.RED_100},
            {'id':2, 'Resp':"2 - Regular - Atende parcialmente aos padrões esperados", 'cor':ft.Colors.ORANGE_100},
            {'id':3, 'Resp':"3 - Satisfatório - Atende os padrões esperados", 'cor':ft.Colors.RED_100},
            {'id':4, 'Resp':"4 - Bom - Demonstra empenho e excelência no atendimento de padrões esperados", 'cor':ft.Colors.YELLOW_100},
            {'id':5, 'Resp':"5 - Excelente - Supera os padrões esperados", 'cor':ft.Colors.GREEN_400},
        ]


        lista_reultados_auto_view.controls.clear()
        scrp_sql = f"SELECT * FROM QuestRH_Respostas WHERE Participante = '{Pessoa}'"
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchall()
        conn.close()
        
        

        
        pilar_atual = None
        competencia_atual = ''
        bloco_perguntas = []
        media_final = 0
        
        # Abrindo conexão
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        
        # Realizando Consulta de Relacoes
        scrp_sql = f"SELECT * FROM QuestRH_Relacoes WHERE Participante = '{Pessoa}'"
        cursor.execute(scrp_sql)
        relacoes = cursor.fetchone()

        #Auto avaliação
        scrp_sql = f"Select Nome, link_foto from QuestRH_Pessoas Where Nome ='{relacoes['Participante']}'"
        cursor.execute(scrp_sql)
        link_foto_resultado = cursor.fetchone()

        
        if link_foto_resultado:
            if os.path.exists(os.path.join(BASE_DIR, "assets", "fotos", f"{link_foto_resultado['Nome']}.jpg")):
                # Note que REMOVEMOS o "assets" da string que vai para o componente
                link_foto = f"/fotos/{link_foto_resultado['Nome'].strip()}.jpg"
            else:
                link_foto = link_foto_resultado['link_foto']
                
        txt_auto.value = f"{relacoes['Participante']}"
        leading_auto.content.value = str(relacoes['Participante'])[0].upper()
        leading_auto.foreground_image_src=f'{link_foto}'
        
       
        list_tile_auto.on_click=lambda e, f = link_foto, n=relacoes['Participante']: abrir_foto_popup(f,n)
        
        #Avaliação 1
        scrp_sql = f"Select Nome, link_foto from QuestRH_Pessoas Where Nome ='{relacoes['Avaliador1']}'"
        cursor.execute(scrp_sql)
        link_foto_resultado = cursor.fetchone()
    
        
        if link_foto_resultado:
            if os.path.exists(os.path.join(BASE_DIR, "assets", "fotos", f"{link_foto_resultado['Nome']}.jpg")):
                # Note que REMOVEMOS o "assets" da string que vai para o componente
                link_foto = f"/fotos/{link_foto_resultado['Nome'].strip()}.jpg"
            else:
                link_foto = link_foto_resultado['link_foto']
                
        txt_av1.value = f"{relacoes['Avaliador1']}"
        leading_av1.content.value = str(relacoes['Avaliador1'])[0].upper()
        leading_av1.foreground_image_src=f'{link_foto}'
        
        
        list_tile_av1.on_click=lambda e, f = link_foto, n=relacoes['Avaliador1']: abrir_foto_popup(f,n)
        
        
        #Avaliação 2
        scrp_sql = f"Select Nome, link_foto from QuestRH_Pessoas Where Nome ='{relacoes['Avaliador2']}'"
        cursor.execute(scrp_sql)
        link_foto_resultado = cursor.fetchone()
        
        
        if link_foto_resultado:
            if os.path.exists(os.path.join(BASE_DIR, "assets", "fotos", f"{link_foto_resultado['Nome']}.jpg")):
                # Note que REMOVEMOS o "assets" da string que vai para o componente
                link_foto = f"/fotos/{link_foto_resultado['Nome'].strip()}.jpg"
            else:
                link_foto = link_foto_resultado['link_foto']
                
        txt_av2.value = f"{relacoes['Avaliador2']}"
        leading_av2.content.value = str(relacoes['Avaliador2'])[0].upper()
        leading_av2.foreground_image_src=f'{link_foto}'
        
        
        list_tile_av2.on_click=lambda e, f= link_foto, n=relacoes['Avaliador2']: abrir_foto_popup(f,n)

        #fecha as conexões
        cursor.close()
        conn.close()

        def adicionar_container_pilar(bloco, row):
            container_pilar = ft.Container(
                content=ft.Column(bloco, spacing=10),
                padding=15,
                margin=ft.margin.only(bottom=20),
                bgcolor=ft.Colors.GREY_50,
                border=ft.border.all(1, ft.Colors.GREY_100),
                border_radius=12,
                expand=True
            )
            

            if row['ID_Rel'] == 0:
                txt_observacoes_auto.value = f"{row['Observacao']}" 
                

                if row['Desempenho_tecnico'] or row['Desempenho_tecnico']>0:
                    #container_desempenho_auto.visible =True
                    txt_desempenho_auto.value = lista_avaliação[row['Desempenho_tecnico'] - 1]['Resp']
                    container_desemp_auto.bgcolor = lista_avaliação[row['Desempenho_tecnico'] - 1]['cor']
                else:
                    txt_desempenho_auto.value =''
                    #container_desempenho_auto.visible =False

                lista_reultados_auto_view.controls.append(container_pilar)
                txt_status_auto.value = 'Realizado'
                container_status_auto.bgcolor = ft.Colors.LIGHT_GREEN_200
                #lista_reultados_auto_view.visible = True
                #container_respostas_auto.visible = True
                #container_obs_auto.visible = True

            elif row['ID_Rel'] == 1:
                txt_observacoes_av1.value = f"{row['Observacao']}" 
                txt_observacoes_av1_e.value = f"{row['obs_e']}"
                txt_observacoes_av1_p.value = f"{row['obs_p']}"
                txt_observacoes_av1_c.value = f"{row['obs_c']}"
            

                if row['Desempenho_tecnico'] or row['Desempenho_tecnico']>0:
                    #container_desempenho_av1.visible =True
                    
                    txt_desempenho_av1.value = lista_avaliação[row['Desempenho_tecnico'] - 1]['Resp']
                    container_desemp_av1.bgcolor = lista_avaliação[row['Desempenho_tecnico'] - 1]['cor']
                else:
                    txt_desempenho_av1.value =''
                    #container_desempenho_av1.visible =False

                lista_reultados_av1_view.controls.append(container_pilar)
                txt_status_av1.value = 'Realizado'
                container_status_av1.bgcolor = ft.Colors.LIGHT_GREEN_200
                #lista_reultados_av1_view.visible = True
                #container_respostas_av1.visible = True
                #container_obs_av1.visible = True
            else:
                txt_observacoes_av2.value = f"{row['Observacao']}" 
                txt_observacoes_av2_e.value = f"{row['obs_e']}"
                txt_observacoes_av2_p.value = f"{row['obs_p']}"
                txt_observacoes_av2_c.value = f"{row['obs_c']}"
                
                if row['Desempenho_tecnico'] or row['Desempenho_tecnico']>0:
                    #container_desempenho_av2.visible =True
                    txt_desempenho_av2.value = lista_avaliação[row['Desempenho_tecnico'] - 1]['Resp']
                    container_desemp_av2.bgcolor = lista_avaliação[row['Desempenho_tecnico'] - 1]['cor']
                else:
                    txt_desempenho_av2.value =''
                    #container_desempenho_av2.visible =False
        
                lista_reultados_av2_view.controls.append(container_pilar)
                txt_status_av2.value = 'Realizado'
                container_status_av2.bgcolor = ft.Colors.LIGHT_GREEN_200
                #lista_reultados_av2_view.visible = True
                #container_respostas_av2.visible = True
                #container_obs_av2.visible = True

        
        # Iterar sobre as linhas do resultado
        for i, row in enumerate(resultados):
            if row['Pilar'] != pilar_atual:
                # Se já havia um pilar em andamento, finalize-o
                if pilar_atual is not None:
                    adicionar_container_pilar(bloco_perguntas, resultados[i-1])

                # Iniciar novo bloco
                pilar_atual = row['Pilar']
                competencia_atual = ''
                bloco_perguntas = [
                    ft.Text(pilar_atual, size=30, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900)
                ]

            # Agrupamento por Competência
            if row['Competencia'] != competencia_atual:
                competencia_atual = row['Competencia']
                bloco_perguntas.append(ft.Text(f"Competência: {competencia_atual}", size=15, weight=ft.FontWeight.BOLD))

            # Texto da pergunta
            texto_pergunta = f"{row['ID_Pergunta']}) {row['Pergunta']}"
            obj_texto = ft.Text(texto_pergunta, size=18)

            # Container da resposta com cor
            container1 = ft.Container(
                bgcolor=lista_avaliação[row['Resposta'] - 1]['cor'],
                border_radius=10,
                padding=10,
                content=ft.Text(lista_avaliação[row['Resposta'] - 1]['Resp'])
            )

            bloco_perguntas.append(
                ft.Column(
                    controls=[obj_texto, container1],
                    spacing=5
                )
            )

            media_final += row['Resposta']

        # Adicionar o último pilar após o loop
        if bloco_perguntas:
            adicionar_container_pilar(bloco_perguntas, resultados[-1])

        nota_final = define_avaliacao_final(Pessoa)
        if nota_final != 0:
            txt_media_final.value = nota_final
            if nota_final == 5: 
                container_media_final.bgcolor=ft.Colors.GREEN_100
                txt_media_final.color=ft.Colors.GREEN_400
            elif nota_final == 4: 
                container_media_final.bgcolor=ft.Colors.LIGHT_GREEN_100
                txt_media_final.color=ft.Colors.LIGHT_GREEN_400
            elif nota_final == 3: 
                container_media_final.bgcolor=ft.Colors.YELLOW_100
                txt_media_final.color=ft.Colors.YELLOW_400
            elif nota_final == 2: 
                container_media_final.bgcolor=ft.Colors.ORANGE_100
                txt_media_final.color=ft.Colors.ORANGE_400
            elif nota_final == 1: 
                container_media_final.bgcolor=ft.Colors.RED_100
                txt_media_final.color=ft.Colors.RED_400
        else:
            txt_media_final.value = f'S/I'
            container_media_final.bgcolor=ft.Colors.GREY_100

        # Exibir painel de respostas

        lista_reultados_auto_view.controls.append(container_desempenho_auto)
        lista_reultados_av1_view.controls.append(container_desempenho_av1)
        lista_reultados_av2_view.controls.append(container_desempenho_av2)


        #lista_reultados_auto_view.controls.append(container_obs_auto)
        #lista_reultados_av1_view.controls.append(container_obs_av1)
        #lista_reultados_av2_view.controls.append(container_obs_av2)

        painel_pend_view.visible = False
        painel_view.visible = False
        painel_resposta_view.visible = True

        lista_reultados_auto_view.update()
        lista_reultados_av1_view.update()
        lista_reultados_av2_view.update()
    
        page.update()

    def muda_cor_dropdown(e: ft.ControlEvent):
        reset_idle_time()
        dropdown = e.control
        valor = dropdown.value
        container = dropdown.meu_container 

        # Exemplo de lógica: muda a cor conforme valor selecionado
        if valor == "1 - Insatisfatório - Não atende ou atende minimamente aos padrões":
            container.bgcolor = ft.Colors.RED_100
        elif valor == "2 - Regular - Atende parcialmente aos padrões esperados":
            container.bgcolor = ft.Colors.ORANGE_100
        elif valor == "3 - Satisfatório - Atende os padrões esperados":
            container.bgcolor = ft.Colors.YELLOW_100
        elif valor == "4 - Bom - Demonstra empenho e excelência no atendimento de padrões esperados":
            container.bgcolor = ft.Colors.LIGHT_GREEN_100
        elif valor == "5 - Excelente - Supera os padrões esperados":
            container.bgcolor = ft.Colors.GREEN_400
        else:
            container.bgcolor = ft.Colors.WHITE
        
        # Atualiza a interface
        container.update()

    
    
    def preparar_relatório_final(e=None, av1 = False, avaliador = ''):
        enviar_msg('Preparando relatório final...',ft.Colors.BLUE_400)
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        
        if avaliador != '':
            scrp_sql = f"Select * from QuestRH_Relacoes Where Avaliador1 = '{avaliador}'"
            cursor.execute(scrp_sql)
            avaliador1 = cursor.fetchall()
        
            if not avaliador1:
                enviar_msg('Você não é avaliador 1, não pode gerar relatório',ft.Colors.RED_400)
                cursor.close()
                conn.close()
                return
        
        sql_texto = '''
        SELECT ID, Participante, Cargo, C_Custo, Local, Avaliacao, ID_Pergunta, Pilar, Competencia, Pergunta,
            -- Transformação das linhas em colunas
            MAX(CASE WHEN ID_Rel = 0 THEN Resposta END) AS Resposta_Av0,
            MAX(CASE WHEN ID_Rel = 1 THEN Resposta END) AS Resposta_Av1,
            MAX(CASE WHEN ID_Rel = 2 THEN Resposta END) AS Resposta_Av2
        FROM 
            QuestRH_Respostas 
        GROUP BY 
            ID_Pergunta, Participante -- Agrupamos para consolidar os avaliadores na mesma linha
        ORDER BY 
            Participante, ID_Pergunta;
        '''
        
        cursor.execute(sql_texto)
        consulta = cursor.fetchall()
        cursor.close()
        conn.close()
        
        
        if av1== True:
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            
            if avaliador != '':
                sql_texto = f'''
                SELECT Participante, Avaliador1 From QuestRH_Relacoes Where Avaliador1 = '{avaliador}'
                '''
            else:
                sql_texto = '''
                SELECT Participante, Avaliador1 From QuestRH_Relacoes
                '''
            
            cursor.execute(sql_texto)
            consulta_avaliadores1 = cursor.fetchall()
            cursor.close()
            conn.close()
            
            df_avaliadores = pd.DataFrame(consulta_avaliadores1)
            avaliadores = df_avaliadores['Avaliador1'].unique()
        
        if consulta and av1 == False:
            participante_anterior = ''
            
            df_geral = pd.DataFrame(consulta)
            participantes = df_geral['Participante'].unique()
            
            nome_arquivo = f"Relatorio_final.xlsx"
            caminho_file = os.path.join("assets", "relatorios",  nome_arquivo)
            
            writer = pd.ExcelWriter(caminho_file, engine='xlsxwriter')
            
            for i, p in enumerate(participantes):
                df_participante = df_geral[df_geral['Participante'] == p]
            
                participante_anterior = p
                grafico64, msgerro = gera_graficos.gera_ninebox(participante=p)
                grafico_pilar64, msgerro = gera_graficos.gera_gráfico_pilar(participante=p)
                grafico_comp64, msgerro = gera_graficos.gera_gráfico_Competencia(participante=p)
                grafico_compar64, msgerro_compar = gera_graficos.gera_gráfico_Comparativo(participante=p)

                # Transformar dados em DataFrame para a tabela
                df = df_participante
                
                nome = p
                # Salvar o DataFrame na aba do colaborador
                nome_aba = str(p)[:31]
                nome_curto = "".join(c for c in str(p) if c not in r"/\?*[]:")[:31]
                df.to_excel(writer, sheet_name=nome_aba, index=False, startrow=45)
                
        

                # --- INSERIR NO EXCEL ---
                worksheet = writer.sheets[nome_aba]
                worksheet.write(0, 0, f"{nome}")
                
                # Inserir a imagem do buffer diretamente na célula E2
                if grafico64:
                    img_bytes = io.BytesIO(base64.b64decode(grafico64))
                    worksheet.insert_image('A2', f'ninbox_{nome}.png', {'image_data': img_bytes,
                                                                        'x_scale': 0.4,
                                                                        'y_scale': 0.4})

                if grafico_pilar64:
                    img_bytes = io.BytesIO(base64.b64decode(grafico_pilar64))
                    worksheet.insert_image('F2', f'pilar_{nome}.png', {'image_data': img_bytes,
                                                                        'x_scale': 0.72,
                                                                        'y_scale': 0.72})

                if grafico_comp64:
                    img_bytes = io.BytesIO(base64.b64decode(grafico_comp64))
                    worksheet.insert_image('A16', f'comp_{nome}.png', {'image_data': img_bytes,
                                                                        'x_scale': 0.57,
                                                                        'y_scale': 0.57})
        
                if grafico_compar64:
                    img_bytes = io.BytesIO(base64.b64decode(grafico_compar64))
                    worksheet.insert_image('A30', f'comparativo_{nome}.png', {'image_data': img_bytes,
                                                                        'x_scale': 0.57,
                                                                        'y_scale': 0.57})

                enviar_msg(f'Preparando relatório final... {i+1}/{len(participantes)} - participante: {p} [{round(((i+1)/len(participantes))*100,2)}%]',ft.Colors.BLUE_600) 
                
            writer.close()
        elif consulta and av1 == True:
            df_geral = pd.DataFrame(consulta)
            participantes = df_geral['Participante'].unique()
            
            nome_arquivo = f"Relatorio_final_av1.xlsx"
            caminho_file = os.path.join("assets", "relatorios",  nome_arquivo)
            
            writer = pd.ExcelWriter(caminho_file, engine='xlsxwriter')
            
            for i, av in enumerate(avaliadores):
                participantes = df_avaliadores[df_avaliadores['Avaliador1'] == av]['Participante'].unique()
                
                df_participante = df_geral[df_geral['Participante'].isin(participantes)]
                
                # 3. Prepara a cláusula WHERE para uma futura consulta SQL
                if len(participantes) > 0:
                    # O join coloca aspas ENTRE os nomes, e o f-string coloca nas extremidades
                    sql_outras = "Participante IN ('" + "', '".join(participantes) + "')"
                else:
                    sql_outras = "Participante IN ('')" # Evita erro de SQL vazio
                
                grafico64, msgerro = gera_graficos.gera_ninebox(outras_condicoes=sql_outras)
                grafico_pilar64, msgerro = gera_graficos.gera_gráfico_pilar(outras_condicoes=sql_outras)
                grafico_comp64, msgerro = gera_graficos.gera_gráfico_Competencia(outras_condicoes=sql_outras)
                grafico_compar64, msgerro_compar = gera_graficos.gera_gráfico_Comparativo(outras_condicoes=sql_outras)

                # Transformar dados em DataFrame para a tabela
                df = df_participante
                
                nome = av
                # Salvar o DataFrame na aba do colaborador
                nome_aba = str(av)[:31]
                df.to_excel(writer, sheet_name=nome_aba, index=False, startrow=45)
                
                # --- INSERIR NO EXCEL ---
                worksheet = writer.sheets[nome_aba]
                worksheet.write(0, 0, f"{nome}")
                
                # Inserir a imagem do buffer diretamente na célula E2
                if grafico64:
                    img_bytes = io.BytesIO(base64.b64decode(grafico64))
                    worksheet.insert_image('A2', f'ninbox_{nome}.png', {'image_data': img_bytes,
                                                                        'x_scale': 0.4,
                                                                        'y_scale': 0.4})

                if grafico_pilar64:
                    img_bytes = io.BytesIO(base64.b64decode(grafico_pilar64))
                    worksheet.insert_image('F2', f'pilar_{nome}.png', {'image_data': img_bytes,
                                                                        'x_scale': 0.72,
                                                                        'y_scale': 0.72})

                if grafico_comp64:
                    img_bytes = io.BytesIO(base64.b64decode(grafico_comp64))
                    worksheet.insert_image('A16', f'comp_{nome}.png', {'image_data': img_bytes,
                                                                        'x_scale': 0.57,
                                                                        'y_scale': 0.57})
        
                if grafico_compar64:
                    img_bytes = io.BytesIO(base64.b64decode(grafico_compar64))
                    worksheet.insert_image('A30', f'comparativo_{nome}.png', {'image_data': img_bytes,
                                                                        'x_scale': 0.57,
                                                                        'y_scale': 0.57})

                enviar_msg(f'Preparando relatório final av1... {i+1}/{len(avaliadores)} - participante: {av} [{round(((i+1)/len(avaliadores))*100,2)}%]',ft.Colors.BLUE_600) 
                
            writer.close()
            
        
        else:
            enviar_msg('Sem dados para realizar a exportação', ft.Colors.RED_400)
        
        try:
            # 2. Certifique-se que a pasta assets existe na VPS 
            if not os.path.exists(os.path.join("assets", "relatorios")):
                os.makedirs(os.path.join("assets", "relatorios"))
                
                
            caminho_url = f"/relatorios/{nome_arquivo}"
            page.launch_url(f"{caminho_url}?t={datetime.datetime.now().timestamp()}")
            
            # Cria uma tarefa em segundo plano para deletar daqui a 30 segundos
            def deletar_depois():
                time.sleep(30) # Tempo suficiente para o navegador completar o download
                if os.path.exists(caminho_file):
                    os.remove(caminho_file)
                    
            threading.Thread(target=deletar_depois, daemon=True).start()
            
            mostrar_alerta_temporario('Exportação realizada com sucesso', ft.Colors.GREEN_400)
        except Exception as ex:
            mostrar_alerta_temporario(f'Erro ao exportar: {ex}', ft.Colors.RED_400)
    
    
    #Função para construir os formulários de respostas
    def abrir_formulario(nome):
        nome_em_avaliacao.value = f'Você está avaliando: {nome}'
        avaliador = texto_ola.value.replace('Olá, ','')
        
        #Limpando os campos
        txt_observacoes.value =''
        txt_observacoes_e.value = ''
        txt_observacoes_p.value = ''
        txt_observacoes_c.value = ''
        form_inputs.clear()
        form_content.controls.clear()
        perguntas_formulario = obter_perguntas(nome)
        dropdown_desempenho.value = 0
        
        
        #Diponibiliza recursos para av1 e av2
        exportar_pdf_btn_bkp.visible = nome != avaliador
        linha_calibracao.visible = nome != avaliador
        exportar_pdf_btn_antes.visible = nome != avaliador

        # Opções padrão
        opcoes_avaliacao = [
            ft.dropdown.Option("1 - Insatisfatório - Não atende ou atende minimamente aos padrões"),
            ft.dropdown.Option("2 - Regular - Atende parcialmente aos padrões esperados"),
            ft.dropdown.Option("3 - Satisfatório - Atende os padrões esperados"),
            ft.dropdown.Option("4 - Bom - Demonstra empenho e excelência no atendimento de padrões esperados"),
            ft.dropdown.Option("5 - Excelente - Supera os padrões esperados")
        ]

        alerta_container_form.visible = False
        txt_msg_alerta.value = ''

        # Agrupar perguntas por Pilar
        pilares_dict = {}
        for row in perguntas_formulario:
            pilar = row['Pilar']

            if pilar not in pilares_dict:
                pilares_dict[pilar] = []

            pilares_dict[pilar].append(row)

        # Construir cada bloco de pilar
        for pilar, perguntas in pilares_dict.items():
            bloco_perguntas = []

            bloco_perguntas.append(
                ft.Text(pilar, size=30, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_900)
            )

            competencia_atual = ''

            for row in perguntas:
                texto_pergunta = f"{row['ID']}) {row['Pergunta']}*"
                obj_texto = ft.Text(texto_pergunta, size=18)

                # Agrupamento por Competência
                if row['Competencia'] != competencia_atual:
                    competencia_atual = row['Competencia']
                    bloco_perguntas.append(
                        ft.Text(f"Competência: {competencia_atual}", size=15, weight=ft.FontWeight.BOLD)
                    )

                dropdown = ft.Dropdown(
                    options=opcoes_avaliacao,
                    hint_text="Avaliação",
                    on_change=muda_cor_dropdown,
                    expand= True
                )

                container1 = ft.Container(
                    bgcolor=ft.Colors.WHITE,
                    border_radius=10,
                    padding=10,
                    content=dropdown
                )
                
                container1.on_click = reset_idle_time
                container1.on_hover =reset_idle_time
                container1.on_long_press=reset_idle_time

                dropdown.on_change=muda_cor_dropdown
                dropdown.on_focus=reset_idle_time
                dropdown.on_blur=reset_idle_time
                
                dropdown.meu_container = container1
                
                
                #Recuperação de resposta da pergunta back-up rascunho
                conn = mysql_connection()
                cursor = conn.cursor(pymysql.cursors.DictCursor)
                texto_sql = f"Select Resposta from QuestRH_Rascunho where Nome_Avaliador = '{avaliador}'"
                texto_sql +=  f"and Participante = '{nome}' and ID_Pergunta = {row['ID']} and Pergunta = '{row['Pergunta']}'"
                cursor.execute(texto_sql)
                resposta= cursor.fetchone()
                conn.close()
                
                if resposta:
                    if  resposta['Resposta'] == 1:
                        dropdown.value = "1 - Insatisfatório - Não atende ou atende minimamente aos padrões"
                        container1.bgcolor =  ft.Colors.RED_100
                    elif  resposta['Resposta'] == 2:
                        dropdown.value ="2 - Regular - Atende parcialmente aos padrões esperados"
                        container1.bgcolor = ft.Colors.ORANGE_100
                    elif resposta['Resposta'] == 3:
                        dropdown.value ="3 - Satisfatório - Atende os padrões esperados"
                        container1.bgcolor = ft.Colors.YELLOW_100
                    elif resposta['Resposta'] == 4:
                        dropdown.value = "4 - Bom - Demonstra empenho e excelência no atendimento de padrões esperados"
                        container1.bgcolor = ft.Colors.LIGHT_GREEN_100
                    elif resposta['Resposta'] == 5:
                        dropdown.value = ("5 - Excelente - Supera os padrões esperados")
                        container1.bgcolor = ft.Colors.GREEN_400
                        # Exemplo de lógica: muda a cor conforme valor selecionado
                
                form_inputs.append(dropdown)
                bloco_perguntas.append( ft.Column(
                    controls=[
                        obj_texto,
                        container1
                    ],
                    spacing=5
                ))

            container_pilar = ft.Container(
                content=ft.Column(bloco_perguntas, spacing=10),
                padding=15,
                margin=ft.margin.only(bottom=20),
                bgcolor=ft.Colors.GREY_50,
                border=ft.border.all(1, ft.Colors.GREY_100),
                border_radius=12,
                expand=True
            )

            form_content.controls.append(container_pilar)

        obj_texto = ft.Text('Desempenho - Execução e entrega de atividades de forma geral:', size=18)
        dropdown_desempenho.options=opcoes_avaliacao

        container_desempenho = ft.Container(
            bgcolor=ft.Colors.WHITE,
            border_radius=10,
            padding=10,
            content=dropdown_desempenho
        )
        
        container_desempenho.on_click = reset_idle_time
        container_desempenho.on_hover =reset_idle_time
        container_desempenho.on_long_press=reset_idle_time

        dropdown_desempenho.on_change=muda_cor_dropdown
        dropdown_desempenho.on_focus=reset_idle_time
        dropdown_desempenho.on_blur=reset_idle_time


        dropdown_desempenho.meu_container = container_desempenho
        
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        texto_sql = f"Select Desempenho_tecnico, Observacao, obs_e, obs_p, obs_c from QuestRH_Rascunho where Nome_Avaliador = '{avaliador}'"
        texto_sql +=  f"and Participante = '{nome}'"
        cursor.execute(texto_sql)
        resposta= cursor.fetchone()
        conn.close()

        if resposta:
            if  resposta['Desempenho_tecnico'] == 1:
                dropdown_desempenho.value = "1 - Insatisfatório - Não atende ou atende minimamente aos padrões"
                container_desempenho.bgcolor = ft.Colors.RED_100
            elif  resposta['Desempenho_tecnico'] == 2:
                dropdown_desempenho.value ="2 - Regular - Atende parcialmente aos padrões esperados"
                container_desempenho.bgcolor = ft.Colors.ORANGE_100
            elif resposta['Desempenho_tecnico'] == 3:
                dropdown_desempenho.value ="3 - Satisfatório - Atende os padrões esperados"
                container_desempenho.bgcolor = ft.Colors.YELLOW_100
            elif resposta['Desempenho_tecnico'] == 4:
                dropdown_desempenho.value = "4 - Bom - Demonstra empenho e excelência no atendimento de padrões esperados"
                container_desempenho.bgcolor = ft.Colors.LIGHT_GREEN_100
            elif resposta['Desempenho_tecnico'] == 5:
                dropdown_desempenho.value = ("5 - Excelente - Supera os padrões esperados")
                container_desempenho.bgcolor = ft.Colors.GREEN_100
                
            txt_observacoes.value = resposta['Observacao']
            txt_observacoes_e.value = resposta['obs_e']
            txt_observacoes_p.value = resposta['obs_p']
            txt_observacoes_c.value = resposta['obs_c']

        form_content.controls.append(ft.Column(
                    controls=[
                        obj_texto,
                        container_desempenho
                    ],
                    spacing=5))
        
        #form_content.controls.append(txt_observacoes)

        painel_view.visible = False
        Gestor_tempo_formulario.visible = True
        reset_idle_time()
        page.update()
        page.run_task(check_idle)

    def validar_usuario(nome):
        if valida_texto(nome) == False:
            return False,'login possui caracteres e palavras inválidas'
        
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(f"SELECT Login, Senha From QuestRH_Pessoas where Nome = '{nome}'")
        resposta = cursor.fetchone()
        conn.close
        
        if resposta:
            if str(resposta['Login']).lower() == nome_cb.value.lower():
                if resposta['Senha'] == senha_txt.value:
                    return True,''
                else:
                    return False, 'Seu login não confere com seu nome, por gentileza, log-se novamente.'
            else:
                return False, 'Seu login não confere com seu nome, por gentileza, log-se novamente.'
        else:
            return False, 'Seu login não confere com seu nome, por gentileza, log-se novamente.'

    #Função para enviar as resposta dos usuários
    def enviar_formulario(e, backup=False,reinicio=True):
        respostas = []
        if backup == False:
            mensagem_aguarde.value = 'Aguarde, enviando respostas...'
        else:
            mensagem_aguarde.value = 'Aguarde, salvando rascunho...'
            
        aguarde_overlay.visible = True
        page.update()

        for i, grupo in enumerate(form_inputs):
            if not grupo.value and backup == False:
                mostrar_alerta_temporario("Preencha todos os campos antes de enviar o furmulário...", ft.Colors.RED_400)
                alerta_container_form.visible = True
                txt_msg_alerta.value = 'Preencha todos os campos antes de enviar o furmulário...'
                aguarde_overlay.visible = False
                page.update()
                return  # Interrompe envio
            else:
                valor_resp_antes = str(grupo.value)
                if valor_resp_antes:
                    valor_resp = str(grupo.value).split(' - ')[0]
                else:
                    valor_resp = ''
                respostas.append({
                    'ID':i+1,
                    'resposta': valor_resp})
        
        if (not dropdown_desempenho.value or dropdown_desempenho.value == 0) and backup == False :
            mostrar_alerta_temporario("Coloque antes o valor do desempenho da pessoa.", ft.Colors.RED_400)
            alerta_container_form.visible = True
            txt_msg_alerta.value = 'Coloque antes o valor do desempenho da pessoa.'
            aguarde_overlay.visible = False
            page.update()
            return  # Interrompe envio
        else:
            if dropdown_desempenho.value or dropdown_desempenho.value > 0:
                valor_desempenho = dropdown_desempenho.value.split(' - ')[0]
            else:
                valor_desempenho = ''
                
        def valida_campos_texto(texto, campo = ''):
            if texto == '':
               return False, 'Campo de ' + campo + ' vazio, por gentileza, coloque suas considerações'
            elif len(texto)<3: 
               return False, 'Campo de ' + campo + ' muito curto, por gentileza, aumente suas considerações'
       
            return True, ''
            
        avaliador = texto_ola.value.replace("Olá, ",'')
        participante = nome_em_avaliacao.value.replace('Você está avaliando: ', "")
        
        if backup == False:
        
            validacao_texto, msg =valida_campos_texto(txt_observacoes.value, 'observações')
            if validacao_texto == False:
                    mostrar_alerta_temporario(msg, ft.Colors.RED_400)
                    alerta_container_form.visible = True
                    txt_msg_alerta.value = msg
                    aguarde_overlay.visible = False
                    page.update()
                    return
            
            if participante != avaliador:  
                validacao_texto, msg =valida_campos_texto(txt_observacoes_e.value, 'calibração do pilar "E"')
                if validacao_texto == False:
                        mostrar_alerta_temporario(msg, ft.Colors.RED_400)
                        alerta_container_form.visible = True
                        txt_msg_alerta.value = msg
                        aguarde_overlay.visible = False
                        page.update()
                        return
                
                validacao_texto, msg =valida_campos_texto(txt_observacoes_p.value, 'calibração do pilar "P"')
                if validacao_texto == False:
                        mostrar_alerta_temporario(msg, ft.Colors.RED_400)
                        alerta_container_form.visible = True
                        txt_msg_alerta.value = msg
                        aguarde_overlay.visible = False
                        page.update()
                        return
                    
                validacao_texto, msg =valida_campos_texto(txt_observacoes_c.value, 'calibração do pilar "C"')
                if validacao_texto == False:
                        mostrar_alerta_temporario(msg, ft.Colors.RED_400)
                        alerta_container_form.visible = True
                        txt_msg_alerta.value = msg
                        aguarde_overlay.visible = False
                        page.update()
                        return
            else:
                txt_observacoes_e.value = 'Não se Aplica'   
                txt_observacoes_p.value = 'Não se Aplica' 
                txt_observacoes_c.value = 'Não se Aplica' 

        
        data_envio = datetime.datetime.now()
        data_envio_formatado = data_envio.strftime("%Y/%m/%d %H:%M:%S")
        
        validacao, erro = validar_usuario(avaliador)
        if validacao == False:
            voltar_login(None)
            enviar_msg(erro, ft.Colors.RED_400)
            page.update()
            return
        
        if valida_texto(participante) == False:
            enviar_msg('Participante inválido.')
            return
        
        #Limpando banco de rascunhos
        conn = mysql_connection()
        cursor = conn.cursor()
        sql_delete = f"DELETE From QuestRH_Rascunho where Participante = '{participante}' and Nome_Avaliador = '{avaliador}'"    
        cursor.execute(sql_delete)
        conn.commit()
        conn.close()

        try:
            #Resgatando dados das demais tabelas
            conn = mysql_connection()

            cursor = conn.cursor(pymysql.cursors.DictCursor)
            scrp_sql = f"SELECT * FROM QuestRH_Pessoas where Nome = '{participante}'"
            cursor.execute(scrp_sql)
            consulta_pessoa = cursor.fetchone()
            cursor.close()
            
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            scrp_sql = f"SELECT * FROM QuestRH_Relacoes where Participante = '{participante}'"
            cursor.execute(scrp_sql)
            consulta_relacao = cursor.fetchone()
            cursor.close()

            cursor = conn.cursor(pymysql.cursors.DictCursor)
            scrp_sql = f"SELECT * FROM QuestRH_Perguntas where Tipo_Avaliacao = '{consulta_relacao['Tipo_Avaliacao']}'"
            cursor.execute(scrp_sql)
            consulta_perguntas = cursor.fetchall()
            cursor.close()

            conn.close()

            if consulta_relacao['Participante'] == avaliador:
                ID_avaliador = 0
            elif consulta_relacao['Avaliador1'] == avaliador:
                ID_avaliador = 1
            elif consulta_relacao['Avaliador2'] == avaliador:
                ID_avaliador = 2
        except Exception as e:
            mostrar_alerta_temporario(f'Houve um erro em tempo de execução: {e}',ft.Colors.RED_400)
            aguarde_overlay.visible = False
            page.update()
            return
        
        
        usuario = nome_cb.value
        computador = platform.node()

        #Inserindo informações no banco de dados
        for row in respostas:
            if row['resposta']:
                resposta = row['resposta']
            else:
                resposta = 0
                
            campos = '''(Participante, Cargo, C_Custo,  Local, Avaliacao, Nome_Avaliador, ID_Rel, ID_Pergunta, Pilar, Competencia, Pergunta, Resposta,Desempenho_tecnico, Observacao, Data_Resp, Computador, Login, Empresa, Sigla_Emp, Grupo_estrategico, obs_e, obs_p, obs_c) 
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'''
            valores = (
                participante,
                consulta_pessoa['Cargo'],
                consulta_pessoa['C_Custo'],
                consulta_pessoa['Local'],
                consulta_relacao['Tipo_Avaliacao'],
                avaliador,
                ID_avaliador,
                int(row['ID']),
                consulta_perguntas[int(row['ID'])-1]['Pilar'],
                consulta_perguntas[int(row['ID'])-1]['Competencia'],
                consulta_perguntas[int(row['ID'])-1]['Pergunta'],
                resposta,
                valor_desempenho,
                txt_observacoes.value,
                data_envio_formatado,
                computador,
                usuario,
                consulta_pessoa['Empresa'],
                consulta_pessoa['Sigla_Emp'],
                consulta_relacao['Grupo_estrategico'],
                txt_observacoes_e.value,
                txt_observacoes_p.value,
                txt_observacoes_c.value
            )
            if backup == False:
                tabela_resp = 'QuestRH_Respostas'
            else:
                tabela_resp = 'QuestRH_Rascunho'
                
            validacao = inserir_banco(tabela_resp, valores, campos)
            
            if validacao ==False:
                mostrar_alerta_temporario("Não foi possível inerir os dados no banco de dados, tente novamente.", ft.Colors.RED_400)
                aguarde_overlay.visible = False
                page.update()
                return
            
        if backup == False:
            participante_realizado.value = participante
            formulario_Envio.visible = True
        
        
        if reinicio:
            Gestor_tempo_formulario.visible = False
            
        mensagem_aguarde.value = 'Aguarde, atualizando o relatório...'
        aguarde_overlay.visible = False
        page.update()
    
    def preparar_pdf(e=None, backup = False, anterior=False, adm= False):
        avaliadores = []
        login = nome_cb.value
        nome_log = texto_ola.value.replace("Olá, ",'')
        avaliador_pendente = ''
        verificacao, erro = validar_usuario(nome_log)
        
        if verificacao == False:
            voltar_login(None)
            enviar_msg(f'Login Inválido:{erro}',ft.Colors.RED_600)
            return
        
        if adm == False:
            enviar_formulario(None, True, False)
            participante = nome_em_avaliacao.value.replace('Você está avaliando: ','')
            avaliador_pendente = texto_ola.value.replace("Olá, ",'')
        else:
            participante = nome_avaliado.value.replace('Você está analisando: ','')
    
        
        if anterior == False:   
            scrp_sql = f"Select Participante, Avaliador1, Avaliador2 from QuestRH_Relacoes where Participante = '{participante}'"
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            cursor.execute(scrp_sql)
            consulta = cursor.fetchone()
            cursor.close()
            avaliadores.append(consulta['Participante'])
            avaliadores.append(consulta['Avaliador1'])
            avaliadores.append(consulta['Avaliador2'])
        else:
            scrp_sql = f"""
SELECT DISTINCT Nome_Avaliador, ID_Rel 
    FROM QuestRH_Resp_BKP 
    WHERE Participante = '{participante}' 
    AND data_bkp = (SELECT MAX(data_bkp) FROM QuestRH_Resp_BKP WHERE Participante = '{participante}') order by ID_Rel;
"""
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            cursor.execute(scrp_sql)
            consulta = cursor.fetchall()
            cursor.close()
            
            if consulta:
                for linha in consulta:
                    avaliadores.append({'nome':linha['Nome_Avaliador'], 'id':linha['ID_Rel']})
            else:
                enviar_msg (f'Não há dados do {participante} no ciclo passado',ft.Colors.RED_600)
                return
            
        texto_html = f"""
<html>
<head>
<style>
    @page {{ size: A4; margin: 1cm; }}
    body {{ font-family: Helvetica, Arial, sans-serif; color: #333; }}
    
    /* Cabeçalho conforme o PDF anexo */
    .header-table {{ width: 100%; border-bottom: 2px solid #444; margin-bottom: 20px; }}
    .header-title {{ font-size: 18pt; font-weight: bold; padding-bottom: 5px; }}
    .info-text {{ font-size: 10pt; padding: 2px 0; }}

    /* Pilar (Faixa escura) [cite: 4] */
    .pilar-header {{ 
        background-color: #0070C0; color: white; padding: 8px; 
        font-size: 11pt; font-weight: bold; margin-top: 15px;
    }}

    /* Competência (Subtítulo) [cite: 5, 7] */
    .competencia-title {{ 
        background-color: #D5EDFF; font-size: 10pt; font-weight: bold; color: #0070C0; 
        padding: 10px 0 5px 5px; border-bottom: 1px solid #eee;
    }}
    
    .quebra-pagina {{page-break-before: always;}}

    /* Tabela de Perguntas e Respostas  */
    .tabela-dados {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; }}
    .td-pergunta {{ width: 50%; font-size: 9pt; padding: 8px 5px; border-bottom: 0.5pt solid #f0f0f0; }}
    .td-resposta {{ width: 50%; padding: 5px; border-bottom: 0.5pt solid #f0f0f0; }}

    /* Estilização das Notas (Cores da sua imagem) */
    .nota-box {{ padding: 6px; border-radius: 3px; font-size: 8.5pt; font-weight: bold; }}
    
    /* Cores baseadas no esquema Flet solicitado */
    .excelente {{ background-color: #66BB6A; color: #ffffff; }} /* GREEN_400 */
    .bom {{ background-color: #DCEDC8; color: #33691E; }}       /* LIGHT_GREEN_100 */
    .mediano {{ background-color: #FFF9C4; color: #F57F17; }}   /* YELLOW_100 */
    .regular {{ background-color: #FFE0B2; color: #E65100; }}    /* ORANGE_100 */
    .ruim {{ background-color: #FFCDD2; color: #B71C1C; }}       /* RED_100 */
    .nao-apurado {{ background-color: #F5F5F5; color: #9E9E9E; border: 1px solid #E0E0E0; }} /* GREY_100 */
    
    
    .observacoes-texto {{
            font-size: 9pt;
            padding: 5px;
            color: #555;
            text-align: justify;
            line-height: 1.4;
        }}
    
    /* Classe para a linha da Média Final unificada */
    .media-final-full {{
        margin-top: 30px;
        padding: 15px 20px;
        text-align: center;
        font-size: 18pt; /* Tamanho de fonte maior e igual para toda a linha */
        font-weight: bold;
        text-transform: uppercase;
        border-radius: 5px; /* Cantos levemente arredondados para manter o padrão das labels */
        page-break-inside: avoid;
    }}
    
    /* Classes para a linha da Média Final (seguindo o mesmo padrão) */
    .media-full-excelente {{ background-color: #66BB6A; color: #ffffff; border: 1px solid #4CAF50; }}
    .media-full-bom {{ background-color: #DCEDC8; color: #33691E; border: 1px solid #C5E1A5; }}
    .media-full-mediano {{ background-color: #FFF9C4; color: #F57F17; border: 1px solid #FFF59D; }}
    .media-full-regular {{ background-color: #FFE0B2; color: #E65100; border: 1px solid #FFCC80; }}
    .media-full-ruim {{ background-color: #FFCDD2; color: #B71C1C; border: 1px solid #EF9A9A; }}
    .media-full-cinza {{ background-color: #F5F5F5; color: #9E9E9E; border: 1px solid #E0E0E0; }}    
      
    .relatorio-header {{
        display: flex;
        align-items: center;
        gap: 15px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }}

    .relatorio-header h1 {{
        color: #333;
        margin: 0;
        font-size: 24px;
    }}

    .badge {{
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: bold;
        text-transform: uppercase;
    }}

    .badge-anterior {{
        background-color: #e3f2fd;
        color: #1976d2;
        border: 1px solid #1976d2;
    }}

    .badge-atual {{
        background-color: #e8f5e9;
        color: #2e7d32;
        border: 1px solid #2e7d32;
    }}  
        
</style>
</head>
<body>
            """
        for i, avaliador in  enumerate(avaliadores):
            
            if anterior == True:
                scrp_sql = f"Select ID_Pergunta, Pilar, Competencia, Pergunta, Resposta, Desempenho_tecnico, Observacao, obs_e, obs_p, obs_c"
                scrp_sql +=  f" from QuestRH_Resp_BKP where Participante = '{participante}' and Nome_Avaliador = '{avaliador['nome']}' "
                scrp_sql += f"AND data_bkp = (SELECT MAX(data_bkp) FROM QuestRH_Resp_BKP WHERE Participante = '{participante}'and Nome_Avaliador = '{avaliador['nome']}')"
            elif avaliador == avaliador_pendente:
                scrp_sql = f"Select ID_Pergunta, Pilar, Competencia, Pergunta, Resposta, Desempenho_tecnico, Observacao, obs_e, obs_p, obs_c"
                scrp_sql +=  f" from QuestRH_Rascunho where Participante = '{participante}' and Nome_Avaliador = '{avaliador}'"
            else:
                scrp_sql = f"Select ID_Pergunta, Pilar, Competencia, Pergunta, Resposta, Desempenho_tecnico, Observacao, obs_e, obs_p, obs_c"
                scrp_sql +=  f" from QuestRH_Respostas where Participante = '{participante}' and Nome_Avaliador = '{avaliador}'"
            
               
            
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            cursor.execute(scrp_sql)
            consulta = cursor.fetchall()
            cursor.close()
            
            if anterior == False:
                scrp_sql = f'Select ROUND(AVG(Resposta)) as "média" from QuestRH_Respostas where Participante = "{participante}" and Nome_Avaliador = "{avaliador}"'
            else:
                scrp_sql= f"""
    SELECT ROUND(AVG(Resposta)) AS média 
    FROM QuestRH_Resp_BKP 
    WHERE Participante = "{participante}" 
      AND Nome_Avaliador = "{avaliador['nome']}"
      AND data_bkp = (
          SELECT MAX(data_bkp) 
          FROM QuestRH_Resp_BKP 
          WHERE Participante = "{participante}" 
            AND Nome_Avaliador = "{avaliador['nome']}"
      )
"""
            
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            cursor.execute(scrp_sql)
            média_resultado= cursor.fetchone()
            cursor.close()
            
            if média_resultado and backup == False and anterior == False:
                nota = média_resultado['média']
                if nota is None:
                    nota = 'N/I'
            elif anterior == False:
                if avaliador == avaliador_pendente:
                    nota ='Rascunho'
                else:
                    nota = 'N/I'
            elif média_resultado and anterior == True:
                nota = média_resultado['média']
                if nota is None:
                    nota = 'N/I'
            else:
                nota = 'N/I'
                
            if i>0:
                texto_html += '<div class="quebra-pagina"></div>'
                
            if anterior == False:    
                texto_html += gera_cabecalho_html(participante, avaliador, nota, id_avaliador = i)
                if consulta:
                    texto_html += gera_conteudo_html(consulta,participante, avaliador)
                else:
                    texto_html += '<div style:"alingment: center, font-weight: bold", font-size: 10pt, padding=5px 8px>Sem informações do avaliador</div>'
            else:
                texto_html += gera_cabecalho_html(participante, avaliador['nome'], nota, id_avaliador = avaliador['id'], anterior=anterior)
                if consulta:
                    texto_html += gera_conteudo_html(consulta,participante, avaliador['nome'])
                else:
                    texto_html += '<div style:"alingment: center, font-weight: bold", font-size: 10pt, padding=5px 8px>Sem informações do avaliador</div>'
           
        cores_linha = {
                    0: 'media-full-cinza',
                    1: 'media-full-ruim', 
                    2: 'media-full-regular', 
                    3: 'media-full-mediano', 
                    4: 'media-full-bom', 
                    5: 'media-full-excelente'
                }
        
        if anterior == False:
            # Validação e tratamento do valor
            valor_limpo = str(txt_media_final.value).replace(',', '.')

            try:
                media_num = float(valor_limpo)
                nota_inteira = int(media_num) 
                
                # Busca a cor ou define 'media-full-mediano' como padrão
                classe_aplicada = cores_linha.get(nota_inteira, 'media-full-mediano')
                
                texto_html += f"""
                <div class="media-final-full {classe_aplicada}">
                    MÉDIA FINAL: {valor_limpo}
                </div>
                """
            except (ValueError, TypeError):
                # Se não for numérico, aplica a classe cinza-claro
                texto_html += f"""
                <div class="media-final-full media-full-cinza">
                    MÉDIA FINAL: NÃO APURADA
                </div>
                """
        else:
            scrp_sql = f"""
SELECT ROUND(AVG(Resposta)) as média 
FROM QuestRH_Resp_BKP 
WHERE Participante = '{participante}' 
  AND data_bkp = (SELECT MAX(data_bkp) FROM QuestRH_Resp_BKP WHERE Participante = '{participante}')
GROUP BY Participante, data_bkp
HAVING COUNT(DISTINCT CASE WHEN id_rel IN (1, 2) THEN id_rel END) = 2
"""
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            cursor.execute(scrp_sql)
            resp_média= cursor.fetchone()
            cursor.close()
            
            if resp_média:
                try:
                    nota_inteira = int(resp_média['média'])
                    # Busca a cor ou define 'media-full-mediano' como padrão
                    classe_aplicada = cores_linha.get(nota_inteira, 'media-full-mediano')
                    
                    texto_html += f"""
                    <div class="media-final-full {classe_aplicada}">
                        MÉDIA FINAL: {nota_inteira}
                    </div>
                    """
                except (ValueError, TypeError):
                    # Se não for numérico, aplica a classe cinza-claro
                    texto_html += f"""
                    <div class="media-final-full media-full-cinza">
                        MÉDIA FINAL: NÃO APURADA
                    </div>
                    """    
            
        texto_html += "</body></html>" 
        #print(texto_html)
        
        
        #with open(f"{BASE_DIR}/assets/relatorios/{login}.txt", "w+b") as f:
            #f.write(texto_html.encode("utf-8"))
        
        pdf_buffer = io.BytesIO()
        pisa_status = pisa.CreatePDF(io.BytesIO(texto_html.encode("utf-8")), dest=pdf_buffer)
        
        if not pisa_status.err:
            # 2. Nome do arquivo dinâmico
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            nome_arquivo = f"Relatorio_{login.replace('.', '_')}.pdf"
            
            caminho_file = os.path.join("assets", "relatorios",  nome_arquivo)
            
            # 2. Certifique-se que a pasta assets existe na VPS 
            if not os.path.exists(os.path.join("assets", "relatorios")):
                os.makedirs(os.path.join("assets", "relatorios"))
                
                
            # Salva fisicamente o arquivo na pasta assets
            with open(caminho_file, "wb") as f:
                f.write(pdf_buffer.getbuffer())
                
                
            caminho_url = f"/relatorios/{nome_arquivo}"
            page.launch_url(f"{caminho_url}?t={datetime.datetime.now().timestamp()}")
            # Cria uma tarefa em segundo plano para deletar daqui a 30 segundos
            def deletar_depois():
                time.sleep(30) # Tempo suficiente para o navegador completar o download
                if os.path.exists(caminho_file):
                    os.remove(caminho_file)
                    
            threading.Thread(target=deletar_depois, daemon=True).start()
            
            enviar_msg(f"Relatório gerado com sucesso para o participante {participante}!", ft.Colors.GREEN_600)
        else:
            print("Erro ao gerar PDF")
            
        #page.launch_url(f"/relatorio_{login}.pdf")    
    
    def gera_conteudo_html(dados={}, participante='', avaliador=''):
        pilar_anterior = ''
        competencia_anterior = ''
        dados_html = ''
        
        classificacao_nota= {0: 'nao-apurado',1: 'ruim', 2: 'regular', 3: 'mediano', 4: 'bom', 5: 'excelente'}
        descricao_nota = {
            0: "Não apurado",
            1: "1 - Insatisfatório - Não atende ou atende minimamente aos padrões",
            2: "2 - Regular - Atende parcialmente aos padrões esperados",
            3: "3 - Satisfatório - Atende os padrões esperados",
            4: "4 - Bom - Demonstra empenho e excelência no atendimento de padrões esperados",
            5: "5 - Excelente - Supera os padrões esperados"
        }
        
        for dado in dados:
            if dado['Pilar'] != pilar_anterior:
                if competencia_anterior != '':
                    dados_html += '</table>'
                pilar_anterior = dado['Pilar']
                dados_html += f'<div class="pilar-header">Pilar: {dado["Pilar"]}</div>'
                
            if dado['Competencia'] != competencia_anterior:
                if competencia_anterior != '':
                    dados_html += '</table>'
                competencia_anterior = dado['Competencia']
                dados_html += f'<div class="competencia-title">Competência: {dado["Competencia"]}</div>'
                dados_html += '<table class="tabela-dados">'
            
            dados_html += f"""
    <tr>
        <td class="td-pergunta">{dado['ID_Pergunta']}) {dado['Pergunta']}</td>
        <td class="td-resposta"><div class="nota-box {classificacao_nota[dado['Resposta']]}">{descricao_nota[dado['Resposta']]}</div></td>
    </tr>
                """
        
        dados_html += '</table>'
        
        dados_html += f'<div class="pilar-header">Desempenho Técnico</div>'
        dados_html += f"""
    <table class="tabela-dados">
        <tr>
            <td class="td-pergunta">Desempenho - Execução e entrega de atividades de forma geral</td>
            <td class="td-resposta"><div class="nota-box {classificacao_nota[dados[0]['Desempenho_tecnico']]}">{descricao_nota[dados[0]['Desempenho_tecnico']]}</div></td>
        </tr>
    </table>
                    """
       

        if participante != avaliador:
            dados_html += f'<div class="pilar-header">Calibração</div>'
            dados_html += f'<div class="competencia-title">Pilar E</div>'
            observacao = str(dados[0]['obs_e']).replace('\n', '<br>')
            dados_html += f'<div class="observacoes-texto">{observacao}</div>'
            
            dados_html += f'<div class="competencia-title">Pilar P</div>'
            observacao = str(dados[0]['obs_p']).replace('\n', '<br>')
            dados_html += f'<div class="observacoes-texto">{observacao}</div>'
            
            dados_html += f'<div class="competencia-title">Pilar C</div>'
            observacao = str(dados[0]['obs_c']).replace('\n', '<br>')
            dados_html += f'<div class="observacoes-texto">{observacao}</div>'
        
        dados_html += f'<div class="pilar-header">Observações</div>'
        observacao = str(dados[0]['Observacao']).replace('\n', '<br>')
        dados_html += f'<div class="observacoes-texto">{observacao}</div>'
        
        return dados_html
    
    def gera_cabecalho_html(participante, avaliador, media, id_avaliador, anterior= False):
        
        if id_avaliador == 0:
            nomenclatura = "Auto Avaliação"
        else:
            nomenclatura = f"Avaliação{id_avaliador}"
        
        classe_ciclo = 'badge-anterior' if anterior else 'badge-atual'
        texto_ciclo = 'CICLO 2025' if anterior else 'CICLO 2026'

        nome_rel = f"""
            <div class="relatorio-header">
                <h1>Relatório de Desempenho</h1>
                <span class="badge {classe_ciclo}">{texto_ciclo}</span>
            </div>
        """
        
        html_content = f"""
    <table class="header-table">
        <tr>
            <td>
                <div class="header-title">{nome_rel}</div>
                <div class="info-text"><strong>Participante:</strong> {participante}</div>
                <div class="info-text"><strong>{nomenclatura}:</strong> {avaliador}</div>
            </td>
            <td align="right" valign="bottom">
                <div class="info-text"><strong>Média:</strong></div>
                <div style="font-size: 20pt; font-weight: bold; color: #2c3e50;">{media}</div>
            </td>
        </tr>
    </table>
        """
        return html_content
    
    def exportar_para_excel(e=None):
        try:
            mostrar_alerta_temporario('Gerando arquivo de respostas, aguarde a finalização...', ft.Colors.BLUE_400)

            # 1. Conexão e consulta
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            scrp_sql = "SELECT * FROM QuestRH_Respostas"
            cursor.execute(scrp_sql)
            consulta = cursor.fetchall()
            cursor.close()
            conn.close()

            if not consulta:
                mostrar_alerta_temporario('Não existem dados para exportar', ft.Colors.ORANGE_400)
                return

            # 2. Criar DataFrame
            df = pd.DataFrame(consulta)

            # 3. Definição de caminhos seguros no servidor/máquina local
            # Certifica que a pasta assets/relatorios existe
            pasta_exportacao = os.path.join("assets", "relatorios")
            if not os.path.exists(pasta_exportacao):
                os.makedirs(pasta_exportacao)

            # Nome único baseado no timestamp
            nome_arquivo = f"exportacao_respostas_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            caminho_file = os.path.join(pasta_exportacao, nome_arquivo)

            # 4. Salvar o arquivo fisicamente no servidor usando o openpyxl
            with pd.ExcelWriter(caminho_file, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            # 5. Criar a URL relativa que o Flet expõe publicamente
            # O Flet mapeia a pasta 'assets' como a raiz '/', então 'assets/relatorios/arq.xlsx' vira '/relatorios/arq.xlsx'
            caminho_url = f"/relatorios/{nome_arquivo}"
            
            # Abre o link real e limpo no navegador do usuário (evita bloqueio de segurança)
            page.launch_url(f"{caminho_url}?t={datetime.datetime.now().timestamp()}")

            # 6. Criar uma tarefa em segundo plano para deletar o arquivo após 45 segundos
            # Isso impede que o seu servidor fique lotado de arquivos antigos de exportação
            def deletar_depois():
                time.sleep(45) # Tempo suficiente para o usuário terminar o download
                if os.path.exists(caminho_file):
                    os.remove(caminho_file)
                    
            threading.Thread(target=deletar_depois, daemon=True).start()

            mostrar_alerta_temporario('Exportação realizada com sucesso', ft.Colors.GREEN_400)
            
        except Exception as ex:
            mostrar_alerta_temporario(f'Erro ao exportar: {ex}', ft.Colors.RED_400)


    def escolher_pasta(e):
        file_picker.get_directory_path()

    def ao_escolher_pasta(e: ft.FilePickerResultEvent):
        if e.path:
            exportar_para_excel(e.path)

    ####################################################################################################################################### 
    ######################################################## Criando Views ################################################################ 
    ####################################################################################################################################### 

     # Overlay de carregamento
    aguarde_overlay = ft.Container(
        content=ft.Column(
            [
                ft.ProgressRing(),
                mensagem_aguarde
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=20
        ),
        alignment=ft.alignment.center,
        bgcolor=ft.Colors.with_opacity(0.8, ft.Colors.BLACK),
        width=float("inf"),
        height=float("inf"),
        visible=False
    )

    # Overlay de confirmação
    confirmacao_overlay = ft.Container(
        content=ft.Column(
            [
                ft.Text(
                    "Deseja realmente continuar com o envio da resposta dessa pessoa? Uma vez realizado, não é possível rever e reenviar o relatório",
                    size=18,
                    weight=ft.FontWeight.BOLD,
                    text_align=ft.TextAlign.CENTER,
                    color=ft.Colors.WHITE
                ),
                ft.Row(
                    [
                        ft.ElevatedButton(
                            "Sim",
                            on_click=lambda e: confirmar_continuacao(e, True),
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.BLUE,
                                color=ft.Colors.WHITE
                            )
                        ),
                        ft.ElevatedButton(
                            "Não",
                            on_click=lambda e: confirmar_continuacao(e, False),
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.RED,
                                color=ft.Colors.WHITE
                            )
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    spacing=20
                )
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=30
        ),
        alignment=ft.alignment.center,
        bgcolor=ft.Colors.with_opacity(0.85, ft.Colors.BLACK),
        width=float("inf"),
        height=float("inf"),
        visible=False
    )
    
    ###################################################### TELA DE DESLOG #####################################################
    
    
    deslog_overlay = ft.Container(
        content=ft.Column(
            [
                msg_deslog,
                ft.Row(
                    [
                        ft.ElevatedButton(
                            "Sim",
                            on_click=lambda e: reset_idle_time(e),
                            style=ft.ButtonStyle(
                                bgcolor=ft.Colors.BLUE,
                                color=ft.Colors.WHITE
                            )
                        )
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    spacing=20
                )
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=30
        ),
        alignment=ft.alignment.center,
        bgcolor=ft.Colors.with_opacity(0.85, ft.Colors.BLACK),
        width=float("inf"),
        height=float("inf"),
        visible=False
    )
    
    page.overlay.append(aguarde_overlay)
    page.overlay.append(confirmacao_overlay)
    page.overlay.append(deslog_overlay)
   
    #############################################################################################################################

     # --- Animação de Entrada ---
    def animate_entrance():
        time.sleep(0.1)
        left_side.opacity = 1
        left_side.offset = ft.Offset(0, 0)
        right_side.opacity = 1
        right_side.offset = ft.Offset(0, 0)
        page.update()

            
    # --- LADO ESQUERDO: Layout com Logo no topo e GIF na base ---
    left_side = ft.Container(
        expand=True,
        opacity=0,
        bgcolor="#fcfcfc",
        offset=ft.Offset(-0.05, 0),
        animate_opacity=ft.Animation(1000, ft.AnimationCurve.EASE_OUT),
        animate_offset=ft.Animation(1000, ft.AnimationCurve.EASE_OUT_QUINT),
        padding=ft.padding.only(top=40, bottom=20, left=40, right=40),
        content=ft.Column(
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            alignment=ft.MainAxisAlignment.CENTER, # Empurra o conteúdo para as extremidades
            controls=[
                
                # Conteúdo Central (Texto)
                ft.Column(
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    controls=[
                        ft.Image(
                        src='/Enind Grupo - Vetor.svg',
                        width=220,
                        height=220,
                        fit=ft.ImageFit.CONTAIN,
                        ),
                        ft.Divider(height=40, color=ft.Colors.TRANSPARENT),
                        ft.Text("Avaliação de Desempenho", size=32, weight="bold", color="#1a237e"),
                        ft.Container(height=10),
                        ft.Text(
                            "O jeito como fazemos as coisas por aqui é a nossa cultura",
                            size=15, 
                            color=ft.Colors.GREY_700, 
                            text_align="center", 
                            italic=True
                        ),
                    ]
                ),
                ft.Divider(height=40, color=ft.Colors.TRANSPARENT),
                # Seu GIF de evolução abaixo de tudo
                ft.Image(
                    src='/evolucao.gif',  # Nome do seu arquivo GIF
                    height=60,
                    width=426
                ),
            ],
        ),
    )

    # --- LADO DIREITO: Login ---
    right_side = ft.Container(
        expand=True,
        bgcolor="#F7F9FC",
        opacity=0,
        offset=ft.Offset(0.05, 0),
        animate_opacity=ft.Animation(1000, ft.AnimationCurve.EASE_OUT),
        animate_offset=ft.Animation(1000, ft.AnimationCurve.EASE_OUT_QUINT),
        content=ft.Column(
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                ft.Text("Login", size=40, weight="black", color="#1a237e"),
                ft.Text("Identifique-se para continuar", color=ft.Colors.GREY_600),
                ft.Container(height=30),
                nome_cb,
                ft.Container(height=10),
                senha_txt,
                ft.Container(height=30),
                ft.Container(
                    content=ft.Text("ENTRAR NO SISTEMA", color="white", weight="bold"),
                    bgcolor="#1a237e",
                    width=380,
                    height=50,
                    border_radius=12,
                    alignment=ft.alignment.center,
                    on_click=validar_login,
                    on_hover=lambda e: setattr(e.control, "bgcolor", "#283593" if e.data == "true" else "#1a237e"),
                    animate=200,
                ),
            ],
        ),
    )


    login_view = ft.Row(
            expand=True,
            spacing=0,
            controls=[left_side, right_side]
        )


    file_picker.on_result = ao_escolher_pasta
    
    exportar_material_btn = ft.ElevatedButton(
        "Visualizar Material de Apoio",
        on_click=escolher_pasta,
        width=250,
        height=50
    )

    exportar_pdf_btn = ft.ElevatedButton(
        "Exportar para PDF",
        icon= ft.Icons.DOWNLOAD_ROUNDED,
        on_click=lambda _: preparar_pdf(adm=True),
        width=250,
        height=50
    )
    
    exportar_pdf_btn_antes = ft.ElevatedButton(
        "PDF Ciclo Passado",
        icon= ft.Icons.CLOUD_DOWNLOAD_ROUNDED,
        on_click=lambda _: preparar_pdf(anterior=True),
        width=250,
        height=50
    )
    
    
    btn_excel_av1 = ft.ElevatedButton(
        "Exportar Avaliador 1",
        icon= ft.Icons.CLOUD_DOWNLOAD_ROUNDED,
        width=250,
        height=50
    )
    
    
    exportar_pdf_btn_antes_adm = ft.ElevatedButton(
        "PDF Ciclo Passado",
        icon= ft.Icons.CLOUD_DOWNLOAD_ROUNDED,
        on_click=lambda _: preparar_pdf(anterior=True, adm=True),
        width=250,
        height=50
    )
    
    exportar_pdf_btn_bkp = ft.ElevatedButton(
        "Exportar para PDF",
        icon= ft.Icons.DOWNLOAD_ROUNDED,
        on_click=lambda _: preparar_pdf(None, True),
        width=250,
        height=50,
        visible = False
    )
    
    
    exportar_btn = ft.ElevatedButton(
        "Respostas",
        icon= ft.Icons.DOWNLOAD_ROUNDED,
        on_click=lambda _: exportar_para_excel(),
        width=150,
        height=35
    )

    exportar_grafico_btn = ft.ElevatedButton(
        "Exportar para Excel",
        icon= ft.Icons.DOWNLOAD_ROUNDED,
        on_click=lambda _: exportar_grafico_excel(),
        width=250,
        height=35
    )

    grafico_btn = ft.ElevatedButton(
        "Gráficos",
        icon=ft.Icons.BAR_CHART_ROUNDED,
        on_click=lambda _: inicializa_grafico(),
        width=150,
        height=35
    )

    grafico_status_btn = ft.ElevatedButton(
        "Status",
        icon=ft.Icons.SPEED_OUTLINED,
        on_click=lambda e: atualizar_painel_status(e),
        width=150,
        height=35
    )

    atualizar_btn = ft.ElevatedButton(
        content=ft.Row(
                [
                    ft.Icon(name=ft.Icons.REFRESH, size=20),
                    ft.Text("Atualizar")
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=10
            ),
            on_click=atualiza_rel,
            width=150,
            height=35
        )

    # Cabeçalho fixo
    cabecalho = ft.Container(
        content=ft.Row([
            ft.Container(ft.Text("Foto", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Nome", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Auto Avaliação", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação 1", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação 2", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Status", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text('Avaliar', weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text('Visualizar', weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text('Feedback', weight=ft.FontWeight.BOLD), expand=1),
        ],alignment=ft.MainAxisAlignment.CENTER),
        padding=10,
        bgcolor=ft.Colors.BLUE_100,
        border_radius=ft.border_radius.only(top_left=10, top_right=10),
        
    )


    # Corpo com rolagem (apenas as linhas rolam)
    corpo_tabela = ft.Container(
        content=lista_view,
        height = page.height * 0.7,
        bgcolor=ft.Colors.translate,
        border_radius=ft.border_radius.only(bottom_left=10, bottom_right=10)
    )


    painel_view = ft.Container(
        content=ft.Column([
            ft.Row([ft.TextButton("Deslogar", icon=ft.Icons.ARROW_BACK, on_click=voltar_login), texto_ola1], spacing= 10),
            container_expiração,
            ft.Row([ft.Row([ft.Text("Painel de Controle", size=25, weight=ft.FontWeight.BOLD), 
                            ft.Row([atualizar_btn, btn_excel_av1], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)]),btn_ver_manual],
                   expand=True, alignment=ft.MainAxisAlignment.SPACE_BETWEEN,vertical_alignment=ft.CrossAxisAlignment.CENTER),
           cabecalho,
           corpo_tabela
        ]),
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

      # Cabeçalho fixo
    cabecalho_pend = ft.Container(
        content=ft.Row([
            ft.Container(ft.Text("Foto", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Participante", weight=ft.FontWeight.BOLD), expand=2),
            ft.Container(ft.Text("Status", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliador1", weight=ft.FontWeight.BOLD), expand=2),
            ft.Container(ft.Text("Status", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliador2", weight=ft.FontWeight.BOLD), expand=2),
            ft.Container(ft.Text("Status", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text('Feedback', weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text('Visualizar', weight=ft.FontWeight.BOLD), expand=1),
            
        ],alignment=ft.MainAxisAlignment.CENTER),
        padding=10,
        bgcolor=ft.Colors.BLUE_100,
        border_radius=ft.border_radius.only(top_left=10, top_right=10),
        
    )

    # Corpo com rolagem (apenas as linhas rolam)
    corpo_tabela_pend = ft.Container(
        content=lista_pend_view,
        expand=True,
        bgcolor=ft.Colors.translate,
        border_radius=ft.border_radius.only(bottom_left=10, bottom_right=10),
        height = page.height * 0.7
    )

    btn_exportar_tabela = ft.ElevatedButton('Pendências', on_click=lambda e:exportar_dados_painel_adm(e), width=150, height=35, icon=ft.Icons.DOWNLOAD_ROUNDED)
    btn_reiniciar_ciclo  = ft.ElevatedButton(
            'Reiniciar Ciclo', 
            on_click=lambda e:reiniciar_ciclo(e), 
            width=150,
            height=35, 
            icon=ft.Icons.CLOUD_UPLOAD, visible = False)
    
    btn_restaurar = ft.ElevatedButton(
        "Restaurar Dados",
        icon=ft.Icons.SETTINGS_BACKUP_RESTORE,
        width=150,
        height=35, 
        on_click=restaurar_backup_dinamico,
        visible= False
    )
    
    baixar_rel_final = ft.ElevatedButton(
        "Relatório Final",
        icon=ft.Icons.DOWNLOAD_FOR_OFFLINE,
        width=150,
        height=35, 
        on_click=lambda e:preparar_relatório_final(e),
        visible= False
    )
    
    baixar_rel_final_av1 = ft.ElevatedButton(
        "Relatório Av1",
        icon=ft.Icons.DOWNLOAD_FOR_OFFLINE,
        width=150,
        height=35, 
        on_click=lambda e:preparar_relatório_final(e, True),
        visible= False
    )
    
    painel_pend_view = ft.Container(
        content=ft.Column([
            ft.Row([ft.TextButton("Deslogar", icon=ft.Icons.ARROW_BACK, on_click=voltar_login),texto_ola2], spacing= 10),
            container_expiração,
            ft.Text("Painel de Controle", size=25, weight=ft.FontWeight.BOLD),
            ft.Container(ft.Row([btn_reiniciar_ciclo,btn_restaurar,baixar_rel_final, baixar_rel_final_av1],spacing=10), expand=True),
            ft.Container(ft.Row([atualizar_btn,exportar_btn, grafico_btn, grafico_status_btn, btn_exportar_tabela, chk_estrategico, chk_nao_estrategico],spacing=10), expand=True),
            cabecalho_pend,
            corpo_tabela_pend
        ]),
        
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

    #form_content = ft.Column(spacing=10,expand=True)

    form_content = ft.ListView(
        spacing=10,
        expand=2,
        auto_scroll=False,  # Mantenha a rolagem manual
        padding=0
        
    )

    container_perguntas = ft.Container(
        content=ft.Row([form_content, container_observacoes]),
        height=page.height * 0.7,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=True
    )

    txt_ociosidade = ft.Text("", size=10, weight=ft.FontWeight.BOLD)
    formulario_view =  ft.Container(
            content=ft.Column([
                ft.Row([
                        ft.Row([ ft.TextButton("Voltar", on_click=voltar_painel, icon=ft.Icons.ARROW_BACK),texto_ola3]),
                        ft.Row([exportar_pdf_btn_bkp, exportar_pdf_btn_antes], spacing=10),],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                ft.Row([nome_em_avaliacao,txt_ociosidade],spacing=10, alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                container_perguntas,
                alerta_container_form,
                ft.Row(
                    controls=[
                        ft.ElevatedButton(
                            "Enviar",
                            on_click=mostrar_confirmacao,
                            width=300,
                            height=50,
                            style=ft.ButtonStyle(
                                shape=ft.RoundedRectangleBorder(radius=10),
                                bgcolor=ft.Colors.BLUE_600,
                                color=ft.Colors.WHITE
                            )
                        )
                    ],
                    alignment=ft.MainAxisAlignment.CENTER
                )
            ], spacing=15),
            bgcolor=ft.Colors.WHITE,
            border_radius=10,
            padding=20,
            opacity=0.92,
        )
    
    Gestor_tempo_formulario = ft.GestureDetector(
        on_pan_start= reset_idle_time,     # movimento de mouse ou dedo
        on_pan_update= reset_idle_time,    # movimento contínuo
        on_tap= reset_idle_time,
        on_hover=reset_idle_time,
        content= formulario_view,
        visible=False
    )
   
    # Cabeçalho fixo
    cabecalho_respostas = ft.Container(
        content=ft.Row([
            ft.Container(ft.Text("Avaliador", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Pilar", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Competência", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Pergunta", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Avaliação", weight=ft.FontWeight.BOLD), expand=2),
        ],alignment=ft.MainAxisAlignment.CENTER),
        padding=10,
        bgcolor=ft.Colors.BLUE_100,
        border_radius=ft.border_radius.only(top_left=10, top_right=10),
        
    )

    container_respostas_auto = ft.Container(
        content=ft.Row([ft.Container(lista_reultados_auto_view,expand=2),
                           ft.Container(container_obs_auto, expand=1)]),
        height= page.height * 0.5,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )
    container_respostas_av1 = ft.Container(
        content=ft.Row([ft.Container(lista_reultados_av1_view,expand=2),
                           ft.Container(container_obs_av1, expand=1)]),
        height= page.height * 0.5,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

    container_respostas_av2 = ft.Container(
        content=ft.Row([ft.Container(lista_reultados_av2_view,expand=2),
                           ft.Container(container_obs_av2, expand=1)]),
        height= page.height * 0.5,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

    painel_resposta_view = ft.Container(
        content=ft.Column([
            ft.Row([
                    ft.Row([ft.TextButton("Voltar", on_click=voltar_painel, icon=ft.Icons.ARROW_BACK),texto_ola4]),
                    ft.Row([exportar_pdf_btn, exportar_pdf_btn_antes_adm], spacing=10)
                ],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            ft.Row([nome_avaliado, container_media_final], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            container_txt_auto,
            container_respostas_auto,
            container_txt_av1,
            container_respostas_av1,
            container_txt_av2,
            container_respostas_av2
        ], spacing=5),
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

    formulario_Envio = ft.Container(
        content=ft.Column([
            ft.Icon(name=ft.Icons.SENTIMENT_SATISFIED_ALT, size=80, color=ft.Colors.GREEN_600),
            ft.Text(
                "Muito obrigado por sua participação!",
                size=25,
                weight=ft.FontWeight.BOLD,
                color=ft.Colors.GREEN_800,
                text_align=ft.TextAlign.CENTER
            ),
            ft.Text(
                f"Suas respostas referente ao participante abaixo foram registradas com sucesso.\nSua colaboração é fundamental para o nosso crescimento.",
                size=16,
                text_align=ft.TextAlign.CENTER,
                color=ft.Colors.GREY_800
            ),
            participante_realizado,
            ft.Row(
                controls=[
                    ft.ElevatedButton(
                        "Voltar ao Painel",
                        on_click=lambda e:voltar_painel(atualizar=True),
                        width=300,
                        height=50,
                        style=ft.ButtonStyle(
                            shape=ft.RoundedRectangleBorder(radius=10),
                            bgcolor=ft.Colors.BLUE_600,
                            color=ft.Colors.WHITE
                        )
                    )
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=10
            )
        ],
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=20),
        bgcolor=ft.Colors.WHITE,
        border_radius=12,
        padding=30,
        opacity=0.95,
        visible=False
    )

    # Cabeçalho fixo
    cabecalho_grafico = ft.Container(
        content=ft.Row([
            ft.Container(ft.Text("Participante", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Tipo_av", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Sigla_Emp", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("C_Custo", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Cargo", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Media_auto", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Media_Avs", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Desemp_auto", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Desemp_Avs", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Performance", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Status_Av", weight=ft.FontWeight.BOLD), expand=1)
        ],alignment=ft.MainAxisAlignment.CENTER),
        padding=10,
        bgcolor=ft.Colors.BLUE_100,
        border_radius=ft.border_radius.only(top_left=10, top_right=10),
    )

    lista_grafico = ft.ListView(expand=True, auto_scroll=False, spacing=10, padding=10)

    container_grafico = ft.Container(
        content=lista_grafico,
        height= page.height * 0.7,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=True
    )


     # Cabeçalho fixo
    cabecalho_dados_bellcurve = ft.Container(
        content=ft.Row([
            ft.Container(ft.Text("Participante", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Média Final", weight=ft.FontWeight.BOLD), expand=1)
        ],alignment=ft.MainAxisAlignment.CENTER),
        padding=10,
        border_radius=ft.border_radius.only(top_left=10, top_right=10),
    )

    lista_grafico_bellcurve = ft.ListView(expand=True, auto_scroll=False, spacing=10, padding=10)

    dropdown_bellcurve_nota = ft.Dropdown(
        label="Notas",
        options=[ft.dropdown.Option("1"), ft.dropdown.Option("2"), ft.dropdown.Option("3"), ft.dropdown.Option("4"), ft.dropdown.Option("5")],
        on_change=lambda e:atualiza_tabela_bellcurve(e),
        expand=True
    )

    container_dados_bellcurve = ft.Container(
        content=lista_grafico_bellcurve,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=True,
        height=page.height * 0.4
    )

    btn_limpar_filtros_bellcurve = ft.TextButton(
        "Limpar Filtros",
        on_click=lambda e:limpar_filtros_bellcurve(e),
        style=ft.ButtonStyle(
            shape=ft.RoundedRectangleBorder(radius=10),
            bgcolor=ft.Colors.BLUE_600,
            color=ft.Colors.WHITE
        )
    )

    # Cabeçalho fixo
    container_bellcurve = ft.Column([
        ft.Row([dropdown_bellcurve_nota, btn_limpar_filtros_bellcurve],spacing=10),
        cabecalho_dados_bellcurve, container_dados_bellcurve],
        expand=3
    )

    chk_estrategico_grafico = ft.Checkbox(label="Estratégico", value=True, on_change=lambda e: atualizar_query(e))
    chk_nao_estrategico_grafico = ft.Checkbox(label="Não Estratégico", value=True, on_change=lambda e: atualizar_query(e))

    cboxPessoa.on_change = atualiza_dados
    cboxPilar.on_change = atualiza_dados
    cboxCompetencia.on_change = atualiza_dados
    dropdown_Performance.on_change = atualiza_dados
    
    # Estilo comum para os containers de imagem (Cards)
    
    container_grafico = ft.Column(
            [   ft.Row([ft.TextButton("Voltar", icon=ft.Icons.ARROW_BACK, on_click=voltar_painel), texto_ola5], spacing= 10),
                ft.Row([ft.Text("📊 Análise Gráfica - Avaliação de Desempenho", size=20, weight="bold"), exportar_grafico_btn], spacing= 10),
                ft.Divider(),
                ft.Row(
                    [
                        ft.Container(cboxPessoa,expand=3),
                        ft.Container(cboxPilar,expand=2),
                        ft.Container(cboxCompetencia,expand=2),
                        ft.Container(dropdown_Performance, expand=2),
                        ft.Container(ft.Column([chk_estrategico_grafico,chk_nao_estrategico_grafico]), expand=1),
                        btn_limpar_campos
                    ],
                    spacing=5
                ),
                botoes,
                ft.Divider(),
                ft.Row([ft.Container(img_Ninebox,expand=2), ft.Container(img_Pilar, expand=3) ], alignment=ft.MainAxisAlignment.CENTER, spacing=5),
                ft.Container(img_Comp),
                ft.Container(img_Compar),
                ft.Row([ft.Container(img_BellCurve,expand=3),container_bellcurve], alignment='top', spacing=5),
                cabecalho_grafico,
                container_grafico
            ],
            expand=True,
            visible=False
        )
    
    txt_assinatura = ft.Text('Powered by ENIND Engenharia - Wagner Barreiro all rights reserved', size=10, color=ft.Colors.GREY_600)

    # containers
    container_realizados = ft.Container(
        content=ft.Column([
            txt_Realizados,
            ft.Text('Avaliações Realizadas',weight='w200', size=15, color=ft.Colors.GREEN)
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
            ft.Text('Avaliações Pendentes',weight='w200', size=15, color=ft.Colors.RED)
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
            ft.Text('Participantes Finalizados',weight='w200', size=15, color=ft.Colors.BLUE)
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
            ft.Text('Participantes Pendentes',weight='w200', size=15, color=ft.Colors.AMBER),
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
        # --- CABEÇALHO ---
        ft.Row([
            ft.TextButton("Voltar", icon=ft.Icons.ARROW_BACK, on_click=voltar_painel),
            ft.VerticalDivider(width=1),
            texto_ola6
        ], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.CENTER),
        
        ft.Row([
            ft.Text("   📊 Análise Gráfica - Status das Avaliações", size=24, weight="bold", color="blue900")
        ]),
        
        ft.Divider(height=10, color="transparent"),
        ft.Container(container_texto_grafico, padding=ft.padding.only(left=15)),

        # --- CORPO DO DASHBOARD ---
        ft.Row([
            
            # COLUNA DA ESQUERDA (Status e Potencial)
            ft.Column([
                # Linha 1: Concluídos e Não Finalizados
                ft.Row([
                    ft.Container(img_Concluidos, expand=1, bgcolor="white", border_radius=10, padding=10, shadow=ft.BoxShadow(blur_radius=5, color="black12")),
                    ft.Container(img_Potencial_nao_finalizados, expand=2, bgcolor="white", border_radius=10, padding=10, shadow=ft.BoxShadow(blur_radius=5, color="black12")),
                ], spacing=15),
                
                # Linha 2: Finalizados e Potencial
                ft.Row([
                    ft.Container(img_Finalizados, expand=1, bgcolor="white", border_radius=10, padding=10, shadow=ft.BoxShadow(blur_radius=5, color="black12")),
                    ft.Container(img_Potencial, expand=2, bgcolor="white", border_radius=10, padding=10, shadow=ft.BoxShadow(blur_radius=5, color="black12")),
                ], spacing=15),
            ], expand=3, spacing=15),

            # COLUNA DA DIREITA (Barra Empresa + Feedbacks abaixo)
            ft.Column([
                # Gráfico de Barras Principal (Topo)
                ft.Container(
                    content=img_Barra_Empresa,
                    expand=2, # Dá mais espaço vertical para as barras
                    bgcolor="white",
                    border_radius=10,
                    padding=20,
                    alignment=ft.alignment.center,
                    shadow=ft.BoxShadow(blur_radius=8, color="black12")
                ),
                
                # Linha de Feedbacks (Base da coluna direita)
                ft.Row([
                    ft.Container(img_feedback_antes, expand=1, bgcolor="white", border_radius=10, padding=10, shadow=ft.BoxShadow(blur_radius=5, color="black12")),
                    ft.Container(img_feedback_atual, expand=1, bgcolor="white", border_radius=10, padding=10, shadow=ft.BoxShadow(blur_radius=5, color="black12")),
                ], spacing=15, expand=1) # expand=1 faz essa linha ser menor que a de cima
                
            ], expand=2, spacing=15),
            
        ], expand=True, spacing=20)

    ], visible=False, expand=True, spacing=10)

    
    #with open(image_path, "rb") as img_file:
            #img_base64 = base64.b64encode(img_file.read()).decode('utf-8')
    
    '''
        content=ft.Image(
            src='/Imagem_Quest.svg',
            fit=ft.ImageFit.COVER,
            expand=True,
            gapless_playback=True
        ),
        '''
    imagem_fundo = ft.Container(
        content=ft.Text(''),
        alignment=ft.alignment.center,
        bgcolor=ft.Colors.TRANSPARENT,
        expand=True
    )

    painel_comum = ft.Container(
        content=ft.Column([
            login_view,
            painel_view,
            Gestor_tempo_formulario,
            formulario_Envio,
        ], alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=30),
        alignment=ft.alignment.center,
    )

    painel_av1 = ft.Container(
        content=ft.Column([
            login_view,
            painel_view,
            painel_resposta_view,
            Gestor_tempo_formulario,
            formulario_Envio,
        ], alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=30),
        alignment=ft.alignment.center,
    )

    painel_admin = ft.Container(
        content=ft.Column([
            login_view,
            painel_pend_view,
            painel_resposta_view,
            container_grafico,
            container_painel_grafico
        ], alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=30),
        alignment=ft.alignment.center,
    )
    
    conteudo_central = ft.Container(
        content=ft.Column([
            login_view,
        ]),
    )
    
    overlay = ft.Column([
        alerta_container,
        login_view,
        ft.Row([txt_assinatura], alignment=ft.MainAxisAlignment.CENTER),
        ], 
    expand=True,
    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
    horizontal_alignment=ft.CrossAxisAlignment.CENTER
    )

    stack = ft.Stack([
        #imagem_fundo,
        overlay
    ])


    #print('✔️ tamanho da tela: ', page.width, 'x', page.height)
    # Adicionando os dois lados
    page.add(
        overlay 
    )

    page.update()
    animate_entrance()
    
    page.on_resized = atualizar_altura_container
    
#ft.app(target=main,view=ft.WEB_BROWSER, port=8000)

#Colocar sempre porta 8000
ft.app(target=main, port=8000,view=ft.WEB_BROWSER, assets_dir="assets")#, host="0.0.0.0")
