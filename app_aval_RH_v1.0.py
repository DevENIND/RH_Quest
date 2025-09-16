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

import os
import tempfile
import secrets
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
image_path = BASE_DIR / "Imagem_Quest.png"

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


def lista_pendencias():
    scrp_sql = f"SELECT * FROM QuestRH_Relacoes order by Participante"
    conn = mysql_connection()
    cursor = conn.cursor(pymysql.cursors.DictCursor)
    cursor.execute(scrp_sql)
    resultados = cursor.fetchall()
    conn.close()

    lista = []

    for row in resultados:
        #Adicionando Auto Avaliação
        Status = define_status(row['Participante'], row['Participante'])
        Resp_aval = captura_valor_nota(row['Participante'],row['Participante'])

        Status1 = define_status(row['Participante'], row['Avaliador1'])
        Resp_aval1 = captura_valor_nota(row['Participante'],row['Avaliador1'])

        Status2 = define_status(row['Participante'], row['Avaliador2'])
        Resp_aval2 = captura_valor_nota(row['Participante'],row['Avaliador2'])

        lista.append(
                {"Participante": row['Participante'], 
                "Status": Status,
                "Avaliacao": Resp_aval,
                "Avaliador1": row['Avaliador1'],
                "Status1": Status1,
                "Avaliacao1": Resp_aval1,
                "Avaliador2": row['Avaliador2'],
                "Status2": Status2,
                "Avaliacao2": Resp_aval2,
                "Questionário": row['Tipo_Avaliacao']
                }
            )
        
    return lista

def define_avaliacao_final(Pessoa):
    scrp_sql = f"SELECT * FROM QuestRH_Relacoes Where Participante = '{Pessoa}'"
    conn = mysql_connection()
    cursor = conn.cursor(pymysql.cursors.DictCursor)
    cursor.execute(scrp_sql)
    resultados = cursor.fetchone()
    conn.close()

    Status = define_status(resultados['Participante'], resultados['Participante'])
    Status1 = define_status(resultados['Participante'], resultados['Avaliador1'])
    Status2 = define_status(resultados['Participante'], resultados['Avaliador2'])

    if Status == 'Realizado' and Status1 == 'Realizado' and Status2 == 'Realizado':
        scrp_sql = f"SELECT ROUND(AVG(Resposta)) as Media FROM QuestRH_Respostas Where Participante = '{Pessoa}' and ID_Rel > 0"
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(scrp_sql)
        resultados = cursor.fetchone()
        conn.close()
        if resultados:
            return resultados['Media']
        else:
            return 0
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
        conn.close()

        for row in resultados:
            if row['Participante'] == Pessoa:
                Avaliador = 0
            elif row['Avaliador1'] == Pessoa:
                Avaliador = 1
            elif row['Avaliador2'] == Pessoa:
                Avaliador = 2
            else:
                Avaliador = ''
            
            if Avaliador == 2 or Avaliador == 0:
                if define_status(row['Participante'],row['Participante']) == 'Realizado':
                    Resp_auto = "Sim"
                else:
                    Resp_auto = "Não"

                if define_status(row['Participante'],row['Avaliador2']) == 'Realizado':
                    Resp_aval2 = "Sim"
                else:
                    Resp_aval2 = "Não"

                if define_status(row['Participante'],row['Avaliador1']) == 'Realizado':
                    Resp_aval1 = "Sim"
                else:
                    Resp_aval1 = "Não"
            else:
                Resp_auto = captura_valor_nota(row['Participante'],row['Participante'])
                Resp_aval1 = captura_valor_nota(row['Participante'],row['Avaliador1']) 
                Resp_aval2 = captura_valor_nota(row['Participante'],row['Avaliador2']) 

            
            if Pessoa == 'Administrador':
                if Resp_auto == '' or Resp_aval1 == '' or Resp_aval2 == '':
                    Status = 'Pendente'
                else:
                    Status = 'Realizado'

            else:
                Status = define_status(row['Participante'], Pessoa)

            lista.append(
                {"nome": row['Participante'], 
                 "Avaliador": Avaliador, 
                 "Questionário": row['Tipo_Avaliacao'],
                 "status": Status,
                 "auto_aval": Resp_auto,
                 "primaria": Resp_aval1, 
                 "secundaria": Resp_aval2}
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
    page.scroll = ft.ScrollMode.AUTO
    page.bgcolor = ft.Colors.TRANSPARENT
    page.window.maximized = True
    data_limite = '2025/09/30 23:59:59'
    #data_limite = '2025/07/21 23:59:59'
    file_picker = ft.FilePicker()
    page.overlay.append(file_picker)
    linhas = []

    # Container de alerta
    alerta_container = ft.Container(
        content=ft.Text("", color=ft.Colors.WHITE),
        bgcolor=ft.Colors.RED_400,
        height=40,
        padding=10,
        alignment=ft.alignment.center,
        visible=False
    )

    # Container de alerta
    alerta_container_form = ft.Container(
        content=ft.Text("", color=ft.Colors.WHITE),
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

    senha_txt =  ft.TextField(label="Digite a senha para entrar", expand=True , visible=True, password=True)
    expiracao_txt = ft.Text(f'Data para envio de formulário expirado no dia {data_limite}.',expand= True, visible= False, size=20, weight=ft.FontWeight.BOLD, bgcolor=ft.Colors.RED_400, text_align=ft.alignment.center)

    mensagem_aguarde = ft.Text(
                    "Aguarde, atualizando o relatório...",
                    size=18,
                    weight=ft.FontWeight.BOLD,
                    text_align=ft.TextAlign.CENTER,
                    color=ft.Colors.WHITE
                )


    erro_login = ft.Text("", color=ft.Colors.RED)
    txt_observacoes = ft.TextField(label="Observações", expand=True, multiline=True,max_length=20000, 
                                   on_change=lambda e: reset_idle_time(e), 
                                   on_click=lambda e: reset_idle_time(e),
                                   on_blur=lambda e: reset_idle_time(e),
                                   on_focus=lambda e: reset_idle_time(e),
                                   on_submit=lambda e: reset_idle_time(e),
                                   on_animation_end=lambda e: reset_idle_time(e),
                                    visible=True)
    nome_cb = ft.TextField(label="Login", expand=True, visible=True)
    form_inputs = []

    txt_observacoes_auto = ft.Text("Observações", expand=True, size=20, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900)
    container_obs_auto = ft.Container(
                            content=txt_observacoes_auto,
                            bgcolor=ft.Colors.GREY_50,      
                            padding=15, 
                            border_radius=10,
                            visible= False
                            )

    
    txt_observacoes_av1 =ft.Text("Observações", expand=True, size=20, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900)
    container_obs_av1 = ft.Container(
                            content=txt_observacoes_av1, 
                            bgcolor=ft.Colors.GREY_50,      
                            padding=15, 
                            border_radius=10,
                            visible= False
                            )

    
    txt_observacoes_av2 = ft.Text("Observações", expand=True, size=20, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900)
    container_obs_av2 = ft.Container(
                            content=txt_observacoes_av2, 
                            bgcolor=ft.Colors.GREY_50,      
                            padding=15, 
                            border_radius=10,
                            visible= False
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
                            border_radius=10,
                            visible= False
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
                            border_radius=10,
                            visible= False
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
                            border_radius=10,
                            visible= False
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

    lista_reultados_auto_view = ft.ListView(expand=True, auto_scroll=False, spacing=10, padding=10,visible=False)
    lista_reultados_av1_view = ft.ListView(expand=True, auto_scroll=False, spacing=10, padding=10,visible=False)
    lista_reultados_av2_view = ft.ListView(expand=True, auto_scroll=False, spacing=10, padding=10,visible=False)

    txt_auto = ft.Text("", size=20, weight=ft.FontWeight.BOLD)
    txt_av1 = ft.Text("", size=20, weight=ft.FontWeight.BOLD)
    txt_av2 = ft.Text("", size=20, weight=ft.FontWeight.BOLD)

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

    txt_media_final = ft.Text("", size=20, weight=ft.FontWeight.BOLD,color=ft.Colors.BLUE_900)

    container_media_final = ft.Container(
                            content=txt_media_final, 
                            bgcolor=ft.Colors.BLUE_100,      
                            padding=15, 
                            border_radius=10,
                            expand= True,
                            alignment=ft.alignment.center,
                            visible= True
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
    
    # textos -> Referencia de tamanhos de fontes =https://flet.dev/docs/controls/text/#font_family
    txt_Realizados = ft.Text('', size=40, weight='w200', font_family='roboto',text_align="center", color=ft.Colors.GREEN)
    txt_Finalizados = ft.Text('', size=40, weight='w200', font_family='roboto',text_align="center",color=ft.Colors.BLUE)
    txt_Pendentes = ft.Text('', size=40, weight='w200', font_family='roboto',text_align="center",color=ft.Colors.RED)
    txt_Pendentes_participantes = ft.Text('', size=40, weight='w200', font_family='roboto',text_align="center",color=ft.Colors.AMBER)

   

    ####################################################################################################################################### 
    ######################################################## Funções dos Objetos ##########################################################
    #######################################################################################################################################
    
    #-------------------------------------------------------------------------------------------------------------------------------------------
    #---------------------------------------------------------------- Gráficos -----------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------------------------------------------------
    def atualiza_tabela_bellcurve(e=None):
        Media = dropdown_bellcurve_nota.value

        if Media:
            float(Media)

        try:
            if not Media is None and Media != '':
                query_sql = f"""
                SELECT Participante, Media from(
                    SELECT Participante, ROUND(AVG(Resposta),0) as Media
                    FROM QuestRH_Respostas
                    WHERE id_rel > 0
                    GROUP BY Participante
                    HAVING COUNT(DISTINCT id_rel) = 2
                    ORDER BY Participante) x
                Where Media = {Media}
            """
            else:
                query_sql =f"""
                SELECT Participante, Media from(
                    SELECT Participante, ROUND(AVG(Resposta),0) as Media
                    FROM QuestRH_Respostas
                    WHERE id_rel > 0
                    GROUP BY Participante
                    HAVING COUNT(DISTINCT id_rel) = 2
                    ORDER BY Participante) x
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
        cboxPessoa.value = None
        cboxPilar.value = None
        cboxCompetencia.value = None
        dropdown_Performance.value = None

        cboxPessoa.update()
        cboxPilar.update()
        cboxCompetencia.update()
        dropdown_Performance.update()

        for chk in dropdown_c_custo.controls[0].controls:
            chk.selected = False

        for chk in dropdown_empresa.controls[0].controls: 
            chk.selected = False
         
        for chk in dropdown_cargo.controls[0].controls:
            chk.selected = False

        txt_outro_query.value = ''
        page.update()
        
        alimenta_competencias()
        alimenta_pessoas()
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


        grafico64, msgerro = gera_graficos.gera_ninebox(pilar=pilar, competencia=competencia, participante=participante,outras_condicoes=txt_outro_query.value)
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
        
        grafico_bell64, msgerro_compar = gera_graficos.gera_bell_curve()
        if grafico_bell64 == None:
            img_BellCurve.visible = False
        else:  
            img_BellCurve.visible = True 

        img_Ninebox.src_base64 = grafico64
        img_Pilar.src_base64 = grafico_pilar64
        img_Comp.src_base64 = grafico_comp64
        img_Compar.src_base64 = grafico_compar64
        img_BellCurve.src_base64 = grafico_bell64

        alimenta_tabela_graficos(Participante=participante, outras_condicoes=txt_outro_query.value, performance=performance)
        
        aguarde_overlay.visible = False
        page.update()
    
    def exportar_grafico_excel(e=None):
        try:
            # Conexão e consulta
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            scrp_sql ="""
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
                    ) as t;
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
            sql_condicao = f" where Participante = '{Participante}' "
        else:
            sql_condicao = ''

        

        scrp_sql =f"""
        SELECT  * From (SELECT Participante, Avaliacao, Sigla_Emp, C_Custo, Cargo, 
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
            ) x
        """

        if performance != '':
            scrp_sql += f" Where x.Performance = '{performance}'"
    

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

        txt_outro_query.value = " and ".join(query_parts) if query_parts else ''
        
        atualiza_dados()
        #return

    #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------- Inicio Função De Diálogo --------------------------------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
    tempo_ociosidade = {'segundos': 0}
 
    def reset_idle_time(e=None):
        #print('timer resetado')
        last_interaction["time"] = time.monotonic()
        txt_ociosidade.value = 'Tempo de inatividade: 0s (00:00:00)'
        tempo_ociosidade['segundos'] = 0
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
                tempo_ociosidade['segundos'] += 1 
                tempo_delta = datetime.timedelta(seconds=tempo_ociosidade['segundos'])
                tempo_formatado = str(tempo_delta)

                txt_ociosidade.value = f'Tempo de inatividade: {tempo_ociosidade['segundos']}s ({tempo_formatado})'
                txt_ociosidade.update()
             


                if remaining > 0:
                    if remaining <= 40:
                        deslog_overlay.visible = True

                    msg_deslog.value = f"Você está muito tempo inativo, está por aí? o questionário será fechado em {remaining} segundos"
                else:
                    deslog_overlay.visible = False
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
                page.launch_url("https://enindengenharia-my.sharepoint.com/:b:/g/personal/wagner_barreiro_enind_com_br/EVTZZEKppENKtxKjAVAmZZcBFQtA84M_e-hGFLQp4EHYYQ?e=YjJhPO")
            elif resultado['Tipo_Avaliacao'] == 'A2':
                page.launch_url("https://enindengenharia-my.sharepoint.com/:b:/g/personal/wagner_barreiro_enind_com_br/EcQaPPgtXNRCkfMWt3uNVtoBWorBkNJnPNAURJrQOTFviw?e=wKcXOf")
            else:
                page.launch_url("https://enindengenharia-my.sharepoint.com/:b:/g/personal/wagner_barreiro_enind_com_br/EWr3VQ6suA1InksNT7CADgYBiFM4i7de--l2KeXY_iwfNA?e=iYMGMr")
        else:
            page.launch_url("https://enindengenharia-my.sharepoint.com/:b:/g/personal/wagner_barreiro_enind_com_br/EVTZZEKppENKtxKjAVAmZZcBFQtA84M_e-hGFLQp4EHYYQ?e=YjJhPO")
    
    def atualizar_altura_container(e):
        altura_tela = e.height
    
        container_obs_auto.height = altura_tela * 0.3
        container_obs_av1.height = altura_tela * 0.3
        container_obs_av1.height = altura_tela * 0.3

        corpo_tabela_pend.height = altura_tela * 0.7
        corpo_tabela.height = altura_tela * 0.7
        container_perguntas.height = altura_tela * 0.7
        page.update()
    
    
    def voltar_login(e):
        nome_cb.value =''
        senha_txt.value =''
        
        page.clean()
        conteudo_central = ft.Container(
            content=ft.Column([
                    login_view,
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=30),
                alignment=ft.alignment.center,
            )
        
        overlay = ft.Column([
            alerta_container,
            conteudo_central,
            alerta_container
            ], 
        expand=True,
        alignment=ft.alignment.bottom_center 
        )

        stack = ft.Stack([
            imagem_fundo,
            overlay
        ])
        page.add(stack)

        login_view.visible =True
        page.update()

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
                questionarios = lista_pendencias()
                montar_tabela_pendencias(questionarios)
        else:
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
        nome_pessoa = texto_ola.value.replace("Olá, ",'')
        aguarde_overlay.visible = True
        page.update()

        if nome_pessoa == 'Administrador':
            painel_pend_view.visible = True
            painel_view.visible = False
            questionarios = lista_pendencias()
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
    nome_em_avaliacao = ft.Text("", size=25, weight=ft.FontWeight.BOLD)
    nome_avaliado = ft.Text("", size=25, weight=ft.FontWeight.BOLD)

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
                    page.clean()

                    overlay = ft.Column([
                        alerta_container,
                        painel_admin,
                        alerta_container
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

                    questionarios = lista_pendencias()
                    montar_tabela_pendencias(questionarios)
                    painel_pend_view.visible = True
                    painel_resposta_view.visible= False 
                    
                elif avaliador1:
                    page.clean()

                    overlay = ft.Column([
                        alerta_container,
                        painel_av1,
                        alerta_container,
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

                    overlay = ft.Column([
                        alerta_container,
                        painel_comum,
                        alerta_container,
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

            btn_ver_resp = (
                ft.TextButton('Visualizar',icon=ft.Icons.VISIBILITY, on_click=lambda e, nome=q["nome"]: abrir_formulario_respostas(nome))
                if (q['Avaliador'] == 1) else ft.Text("")
            )

             # Dados que vão rolar
           
            linha = ft.Container(
                content=ft.Row([
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
                        ft.Container(btn_ver_resp, expand=1)
                    ], spacing=10),
                padding=10,
                bgcolor=ft.Colors.TRANSPARENT,
                border_radius=8,
                margin=ft.margin.only(bottom=4)
            )
            
            lista_view.controls.append(linha)

            
        if data_atual >=data_fechamento:
            expiracao_txt.visible = True

        lista_pend_view.visible = True
        lista_view.update()
        page.update()

    #Função para criar a tabela de visualização - Painel de controle - PENDENCIAS
    def montar_tabela_pendencias(questionarios):
        nonlocal data_limite
        nonlocal linhas
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
                ft.TextButton('Ver', icon=ft.Icons.VISIBILITY, on_click=lambda e, nome=q["Participante"]: abrir_formulario_respostas(nome))
            )

             # Dados que vão rolar
            linha = ft.Container(
                content=ft.Row([
                        ft.Container(ft.Text(q["Participante"]), expand=3),
                        ft.Container(
                            content=ft.Text(Status),
                            bgcolor=bg_cor,
                            padding=ft.padding.symmetric(horizontal=6, vertical=2),
                            border_radius=5,
                            alignment=ft.alignment.center,
                            expand=1
                        ),
                        ft.Container(ft.Text(q["Avaliacao"]), expand=1, alignment=ft.alignment.center),
                        ft.Container(ft.Text(q["Avaliador1"]), expand=3),
                        ft.Container(
                            content=ft.Text(q["Status1"]),
                            bgcolor=bg_cor1,
                            padding=ft.padding.symmetric(horizontal=6, vertical=2),
                            border_radius=5,
                            alignment=ft.alignment.center,
                            expand=1
                        ),
                        ft.Container(ft.Text(q["Avaliacao1"]), expand=1, alignment=ft.alignment.center),
                        ft.Container(ft.Text(q["Avaliador2"]), expand=3),
                         ft.Container(
                            content=ft.Text(q["Status2"]),
                            bgcolor=bg_cor2,
                            padding=ft.padding.symmetric(horizontal=6, vertical=2),
                            border_radius=5,
                            alignment=ft.alignment.center,
                            expand=1
                        ),
                        ft.Container(ft.Text(q["Avaliacao2"]), expand=1, alignment=ft.alignment.center),
                         ft.Container(btn_ver_resp, expand=1),
                    ], spacing=10),
                padding=10,
                bgcolor=ft.Colors.TRANSPARENT,
                border_radius=8,
                margin=ft.margin.only(bottom=4)
            )
            
            lista_pend_view.controls.append(linha)
        
        

        if data_atual >=data_fechamento:
            expiracao_txt.visible = True

        lista_pend_view.visible = True
        lista_pend_view.update()
        page.update()
    
    

     #Função para construir os formulários de respostas
    def abrir_formulario_respostas(Pessoa):
        container_respostas_auto.visible = False
        container_respostas_av1.visible = False
        container_respostas_av2.visible = False
        
        container_obs_auto.visible =False
        container_obs_av1.visible = False 
        container_obs_av2.visible = False    

        container_desempenho_auto.visible =False
        container_desempenho_av1.visible =False
        container_desempenho_av2.visible =False

        txt_av1.value = ''
        txt_av2.value  = ''
        txt_auto.value  = ''

        txt_av1.visible =False
        txt_av2.visible =False
        txt_auto.visible =False

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
                txt_observacoes_auto.value = f"Observação: {row['Observacao']}" 
                txt_auto.value = f"Auto Avaliação: {row['Nome_Avaliador']}"
                txt_auto.visible = True

                if row['Desempenho_tecnico'] or row['Desempenho_tecnico']>0:
                    container_desempenho_auto.visible =True
                    txt_desempenho_auto.value = lista_avaliação[row['Desempenho_tecnico'] - 1]['Resp']
                    container_desemp_auto.bgcolor = lista_avaliação[row['Desempenho_tecnico'] - 1]['cor']
                else:
                    txt_desempenho_auto.value =''
                    container_desempenho_auto.visible =False

                lista_reultados_auto_view.controls.append(container_pilar)
                lista_reultados_auto_view.visible = True
                container_respostas_auto.visible = True
                container_obs_auto.visible = True

            elif row['ID_Rel'] == 1:
                txt_observacoes_av1.value = f"Observação: {row['Observacao']}" 
                txt_av1.value = f"Avaliador1: {row['Nome_Avaliador']}"
                txt_av1.visible = True

                if row['Desempenho_tecnico'] or row['Desempenho_tecnico']>0:
                    container_desempenho_av1.visible =True
                    txt_desempenho_av1.value = lista_avaliação[row['Desempenho_tecnico'] - 1]['Resp']
                    container_desemp_av1.bgcolor = lista_avaliação[row['Desempenho_tecnico'] - 1]['cor']
                else:
                    txt_desempenho_av1.value =''
                    container_desempenho_av1.visible =False

                lista_reultados_av1_view.controls.append(container_pilar)
                lista_reultados_av1_view.visible = True
                container_respostas_av1.visible = True
                container_obs_av1.visible = True
            else:
                txt_observacoes_av2.value = f"Observação: {row['Observacao']}" 
                txt_av2.value = f"Avaliador2: {row['Nome_Avaliador']}"
                txt_av2.visible = True

                if row['Desempenho_tecnico'] or row['Desempenho_tecnico']>0:
                    container_desempenho_av2.visible =True
                    txt_desempenho_av2.value = lista_avaliação[row['Desempenho_tecnico'] - 1]['Resp']
                    container_desemp_av2.bgcolor = lista_avaliação[row['Desempenho_tecnico'] - 1]['cor']
                else:
                    txt_desempenho_av2.value =''
                    container_desempenho_av2.visible =False
        
                lista_reultados_av2_view.controls.append(container_pilar)
                lista_reultados_av2_view.visible = True
                container_respostas_av2.visible = True
                container_obs_av2.visible = True

        
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
            txt_media_final.value = f'A média final é: {nota_final}'
        else:
            txt_media_final.value = f'A média final ainda não foi definida'

        # Exibir painel de respostas

        lista_reultados_auto_view.controls.append(container_desempenho_auto)
        lista_reultados_av1_view.controls.append(container_desempenho_av1)
        lista_reultados_av2_view.controls.append(container_desempenho_av2)


        lista_reultados_auto_view.controls.append(container_obs_auto)
        lista_reultados_av1_view.controls.append(container_obs_av1)
        lista_reultados_av2_view.controls.append(container_obs_av2)

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

    #Função para construir os formulários de respostas
    def abrir_formulario(nome):
        nome_em_avaliacao.value = f'Você está avaliando: {nome}'
        txt_observacoes.value =''
        form_inputs.clear()
        form_content.controls.clear()
        perguntas_formulario = obter_perguntas(nome)
        dropdown_desempenho.value = 0

        # Opções padrão
        opcoes_avaliacao = [
            ft.dropdown.Option("1 - Insatisfatório - Não atende ou atende minimamente aos padrões"),
            ft.dropdown.Option("2 - Regular - Atende parcialmente aos padrões esperados"),
            ft.dropdown.Option("3 - Satisfatório - Atende os padrões esperados"),
            ft.dropdown.Option("4 - Bom - Demonstra empenho e excelência no atendimento de padrões esperados"),
            ft.dropdown.Option("5 - Excelente - Supera os padrões esperados")
        ]

        alerta_container_form.visible = False
        alerta_container_form.content.value = ''

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
                texto_pergunta = f"{row['ID']}) {row['Pergunta']}"
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

        form_content.controls.append(ft.Column(
                    controls=[
                        obj_texto,
                        container_desempenho
                    ],
                    spacing=5))
        
        form_content.controls.append(txt_observacoes)

        painel_view.visible = False
        Gestor_tempo_formulario.visible = True
        reset_idle_time()
        page.update()
        page.run_task(check_idle)


    #Função para enviar as resposta dos usuários
    def enviar_formulario(e):
        respostas = []
        mensagem_aguarde.value = 'Aguarde, enviando respostas...'
        aguarde_overlay.visible = True
        page.update()

        for i, grupo in enumerate(form_inputs):
            if not grupo.value:
                mostrar_alerta_temporario("Preencha todos os campos antes de enviar o furmulário...", ft.Colors.RED_400)
                alerta_container_form.visible = True
                alerta_container_form.content.value = 'Preencha todos os campos antes de enviar o furmulário...'
                aguarde_overlay.visible = False
                page.update()
                return  # Interrompe envio
            else:
                valor_resp = str(grupo.value).split(' - ')[0]
                respostas.append({
                    'ID':i+1,
                    'resposta': valor_resp})
        
        if not dropdown_desempenho.value or dropdown_desempenho.value == 0 :
            mostrar_alerta_temporario("Coloque antes o valor do desempenho da pessoa.", ft.Colors.RED_400)
            alerta_container_form.visible = True
            alerta_container_form.content.value = 'Coloque antes o valor do desempenho da pessoa.'
            aguarde_overlay.visible = False
            page.update()
            return  # Interrompe envio
        else:
            valor_desempenho = dropdown_desempenho.value.split(' - ')[0]

        obs_us = txt_observacoes.value
        if valida_texto(obs_us) == False:
            mostrar_alerta_temporario('Campo de obsevações contem palavras não permitidas, por gentileza, analise o texto', ft.Colors.RED_400)
            alerta_container_form.visible = True
            alerta_container_form.content.value = 'Campo de obsevações contem palavras não permitidas, por gentileza, analise o texto'
            aguarde_overlay.visible = False
            page.update()
            return

        avaliador = texto_ola.value.replace("Olá, ",'')
        participante = nome_em_avaliacao.value.replace('Você está avaliando: ', "")
        data_envio = datetime.datetime.now()
        data_envio_formatado = data_envio.strftime("%Y/%m/%d %H:%M:%S")

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

            campos = '''(Participante, Cargo, C_Custo,  Local, Avaliacao, Nome_Avaliador, ID_Rel, ID_Pergunta, Pilar, Competencia, Pergunta, Resposta,Desempenho_tecnico, Observacao, Data_Resp, Computador, Login, Empresa, Sigla_Emp) 
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'''
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
                int(row['resposta']),
                valor_desempenho,
                obs_us,
                data_envio_formatado,
                computador,
                usuario,
                consulta_pessoa['Empresa'],
                consulta_pessoa['Sigla_Emp']
            )

            validacao = inserir_banco('QuestRH_Respostas',valores, campos)
            if validacao ==False:
                mostrar_alerta_temporario("Não foi possível inerir os dados no banco de dados, tente novamente.", ft.Colors.RED_400)
                aguarde_overlay.visible = False
                page.update()
                return
        
        participante_realizado.value = participante
        formulario_Envio.visible = True
        Gestor_tempo_formulario.visible = False
        mensagem_aguarde.value = 'Aguarde, atualizando o relatório...'
        aguarde_overlay.visible = False
        page.update()
    
    def exportar_para_excel(e=None):
        try:
            # Conexão e consulta
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            scrp_sql = "SELECT * FROM QuestRH_Respostas"
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
            nome_arquivo = f"exportacao_respostas_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
             
            # Codificar para base64
            b64 = base64.b64encode(output.read()).decode()

            # Criar link de download
            link_download = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"

            # Abrir o link no navegador (força o download)
            page.launch_url(link_download, web_window_name=nome_arquivo)

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

    login_view = ft.Container(
        content=ft.Column([
            ft.Text("Login", size=30, weight=ft.FontWeight.BOLD),
            nome_cb,
            senha_txt,
            ft.ElevatedButton("Entrar", on_click=validar_login),
            erro_login
        ],
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        width=600,
        spacing=15),
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=True
    )

    file_picker.on_result = ao_escolher_pasta
    
    exportar_material_btn = ft.ElevatedButton(
        "Visualizar Material de Apoio",
        on_click=escolher_pasta,
        width=250,
        height=50
    )

    exportar_btn = ft.ElevatedButton(
        "Exportar para Excel",
        icon= ft.Icons.DOWNLOAD_ROUNDED,
        on_click=lambda _: exportar_para_excel(),
        width=250,
        height=50
    )

    exportar_grafico_btn = ft.ElevatedButton(
        "Exportar para Excel",
        icon= ft.Icons.DOWNLOAD_ROUNDED,
        on_click=lambda _: exportar_grafico_excel(),
        width=250,
        height=50
    )

    grafico_btn = ft.ElevatedButton(
        "Gráficos",
        icon=ft.Icons.BAR_CHART_ROUNDED,
        on_click=lambda _: inicializa_grafico(),
        width=250,
        height=50
    )

    grafico_status_btn = ft.ElevatedButton(
        "Status",
        icon=ft.Icons.SPEED_OUTLINED,
        on_click=lambda e: atualizar_painel_status(e),
        width=250,
        height=50
    )

    atualizar_btn = ft.ElevatedButton(
        content=ft.Row(
                [
                    ft.Icon(name=ft.Icons.REFRESH, size=20),
                    ft.Text("Atualizar Relatório")
                ],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=10
            ),
            on_click=atualiza_rel,
            width=250,
            height=50
        )

    # Cabeçalho fixo
    cabecalho = ft.Container(
        content=ft.Row([
            ft.Container(ft.Text("Nome", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Auto Avaliação", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação 1", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação 2", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Status", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text('Avaliar', weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text('Visualizar', weight=ft.FontWeight.BOLD), expand=1),
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
            expiracao_txt,
            ft.Row([ft.Text("Painel de Controle", size=25, weight=ft.FontWeight.BOLD),btn_ver_manual],
                   expand=True, alignment=ft.MainAxisAlignment.SPACE_BETWEEN,vertical_alignment=ft.CrossAxisAlignment.CENTER),
            atualizar_btn,
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
            ft.Container(ft.Text("Participante", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Status", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliador1", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Status", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliador2", weight=ft.FontWeight.BOLD), expand=3),
            ft.Container(ft.Text("Status", weight=ft.FontWeight.BOLD), expand=1),
            ft.Container(ft.Text("Avaliação", weight=ft.FontWeight.BOLD), expand=1),
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

    painel_pend_view = ft.Container(
        content=ft.Column([
            ft.Row([ft.TextButton("Deslogar", icon=ft.Icons.ARROW_BACK, on_click=voltar_login),texto_ola2], spacing= 10),
            expiracao_txt,
            ft.Text("Painel de Controle", size=25, weight=ft.FontWeight.BOLD),
            ft.Row([atualizar_btn,exportar_btn, grafico_btn, grafico_status_btn],spacing=10),
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
        expand=True,
        auto_scroll=False,  # Mantenha a rolagem manual
        padding=0
    )

    container_perguntas = ft.Container(
        content=form_content,
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
                ft.Row([ ft.TextButton("Voltar", on_click=voltar_painel, icon=ft.Icons.ARROW_BACK),texto_ola3]),
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
        content=lista_reultados_auto_view,
        height= page.height * 0.3,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )
    container_respostas_av1 = ft.Container(
        content=lista_reultados_av1_view,
        height= page.height * 0.3,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

    container_respostas_av2 = ft.Container(
        content=lista_reultados_av2_view,
        height= page.height * 0.3,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

    painel_resposta_view = ft.Container(
        content=ft.Column([
            ft.Row([ft.TextButton("Voltar", on_click=voltar_painel, icon=ft.Icons.ARROW_BACK),texto_ola4]),
            nome_avaliado,
            txt_auto,
            container_respostas_auto,
            txt_av1,
            container_respostas_av1,
            txt_av2,
            container_respostas_av2,
            container_media_final 
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
        expand=True
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

    cboxPessoa.on_change = atualiza_dados
    cboxPilar.on_change = atualiza_dados
    cboxCompetencia.on_change = atualiza_dados
    dropdown_Performance.on_change = atualiza_dados
    
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
    
    txt_assinatura = ft.Text('Powered by ENIND Engenharia - Wagner Barreiro all rights reserved', size=10, font_family='roboto', color=ft.Colors.GREY_600)

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
        ft.Row([ft.TextButton("Voltar", icon=ft.Icons.ARROW_BACK, on_click=voltar_painel), texto_ola6], spacing= 10),
        ft.Row([ft.Text("📊 Análise Gráfica - Status das Avaliações", size=20, weight="bold")], spacing= 10),
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
        ])
    ],
    visible=False)

    
    with open(image_path, "rb") as img_file:
            img_base64 = base64.b64encode(img_file.read()).decode('utf-8')

    imagem_fundo = ft.Container(
        content=ft.Image(
            src=f"data:image/png;base64,{img_base64}",
            fit=ft.ImageFit.COVER,
            expand=True,
            gapless_playback=True
        ),
        alignment=ft.alignment.center,
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
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=30),
        alignment=ft.alignment.center,
        expand=True,
    )
    
    overlay = ft.Column([
        alerta_container,
        conteudo_central,
        alerta_container,
        ft.Row([txt_assinatura], alignment=ft.MainAxisAlignment.CENTER),
        ], 
    expand=True,
    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
    horizontal_alignment=ft.CrossAxisAlignment.CENTER
    )

    stack = ft.Stack([
        imagem_fundo,
        overlay
    ])

    page.theme_mode = ft.ThemeMode.LIGHT
    page.on_resized = atualizar_altura_container



    #print('✔️ tamanho da tela: ', page.width, 'x', page.height)
    page.add(stack) 
    
    
#ft.app(target=main,view=ft.WEB_BROWSER)

#Colocar sempre porta 8000
ft.app(target=main,view=ft.WEB_BROWSER, port=8000, host="0.0.0.0")