import flet as ft
import pymysql
import random
import string
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

import sys
import os
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
        usuario = os.getlogin()
    except:
        return os.environ.get("USERNAME")
    computador = platform.node()

    data_login = datetime.datetime.now()
    data_login_formatado = data_login.strftime("%Y/%m/%d %H:%M:%S")

    Campos = '(Nome, Data_Acesso, Usuario_Acesso, Maquina_Acesso) VALUES (%s, %s, %s, %s)'

    dados = (
        Pessoa,
        data_login_formatado,
        usuario,
        computador
    )

    return inserir_banco('QuestRH_Acessos', dados , Campos)
        
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

###########################################################################################################################################      
############################################################ Inicio da Aplicação ##########################################################
########################################################################################################################################### 
IDLE_TIMEOUT = 300  # segundos -> 5*60 - 5 Minutos

def main(page: ft.Page):
    print(f'🚀 Iniciando aplicação... a imagem está no diretório:{image_path}')
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
    txt_observacoes = ft.TextField(label="Observações", expand=True, multiline=True,max_length=1000)
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

    ####################################################################################################################################### 
    ######################################################## Funções dos Objetos ##########################################################
    #######################################################################################################################################
    
    last_interaction = {"time": time.monotonic()}
 
    def reset_idle_time(e=None):
        last_interaction["time"] = time.monotonic()
        deslog_overlay.visible = False
        page.update()

    async def check_idle():
        while True:
            if Gestor_tempo_formulario.visible == True:
                await asyncio.sleep(1)
                now = time.monotonic()
                elapsed = now - last_interaction["time"]
                remaining = int(IDLE_TIMEOUT - elapsed)

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
        altura_tela = page.window.height

        #container_obs_auto.height = altura_tela * 0.3
        #container_obs_av1.height = altura_tela * 0.3
        #container_obs_av1.height = altura_tela * 0.3

        #corpo_tabela_pend.height = altura_tela * 0.6
        #corpo_tabela.height = altura_tela * 0.6
        #container_perguntas.height = altura_tela * 0.6 
        page.update()
    
    
    def voltar_login(e):
        nome_cb.value =''
        senha_txt.value =''
        login_view.visible =True
        Gestor_tempo_formulario.visible = False
        formulario_Envio.visible = False
        painel_view.visible = False
        painel_pend_view.visible = False
        page.update()

    def voltar_login(e):
        nome_cb.value =''
        senha_txt.value =''
        login_view.visible =True
        Gestor_tempo_formulario.visible = False
        formulario_Envio.visible = False
        painel_view.visible = False
        painel_pend_view.visible = False
        page.update()



    def voltar_painel(e=None, atualizar = False):
        mensagem_aguarde.value = 'Aguarde, atualizando o relatório...'
        aguarde_overlay.visible = True

        page.update()
        nome_pessoa = texto_ola.value.replace("Olá, ","")

        if nome_pessoa == 'Administrador':
            painel_pend_view.visible = True
            painel_view.visible = False
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
            mostrar_alerta_temporario('Cancelamento de resposta realizado com sucess', ft.Colors.GREEN_400)
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
                texto_ola.value = f"Olá, {resultados['Nome']}"
                texto_ola1.value = f"Olá, {resultados['Nome']}"
                texto_ola2.value = f"Olá, {resultados['Nome']}"
                texto_ola3.value = f"Olá, {resultados['Nome']}"
                texto_ola4.value = f"Olá, {resultados['Nome']}"
                aguarde_overlay.visible = True
                page.update()
                
                registra_login(nome_cb.value)

                if texto_ola.value == 'Olá, Administrador':
                    questionarios = lista_pendencias()
                    montar_tabela_pendencias(questionarios)
                    painel_pend_view.visible = True
                else:
                    questionarios = obter_questionarios(resultados['Nome'])
                    montar_tabela(questionarios)
                    login_view.visible = False
                    painel_view.visible = True
                
                painel_resposta_view.visible= False    
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
        
        dropdown_desempenho.on_change=muda_cor_dropdown
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

            campos = '''(Participante, Cargo, C_Custo,  Local, Avaliacao, Nome_Avaliador, ID_Rel, ID_Pergunta, Pilar, Competencia, Pergunta, Resposta,Desempenho_tecnico, Observacao, Data_Resp, Computador, Login) 
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'''
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
                usuario
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
        on_click=lambda _: exportar_para_excel(),
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
        expand=True,
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
        border_radius=ft.border_radius.only(bottom_left=10, bottom_right=10)
    )

    painel_pend_view = ft.Container(
        content=ft.Column([
            ft.Row([ft.TextButton("Deslogar", icon=ft.Icons.ARROW_BACK, on_click=voltar_login),texto_ola2], spacing= 10),
            expiracao_txt,
            ft.Text("Painel de Controle", size=25, weight=ft.FontWeight.BOLD),
            ft.Row([atualizar_btn,exportar_btn],spacing=10),
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
        expand=True,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=True
    )

    formulario_view =  ft.Container(
            content=ft.Column([
                ft.Row([ ft.TextButton("Voltar", on_click=voltar_painel, icon=ft.Icons.ARROW_BACK),texto_ola3]),
                nome_em_avaliacao,
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
        height= 300,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )
    container_respostas_av1 = ft.Container(
        content=lista_reultados_av1_view,
        height= 300,
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

    container_respostas_av2 = ft.Container(
        content=lista_reultados_av2_view,
        height= 300,
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
    
    conteudo_central = ft.Container(
        content=ft.Column([
            login_view,
            painel_view,
            painel_pend_view,
            painel_resposta_view,
            Gestor_tempo_formulario,
            formulario_Envio
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

    page.theme_mode = ft.ThemeMode.LIGHT
    page.on_resized = atualizar_altura_container

    stack = ft.Stack([
        imagem_fundo,
        overlay
    ])

   
    page.add(stack) 
    
ft.app(target=main,view=ft.WEB_BROWSER)
'''
ft.app(target=main, 
        view=ft.WEB_BROWSER,  
        port=8000,
        host="0.0.0.0"
    )

'''