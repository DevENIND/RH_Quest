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

import sys
import os
#Instalando aplicativo para local Windows
#pyinstaller --noconfirm --onefile --add-data "Imagem_Quest.png;." app.py

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

def prepara_corpo_email(lista_pendentes, log_avaliador, senha_avaliador):
    agora = datetime.datetime.now()
    fuso_horario =  timezone('America/Sao_Paulo') #Horario de São Paulo
    agora = agora.astimezone(fuso_horario)
    hora_agora = agora.time().hour

    
    corpo = 'Prezado(a),'

    if hora_agora >= 0 and hora_agora <= 12:
        corpo = f'Bom Dia!'
    elif hora_agora >= 13 and hora_agora <= 18:
        corpo = f'Boa Tarde!'
    elif hora_agora > 19:
        corpo = f'Boa Noite!'
    

   
    corpo += f"<br><br>Segue abaixo listagem de acessos referente a Avaliação de Desempenho da ENIND <b>essa é uma informação confidencial, não repasse esse e-mail ou responda.</b> "
    corpo += f"<br><br>Você é o responsável pela segurança e acesso das credenciais fornecidas na listagem abaixo."
    corpo += f'<br><br>🔐 <b>Segue abaixo sua credencial de acesso:</b> <br>Login: {log_avaliador} <br> Senha: {senha_avaliador}'
    corpo += f"<br><br><b>⚠️ ** ATENÇÃO **</b> Acesse a área remota (acesso a rede da sede) abra o aplicativo constante na área de trabalho para realizar as avaliações."
    corpo += f"<br><br><b>❓ Possui dúvidas ou problemas contate:</b>"
    corpo += f"<br>sobre o <b>aplicativo:</b> larissa.schons@enind.com.br"
    corpo += f"<br>sobre o acesso a <b>area de trabalho remota:</b> arthur.magno@enind.com.br"
    corpo += f"<br><br>{lista_pendentes}<br><br>"
    corpo += f"<b>E-mail automático, por gentileza, não responda.</b>"
    corpo += f"<br> Atenciosamente,"

    return corpo


def enviar_email(para, assunto, corpo, cc=''):
    try:
        #sender = 'NF@enind.com.br'
        #password = 'Enind@2020'

        sender = 'nao-responda@enind.com.br'
        password = 'N102030r!@'

        lista_para = [email.strip() for email in para.split(",") if email.strip()]
        lista_cc = [email.strip() for email in cc.split(",") if email.strip()]
        todos_destinatarios = lista_para + lista_cc

        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] =  ", ".join(lista_para)
        msg['CC'] = ", ".join(lista_cc)
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
            smtp_server.sendmail(sender, todos_destinatarios, raw)
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

        if resultados['Tipo_Avaliacao']=='A3' and Participante == Pessoa:
            return 'Realizado'
        else:
            return 'Pendente'

###########################################################################################################################################      
########################################################### Inicio da Aplicação ############################################################
########################################################################################################################################### 
lista_email = []

scrp_cabec_tabela = '''<h3>📌Relatório de Acessos e Pendências</h3><table border="1" cellpadding="5" cellspacing="0">
            <tr>
                <th>Nome</th>
                <th>Tipo_Avaliador</th>
                <th>Usuários</th>
                <th>Senha</th>
            </tr>
    '''

resultados = pd.read_excel('lista_emails.xlsx')
resultados.columns = resultados.columns.str.strip()
scrp_tabela = ''
email_enviar = ''

bol_enviar = False
start = True
for index, row in resultados.iterrows():
    if index == 0:
        scrp_tabela = scrp_cabec_tabela
    
    email = row['email_repres']
    status = define_status(row['Participante'],row['Avaliador'])
   
    '''
    if email_enviar == 'wilson.fernandes@enind.com.br':
        start = True
    else:
        scrp_tabela = scrp_cabec_tabela
    '''

    if email_enviar != email and index > 0 and bol_enviar == True and start == True: 
        scrp_tabela += '</table>'
        corpo_email = f'email a enviar para: <b>{email_enviar}</b><br><br>'
        corpo_email += prepara_corpo_email(scrp_tabela, log_avaliador, senha_avaliador)
        boll_env = enviar_email('wagner.barreiro@enind.com.br, larissa.schons@enind.com.br','Avaliação de Desempenho ENIND', corpo_email)

        if boll_env == False:
            print(f'❌ Erro ao enviar email para:{email_enviar} referente a {nome_enviar}')
        else:
            print(f'✔️ Email enviado para:{email_enviar} referente a {nome_enviar}')
        bol_enviar = False

        scrp_tabela = scrp_cabec_tabela

    if status == 'Pendente':

        bol_enviar = True
        email_enviar = email
        nome_enviar = row['Avaliador']
        log_avaliador = row['login_aval']
        senha_avaliador = row['senha_aval']


        if row['Resp_Av1'] == 'Sim':
             scrp_tabela += f'''
                    <tr>            
                        <td>{row['Avaliador']}</td>
                        <td></td>
                        <td>{row['login_aval']}</td>
                        <td>{row['senha_aval']}</td>
                    </tr> 
            '''
        elif row['Resp_Part'] == 'Sim':
            scrp_tabela += f'''
                    <tr>            
                        <td>{row['Participante']}</td>
                        <td>{row['Tipo_aval']}</td>
                        <td>{row['login_part']}</td>
                        <td>{row['senha_part']}</td>
                    </tr> 
            '''
        else:
            scrp_tabela += f'''
                        <td>{row['Participante']}</td>
                        <td>{row['Tipo_aval']}</td>
                        <td></td>
                        <td></td>
                    </tr> 
            '''

if bol_enviar == True:             
    scrp_tabela += '</table>'
    corpo_email = f'email a enviar para: <b>{email_enviar}</b><br><br>'
    corpo_email += prepara_corpo_email(scrp_tabela, log_avaliador, senha_avaliador)
    boll_env = enviar_email('wagner.barreiro@enind.com.br, larissa.schons@enind.com.br','Avaliação de Desempenho ENIND', corpo_email)

    if boll_env == False:
        print(f'❌ Erro ao enviar email para:{email_enviar} referente a {nome_enviar}')
    else:
        print(f'✔️ Email enviado para:{email_enviar} referente a {nome_enviar}')
    bol_enviar = False

    scrp_tabela = scrp_cabec_tabela
