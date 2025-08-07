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
        

def obter_usuarios():
    try:
        conn = mysql_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT Nome FROM QuestRH_Pessoas")
        resultados = cursor.fetchall()
        conn.close()
        return sorted([row[0] for row in resultados])
    except Exception as e:
        print("Erro ao conectar ao banco de dados:", e)
        return []

def captura_valor_nota(Participante, Avaliador):
    try:
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        cursor.execute(f"SELECT SUM(Resposta) as Soma, COUNT(Resposta) as Contagem FROM QuestRH_Respostas where Participante = '{Participante}' and Nome_Avaliador = '{Avaliador}'")
        resultados = cursor.fetchall()
        conn.close()
        #print(resultados[0]['Soma'])
        try:
            valor_nota = resultados[0]['Soma']//resultados[0]['Contagem']
        except Exception as e:
            valor_nota = ''
        return valor_nota
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
        return 'Pendente'

def valida_texto(texto):
    NaoPermitidos = f"SELECT,DELETE,INSERT,',%,{chr(34)},TRUNCATE,DROP,JOIN"
    palavras = NaoPermitidos.split(",")

    for palavra in palavras:
            if palavra in texto.upper():
                return False
    
    return True
    
    
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





usuarios = obter_usuarios()

###########################################################################################################################################      
############################################################ Inicio da Aplicação ##########################################################
########################################################################################################################################### 

def main(page: ft.Page):
    codigo_enviado = ""
    nome_logado = ""
    page.title = "Sistema de Avaliação ENIND"
    page.scroll = ft.ScrollMode.AUTO
    page.bgcolor = ft.Colors.WHITE
    page.window.maximized = True
    data_limite = '2025/09/30 23:59:59'
    #data_limite = '2025/07/21 23:59:59'
    file_picker = ft.FilePicker()
    page.overlay.append(file_picker)

    # Container de alerta
    alerta_container = ft.Container(
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
    #email_txt = ft.TextField(label="Email", expand=True, visible=False)
    #cpf_txt = ft.TextField(label="CPF", expand=True, visible=False)
    senha_txt =  ft.TextField(label="Digite a senha para entrar", expand=True , visible=False, password=True)
    expiracao_txt = ft.Text(f'Data para envio de formulário expirado no dia {data_limite}.',expand= True, visible= False, size=20, weight=ft.FontWeight.BOLD, bgcolor=ft.Colors.RED_400, text_align=ft.alignment.center)

    mensagem_aguarde = ft.Text(
                    "Aguarde, atualizando o relatório...",
                    size=18,
                    weight=ft.FontWeight.BOLD,
                    text_align=ft.TextAlign.CENTER,
                    color=ft.Colors.WHITE
                )

    ############################# Campo de Data ################################
     # Gera listas de opções
    
    dias = [ft.dropdown.Option(str(d).zfill(2)) for d in range(1, 32)]
    meses = [
        ft.dropdown.Option("01", "Janeiro"),
        ft.dropdown.Option("02", "Fevereiro"),
        ft.dropdown.Option("03", "Março"),
        ft.dropdown.Option("04", "Abril"),
        ft.dropdown.Option("05", "Maio"),
        ft.dropdown.Option("06", "Junho"),
        ft.dropdown.Option("07", "Julho"),
        ft.dropdown.Option("08", "Agosto"),
        ft.dropdown.Option("09", "Setembro"),
        ft.dropdown.Option("10", "Outubro"),
        ft.dropdown.Option("11", "Novembro"),
        ft.dropdown.Option("12", "Dezembro")
    ]
    anos = [ft.dropdown.Option(str(y)) for y in range(1940, datetime.datetime.now().year + 1)]

    campo_data = ft.TextField(label="Data selecionada", read_only=True, expand=True)

    dia_dropdown = ft.Dropdown(label="Dia", options=dias)
    mes_dropdown = ft.Dropdown(label="Mês", options=meses)
    ano_dropdown = ft.Dropdown(label="Ano", options=anos)

    def atualizar_data(e):
        if dia_dropdown.value and mes_dropdown.value and ano_dropdown.value:
            campo_data.value = f"{ano_dropdown.value}/{mes_dropdown.value}/{dia_dropdown.value}"
            page.update()
    
            
    # Conecta todos os eventos
    dia_dropdown.on_change = atualizar_data
    mes_dropdown.on_change = atualizar_data
    ano_dropdown.on_change = atualizar_data


    data_row = ft.Row(
        [dia_dropdown, mes_dropdown, ano_dropdown, campo_data] ,
        visible=False,
        alignment=ft.MainAxisAlignment.START,
        spacing=10
    )

    
    
    erro_login = ft.Text("", color=ft.Colors.RED)
    codigo_input = ft.TextField(label="Digite o código recebido", expand=True)
    txt_observacoes = ft.TextField(label="Observações", expand=True, multiline=True,max_length=500)

    form_inputs = []

    # Elemento de saudação
    texto_ola = ft.Text("", size=15, weight=ft.FontWeight.BOLD)
    texto_ola1 = ft.Text("", size=15, weight=ft.FontWeight.BOLD)
    texto_ola2 = ft.Text("", size=15, weight=ft.FontWeight.BOLD)

    participante_realizado = ft.Text("", size=15, weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER, color=ft.Colors.GREY_800)

    ####################################################################################################################################### 
    ######################################################## Funções dos Objetos ##########################################################
    ####################################################################################################################################### 
    last_interaction = {"time": time.monotonic()}

    def reset_idle_time(e=None):
        last_interaction["time"] = time.monotonic()
        deslog_overlay.visible = False
        page.update()

    def voltar_login(e):
        email_validacao_view.visible =False
        login_view.visible =True
        page.update()

    def voltar_painel(e):
        painel_view.visible = True
        formulario_view.visible = False
        formulario_Envio.visible = False
        questionarios = obter_questionarios(nome_cb.value)
        montar_tabela(questionarios)
        page.update()

    def atualiza_rel(e):
        aguarde_overlay.visible = True
        page.update()

        questionarios = obter_questionarios(nome_cb.value)
        montar_tabela(questionarios)
        registra_login(nome_cb.value)
        exportar_btn.visible = False
        
        painel_view.visible = True

        aguarde_overlay.visible = False
        page.update()

    # Função chamada ao clicar "Sim" ou "Não"
    def confirmar_continuacao(e, resposta):
        confirmacao_overlay.visible = False
        if resposta:
            enviar_formulario(e)
        else:
            mostrar_alerta_temporario('Cancelamento de resposta realizado com sucess', ft.Colors.GREEN_400)
        page.update()

    # Botão principal para exibir o overlay
    def mostrar_confirmacao(e):
        confirmacao_overlay.visible = True
        page.update()

    def ao_selecionar_nome(e):
        texto_ola.value = f'Olá, {nome_cb.value}'
        texto_ola1.value = f'Olá, {nome_cb.value}'
        texto_ola2.value = f'Olá, {nome_cb.value}'

        senha_txt.value = ''
        painel_view.visible = False
        page.update()

        try:
            conn = mysql_connection()
            cursor = conn.cursor(pymysql.cursors.DictCursor)  # Retorna os resultados como dicionário
            scrp_sql = "SELECT * FROM QuestRH_Pessoas WHERE Nome = %s"
            cursor.execute(scrp_sql, (nome_cb.value,))
            resultado = cursor.fetchone()
            
            if resultado:
                email = resultado['Email']
                if email:  # Se tem email (não é None nem vazio)
                    senha_txt.visible = False
                else:
                    if nome_cb.value != 'Administrador':
                        senha_txt.visible = False
                    else:
                        senha_txt.visible = True

            elif nome_cb.value != 'Administrador':
                senha_txt.visible = False
            else:
                senha_txt.visible = True

        except Exception as ex:
            mostrar_alerta_temporario(f'erro em tempo de execução: {ex}', ft.Colors.RED_400)
        finally:
            if conn:
                conn.close()
        
        page.update()
        # Aqui você pode preencher campos automaticamente ou carregar dados do banco

    nome_cb = ft.Dropdown(
        label="Nome",
        options=[ft.dropdown.Option(n) for n in usuarios],
        expand=True,
        on_change=ao_selecionar_nome
    )

    
    # Elemento de avaliação
    nome_em_avaliacao = ft.Text("", size=25, weight=ft.FontWeight.BOLD)

    
    codigo_input = ft.TextField(label="Digite o código recebido", expand=True)

    def gerar_codigo():
        return ''.join(random.choices(string.digits, k=6))

    def reenviar_email(e):
        nonlocal codigo_enviado
        codigo_enviado = gerar_codigo()
        #texto_ola.value = f'Olá, {nome_cb.value}'
        corpo = prepara_corpo_email_Codigo(codigo_enviado)
        envio, erro = enviar_email(email_txt.value,'Confirmação de Identidade - Questionário ENIND',corpo)

        if envio == False:
            mostrar_alerta_temporario(f'Houve um erro ao enviar o email: {erro}!',ft.Colors.RED_400)
            return
        
        mostrar_alerta_temporario('Email enviado com sucesso!',ft.Colors.GREEN_700)
        page.update()

    #Função para validar o código do email
    def validar_codigo_email(e):
        if codigo_input.value == codigo_enviado:
            email_validacao_view.visible = False
            aguarde_overlay.visible = True
            page.update()
            questionarios = obter_questionarios(nome_cb.value)
            montar_tabela(questionarios)
            registra_login(nome_cb.value)
            painel_view.visible = True
            aguarde_overlay.visible = False
            page.update()
        else:
            mostrar_alerta_temporario('Código não coincide, por gentileza, tente novamente',ft.Colors.RED_400)
        page.update()
    
    
    #Função para validar o Login
    def validar_login(e):
        #texto_ola.value = f"Olá, {nome_cb.value}"
        if nome_cb.value == 'Administrador':
            if senha_txt.value == "RH*123":
                texto_ola.value = f"Olá, Administrador"
                aguarde_overlay.visible = True
                page.update()
                questionarios = obter_questionarios('Administrador')
                montar_tabela(questionarios)
                registra_login(nome_cb.value)
                exportar_btn.visible = True
                painel_view.visible = True
                aguarde_overlay.visible = False
                page.update()
                return
            else:
                mostrar_alerta_temporario('Senha incorreta',ft.Colors.RED_400)
                return
        if not nome_cb.value:
            mostrar_alerta_temporario("Defina antes com quem deseja realizar o login.",ft.Colors.RED_400)
            return
        if  email_txt.visible == True:
            if ("@enind.com.br" in email_txt.value) or ("@enindservicos.com.br" in email_txt.value):
                erro_login.value = ""
                codigo_input.value = ""
                login_view.visible = False
                texto_ola.value = f"Olá, {nome_cb.value}"
                exportar_btn.visible = False
                email_validacao_view.visible = True
                page.update()
            else:
               mostrar_alerta_temporario("E-mail inválido, é necessário ser da ENIND",ft.Colors.RED_400)
               return
        else:
            try:
                conn = mysql_connection()
                cursor = conn.cursor(pymysql.cursors.DictCursor)
                scrp_sql = f"SELECT * FROM QuestRH_Pessoas WHERE Nome = '{nome_cb.value}'"
                #print(scrp_sql)
                cursor.execute(scrp_sql)
                resultado = cursor.fetchone()
                conn.close()

                if resultado['CPF'] == cpf_txt.value and datetime.datetime.strftime(resultado['Data_Nasc'],'%Y/%m/%d') == campo_data.value:
                    erro_login.value = ""
                    login_view.visible = False
                    #texto_ola.value = f"Olá, {nome_cb.value}"
                    aguarde_overlay.visible = True
                    page.update()
                    questionarios = obter_questionarios(nome_cb.value)
                    montar_tabela(questionarios)
                    registra_login(nome_cb.value)
                    exportar_btn.visible = False
                    painel_view.visible = True
                    aguarde_overlay.visible =  False
                    page.update()
                else:
                    mostrar_alerta_temporario("CPF e data de nascimento não coincidem", ft.Colors.RED_400)
                    return
                
            except Exception as ex:
                erro_login.value = f"Erro na validação: {ex}"
                print(f"Erro na validação: {ex}")

        page.update()

   

    tabela = ft.DataTable(columns=[
        ft.DataColumn(label=ft.Text("Nome")),
        ft.DataColumn(label=ft.Text("Questionário")),
        ft.DataColumn(label=ft.Text("Avaliador")),
        ft.DataColumn(label=ft.Text("Status")),
        ft.DataColumn(label=ft.Text("Auto Avaliação")),
        ft.DataColumn(label=ft.Text("Avaliação1")),
        ft.DataColumn(label=ft.Text("Avaliação2")),
        ft.DataColumn(label=ft.Text("Ação")),
    ])


    #Função para criar a tabela de visualização - Painel de controle
    def montar_tabela(questionarios):
        nonlocal data_limite
        data_atual = datetime.datetime.now()
        data_fechamento = datetime.datetime.strptime(data_limite,"%Y/%m/%d %H:%M:%S") 
        tabela.rows.clear()
        
        for q in questionarios:
            if q["status"] == "Pendente":
                bg_cor = ft.Colors.YELLOW_100
            elif q["status"] == "Realizado":
                bg_cor = ft.Colors.GREEN_100
            else:
                bg_cor = None

            btn_avaliar = (
                ft.TextButton("Avaliar", on_click=lambda e, nome=q["nome"]: abrir_formulario(nome))
                if (texto_ola.value != "Olá, Administrador") and (q["status"] == "Pendente") and (data_atual <= data_fechamento) and (q['Questionário'] != 'A3' or (q['Questionário'] == 'A3' and q['Avaliador'] != 0)) else ft.Text("")
            )
            tabela.rows.append(ft.DataRow(cells=[
                ft.DataCell(ft.Text(q["nome"])),
                ft.DataCell(ft.Text(q["Questionário"])),
                ft.DataCell(ft.Text(q["Avaliador"])),
                 ft.DataCell(
                    ft.Container(
                        content=ft.Text(q["status"]),
                        bgcolor=bg_cor,
                        padding=5,
                        border_radius=5,
                    )
                ),
                ft.DataCell(ft.Text(q["auto_aval"])),
                ft.DataCell(ft.Text(q["primaria"])),
                ft.DataCell(ft.Text(q["secundaria"])),
                ft.DataCell(btn_avaliar),
            ]))

        if data_atual >=data_fechamento:
            expiracao_txt.visible = True
            page.update()
   
    
    def muda_cor_dropdown(e: ft.ControlEvent):
        dropdown = e.control
        valor = dropdown.value
        container = dropdown.meu_container 

        # Exemplo de lógica: muda a cor conforme valor selecionado
        if valor == "1 - Insatisfatório":
            container.bgcolor = ft.Colors.RED_100
        elif valor == "2 - Regular":
            container.bgcolor = ft.Colors.ORANGE_100
        elif valor == "3 - Satisfatório":
            container.bgcolor = ft.Colors.YELLOW_100
        elif valor == "4 - Bom":
            container.bgcolor = ft.Colors.LIGHT_GREEN_100
        elif valor == "5 - Excelente":
            container.bgcolor = ft.Colors.GREEN_400
        else:
            container.bgcolor = ft.Colors.WHITE
        
        # Atualiza a interface
        container.update()

    #Função para construir os formulários de respostas
    def abrir_formulario(nome):
        nome_em_avaliacao.value = f'Você está avaliando: {nome}'
        form_inputs.clear()
        form_content.controls.clear()
        perguntas_formulario = obter_perguntas(nome)

        # Opções padrão
        opcoes_avaliacao = [
            ft.dropdown.Option("1 - Insatisfatório"),
            ft.dropdown.Option("2 - Regular"),
            ft.dropdown.Option("3 - Satisfatório"),
            ft.dropdown.Option("4 - Bom"),
            ft.dropdown.Option("5 - Excelente")
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

                '''
                linha_pergunta = ft.Row(
                    controls=[
                        obj_texto,
                        container1
                    ],
                    spacing=10,
                    alignment=ft.MainAxisAlignment.START,
                    wrap=True
                )
                '''

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


        painel_view.visible = False
        formulario_view.visible = True
        page.update()


    #Função para enviar as resposta dos usuários
    def enviar_formulario(e):
        respostas = []
        mensagem_aguarde.value = 'Aguarde, enviando respostas...'
        aguarde_overlay.visible = True
        page.update()

        for i, grupo in enumerate(form_inputs):
            if not grupo.value:
                mostrar_alerta_temporario("Preencha todos os campos antes de enviar o furmulário...", ft.Colors.RED_400)
                aguarde_overlay.visible = False
                page.update()
                return  # Interrompe envio
            else:
                valor_resp = str(grupo.value).split(' - ')[0]
                respostas.append({
                    'ID':i+1,
                    'resposta': valor_resp})

        obs_us = txt_observacoes.value
        if valida_texto(obs_us) == False:
            mostrar_alerta_temporario('Campo de obsevações constem palavaras não permitidas, por gentileza, analise o texto', ft.Colors.RED_400)
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
        
        try:
            usuario = os.getlogin()
        except:
            return os.environ.get("USERNAME")
        computador = platform.node()

        #Inserindo informações no banco de dados
        for row in respostas:

            campos = '''(Participante, Cargo, C_Custo,  Local, Avaliacao, Nome_Avaliador, ID_Rel, ID_Pergunta, Pilar, Competencia, Pergunta, Resposta, Observacao, Data_Resp, Computador, Login) 
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'''
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
        formulario_view.visible = False
        mensagem_aguarde.value = 'Aguarde, atualizando o relatório...'
        aguarde_overlay.visible = False
        page.update()
    
    def exportar_para_excel(pasta):
        # Simulação de dados para exportar
        conn = mysql_connection()
        cursor = conn.cursor(pymysql.cursors.DictCursor)
        scrp_sql = f"SELECT * FROM QuestRH_Respostas"
        cursor.execute(scrp_sql)
        consulta = cursor.fetchall()
        cursor.close()
        conn.close()

        df = pd.DataFrame(consulta)

        nome_arquivo = f"exportacao_respostas_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caminho_completo = os.path.join(pasta, nome_arquivo)

        try:
            df.to_excel(caminho_completo, index=False, engine='openpyxl')
            mostrar_alerta_temporario('Exportação realizada com sucesso', ft.Colors.GREEN_400)
        except Exception as e:
            mostrar_alerta_temporario('Não foi possível salvar o arquivo', ft.Colors.RED_400)

        page.update()

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
                    "Deseja realmente continuar com o cadastro da resposta dessa pessoa? Uma vez realizado, não é possível rever e reenviar o relatório",
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
    
    page.overlay.append(aguarde_overlay)
    page.overlay.append(confirmacao_overlay)

    login_view = ft.Container(
        content=ft.Column([
            ft.Text("Login", size=30, weight=ft.FontWeight.BOLD),
            nome_cb,
            email_txt,
            cpf_txt,
            data_row,
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

   

    email_validacao_view = ft.Container(
        content=ft.Column([
            ft.TextButton("Retornar", on_click=voltar_login),
            texto_ola,
            ft.Text(f"Validação de E-mail", size=25, weight=ft.FontWeight.BOLD),
            codigo_input,
            ft.TextButton("Reenviar Código", on_click=reenviar_email),
            ft.ElevatedButton("Validar E-mail", on_click=validar_codigo_email),
        ], spacing=15),
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

    file_picker.on_result = ao_escolher_pasta

    exportar_btn = ft.ElevatedButton(
        "Exportar para Excel",
        on_click=escolher_pasta,
        width=250,
        height=50,
        visible=False
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



    painel_view = ft.Container(
        content=ft.Column([
            texto_ola1,
            expiracao_txt,
            exportar_btn,
            ft.Text("Painel de Controle", size=25, weight=ft.FontWeight.BOLD),
            atualizar_btn,
            tabela
        ], spacing=15),
        bgcolor=ft.Colors.WHITE,
        border_radius=10,
        padding=20,
        opacity=0.92,
        visible=False
    )

    form_content = ft.Column(spacing=10,expand=True)
    formulario_view = ft.GestureDetector(
        on_pan_start=reset_idle_time,     # movimento de mouse ou dedo
        on_pan_update=reset_idle_time,    # movimento contínuo
        on_tap=reset_idle_time,           # clique
        content=ft.Container(
            content=ft.Column([
                ft.TextButton("Retornar para o painel de controle", on_click=voltar_painel),
                texto_ola2,
                nome_em_avaliacao,
                form_content,
                txt_observacoes,
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
            visible=False
        )
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
                        on_click=voltar_painel,
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
    
    


    imagem_fundo = ft.Container(
        content=ft.Image(
            src=caminho_recurso("Imagem_Quest.png"),
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
            email_validacao_view,
            painel_view,
            formulario_view,
            formulario_Envio
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=30),
        alignment=ft.alignment.center,
        expand=True
    )

    overlay = ft.Column([
        conteudo_central,
        alerta_container
        ], 
    expand=True,
    alignment=ft.alignment.bottom_center 
    )



    stack = ft.Stack([
        #imagem_fundo,
        overlay
    ])

   
    page.add(stack)

print('Aguarde, abrindo aplicativo, isso pode demorar até 60 segundos')
ft.app(target=main)


