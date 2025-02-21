import json
import logging
import smtplib
import time
import tkinter as tk
from datetime import datetime, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import ttk

import gspread
import matplotlib.pyplot as plt
import pandas as pd
import pytz
import requests
import schedule
from oauth2client.service_account import ServiceAccountCredentials

import config

# Configurações de log
logging.basicConfig(filename='log.log', level=logging.INFO)

# Configurações Google Sheets
scopes = ['https://spreadsheets.google.com/feeds', 
          'https://www.googleapis.com/auth/drive'
]
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    filename=config.filename,
    scopes=scopes
    )
client = gspread.authorize(credentials)

# Autenticação com Google Sheets
planilha_completa = client.open(
    title=config.title,
    folder_id=config.folder_id
    )
planilha = planilha_completa.get_worksheet(0)

def iniciar_processo():
    print("\n***************** INICIO DO PROCESSO *****************\n")

def encerrar_processo():
    print("\n***************** ✅ Processo concluído! *****************\n")

################################################################
# PROCESSO
def verificarNovasLinhas():
    try:
        dados = planilha.get_all_values()

        if len(dados) <= 1:
            print("\n❌ Nenhuma nova linha encontrada na planilha.")
            return []

        header = dados[0]

        if "Processado" not in header:
            print("⚠️ Atenção: Adicione uma coluna chamada 'Processado' na planilha.")
            return []

        col_index = header.index("Processado")  # Índice da coluna "Processado"

        # Filtra apenas as linhas novas (onde "Processado" está vazio)
        novas_linhas = [linha for linha in dados[1:] if len(linha) > col_index and linha[col_index] == ""]

        if novas_linhas:
            print(f"\n========================= GOOGLE SHEETS =========================")
            print(f"🔍 Encontradas {len(novas_linhas)} novas linhas para processamento.")
            return novas_linhas  # Retorna apenas novas linhas

        print("\n🚫  Nenhuma nova linha para processar no momento.\n")
        return []

    except Exception as e:
        print(f"❌ Erro ao verificar novas linhas: {e}")
        return []

    
def verificar_eProcessar():
    novas_linhas = verificarNovasLinhas()
    if novas_linhas:
        processar_linhas(novas_linhas)
        enviarEmail()
    
def processar_linhas(linhas):
        if not linhas:
            return
        
        if not schedule.jobs:
            show_table()
        
        # Verifique se houver linhas para processar
        linhas_processadas = set()
        for index, linha in enumerate(linhas, start=1):
            print("================= DADOS DA LINHA =================\n")
            linha_str = str(linha)
            
            if linha_str in linhas_processadas:
                continue 
            
            linhas_processadas.add(linha_str)

            # Verifique se o Tipo do projeto
            get_tipo(linha)
                
            # Verifique se o Status de projeto existe
            get_Status(linha)

            # Crie card no Trello e informe prazo com base na planilha
            post_CardTrello(linha)

            # Atualizar a planilha para marcar a linha como "Processado"
            marcarComoProcessado(linha)

            print(f"⚙️  Processando linha {index}/{len(linhas)}: {linha_str}")

def marcarComoProcessado(novas_linhas):
    dados = planilha.get_all_values()  # Obtém todos os valores da planilha

    if not dados:
        print("Erro: A planilha está vazia ou não foi carregada corretamente!")
        return

    cabecalho = [col.strip() for col in dados[0]]  # Remove espaços extras dos nomes das colunas

    # Verifica se as colunas esperadas existem
    if "Processado" not in cabecalho:
        print("Erro: A planilha não contém as colunas necessárias ('Processado').")
        print(f"Colunas encontradas: {cabecalho}")
        return

    col_processado = cabecalho.index("Processado") + 1  # Encontra o índice correto da coluna "Processado"
    col_tarefa = cabecalho.index("Tarefa") + 1  # Encontra o índice correto da coluna "Tarefa"

    for i, row in enumerate(dados[1:], start=2):  # Começa da segunda linha (dados reais)
        if row[col_tarefa - 1] in novas_linhas:  # Se a tarefa está na lista de novas
            planilha.update_cell(i, col_processado, "Sim")  # Marca como processado

    print("✅ Linhas marcadas como processadas no Google Sheets e Trello.")

def agendar_verificacao():
    # Agendar a verificação periódica
    schedule.every(1).minutes.do(get_novaLinha)  # Verificar a cada 5 minutos

#########################################################
# API Google Sheets
def get_novaLinha():
    try:
        dados = planilha.get_all_values()

        if len(dados) <= 1:
            print("❌ Nenhuma nova linha encontrada na planilha.")
            return []

        header = dados[0]

        if "Processado" not in header:
            print("⚠️ Atenção: Adicione uma coluna chamada 'Processado' na planilha.")
            return []

        col_index = header.index("Processado")  # Pega o índice da coluna "Processado"

        # 🔍 Filtra apenas as linhas onde "Processado" está vazio
        novas_linhas = [linha for linha in dados[1:] if len(linha) > col_index and linha[col_index].strip() == ""]

        if novas_linhas:
            print(f"🔍 {len(novas_linhas)} novas linhas encontradas para processamento.")
            return novas_linhas
        return []

    except Exception as e:
        print(f"❌ Erro ao verificar novas linhas: {e}")
        return []

def get_tipo(linha):
    print(f"🔬 O projeto é do tipo: {linha[2]}")
    tipo = linha[2]  # Obtém o valor da coluna "Tipo" apenas para esta linha
    return tipo

def get_Status(linha):
    status = linha[3]  # Obtém o valor da coluna "Status"
    print(f"📌 Inserir projeto no card {status}.")
        
def show_table(): # Usado apenas para debug 
    dados = planilha.get_all_values()
    df = pd.DataFrame(dados)

    # Criar nova janela apenas quando necessário
    window = tk.Tk()
    window.title("Tabela de Dados")
    
    frame = ttk.Frame(window)
    frame.pack(fill="both", expand=True)

    tree = ttk.Treeview(frame, columns=list(df.columns), show="headings")
    
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)

    for row in df.itertuples(index=False):
        tree.insert("", "end", values=row)

    tree.pack(fill="both", expand=True)
    
    # Fechar a janela após 5 segundos
    window.after(5000, window.destroy)
    window.mainloop()

#########################################################
# API Trello
def post_CardTrello(linha):
    # Criar Trello
    projeto_name = linha[0]  # Nome do projeto
    tarefa_name = linha[1]  # Nome da tarefa
    status = linha[3]  # Status da tarefa
    data = linha[4]  # Data da tarefa

    if status in config.status_map:
        card_description = f"Essa tarefa faz parte do projeto: {projeto_name}"

    # Ajustar o fuso horário para UTC
    br_tz = pytz.timezone("America/Sao_Paulo")  # Define fuso do Brasil
    due_date = datetime.strptime(data, "%d/%m/%Y")  # Converte a data da string
    due_date = br_tz.localize(due_date)  # Aplica fuso horário local
    due_date = due_date.astimezone(pytz.utc)  # Converte para UTC
    due_date = due_date.strftime("%Y-%m-%dT%H:%M:%SZ")  # Formata para ISO 8601

    createCardEndpoint = config.mainTrelloEndpoint + "cards"
    jsonObj = {
        "key": config.key,
        "token" : config.token,
        "idList": config.status_map[status],
        "name": tarefa_name,
        "desc": card_description,
        "due": due_date  
    }

    new_card = requests.post(createCardEndpoint, params=jsonObj)

    card_data = json.loads(new_card.text)
    print(f"""🆔 ID Card: {card_data['id']}
📛 Titulo: {card_data['name']}
📆 Prazo: {card_data['due']}""")

#########################################################
# API Gmail
def enviarEmail():
    try:
        # Faça um for para percorrer todas as linhas
        dados = planilha.get_all_values()
        projeto = dados[1][0]  # Nome do projeto (linha 2, coluna 1)
        tarefa = dados[1][1]  # Nome da tarefa (linha 2, coluna 2)
        prazo = dados[1][4]  # Data de entrega (linha 2, coluna 4)

        print("\n================= ENVIANDO EMAIL =================\n")
        email = MIMEMultipart(config.emailFrom)
        email["From"] = config.emailFrom
        email["To"] = config.emailTo
        email["Subject"] = f"🚨 Novas Demandas de Projetos!"

        body = f"""
        <html>
          <body>
            <p>Olá, tudo bem?</p>
            <p>Precisamos entregar algumas tarefas que estão pendentes.</p>
            <p>Acione o time de qualidade para acompanhar o andamento das tarefas.</p>
            <p>Verifique o link: <a href="https://trello.com/b/IDoYrbpv/teste-criacao-de-backlog">https://trello.com/b/IDoYrbpv/teste-criacao-de-backlog</a></p>
            <p>Qualquer duvida, não hesite em entrar em contato com o time de qualidade.</p>
            <p>Obrigado!</p>
            <p>Vinícius Franco,<br>InnovAI. 🤖</p>
          </body>
        </html>
        """

        email.attach(MIMEText(body, "html"))
        
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()
        s.login(config.emailFrom, config.senha)
        s.sendmail(config.emailFrom, config.emailTo, email.as_string())
        s.quit()

        print("🟢 Email enviado com sucesso!")
        encerrar_processo()
    except smtplib.SMTPAuthenticationError:
        print("❌ Erro: Não foi possível autenticar no servidor SMTP. Verifique o e-mail e a senha configurados.")
    except ValueError as ve:
        print(f"❌ Erro de configuração: {ve}")
    except Exception as e:
        print(f"❌ Erro inesperado ao enviar o e-mail: {e}")

#########################################################
# Testando conexão com o servidores
def testandoConexaoSMTP():
    try:
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()
        s.login(config.emailFrom, config.senha)
        print("🟢 Conexão com o servidor SMTP estabelecida com sucesso!")
    except smtplib.SMTPAuthenticationError:
        print("❌ Erro: Não foi possível autenticar no servidor SMTP. Verifique o e-mail e a senha configurados.")
    except Exception as e:
        print(f"❌ Erro inesperado ao conectar ao servidor SMTP: {e}")

def testandoConexaoTrello():
    try:
        username = "vinicius_franco_"
        response = requests.get(f"https://api.trello.com/1/members/me/boards", params={"key": config.key, "token": config.token})
        if response.status_code == 200:
            print("🟢 Conexão com o servidor Trello estabelecida com sucesso!")
        else:
            print(f"❌ Erro: Não foi possível conectar ao servidor Trello. Código de status: {response.status_code}")
    except Exception as e:
        print(f"❌ Erro inesperado ao conectar ao servidor Trello: {e}")

def testandoConexaoGoogleSheets():
    try:
        scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(config.filename, scopes=scopes)
        client = gspread.authorize(creds)
        planilha = client.open(config.title).sheet1
        print("🟢 Conexão com o servidor Google Sheets estabelecida com sucesso!")
    except Exception as e:
        print(f"❌ Erro inesperado ao conectar ao servidor Google Sheets: {e}")

def testandoConexaoGmail():
    try:
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()
        s.login(config.emailFrom, config.senha)
        print("🟢 Conexão com o servidor Gmail estabelecida com sucesso!")
    except smtplib.SMTPAuthenticationError:
        print("❌ Erro: Não foi possível autenticar no servidor Gmail. Verifique o e-mail e a senha configurados.")
    except Exception as e:
        print(f"❌ Erro inesperado ao conectar ao servidor Gmail: {e}")

def conectar_servidores():
    print("⌛ Iniciando conexão com o servidor SMTP...")
    testandoConexaoSMTP()
    time.sleep(3)
    
    print("\n⌛ Iniciando conexão com o servidor Trello...")
    testandoConexaoTrello()
    time.sleep(3)
    
    print("\n⌛ Iniciando conexão com o servidor Google Sheets...")
    testandoConexaoGoogleSheets()
    time.sleep(3)
    
    print("\n⌛ Iniciando conexão com o servidor Gmail...")
    testandoConexaoGmail()
    time.sleep(3)
