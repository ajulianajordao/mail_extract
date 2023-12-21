import os
import logging
from exchangelib import Credentials, Account, DELEGATE, FileAttachment
import openpyxl
from datetime import datetime
from dotenv import load_dotenv
import pandas as pd
import easygui
from tkinter import *
from tkinter import ttk
from threading import Thread
from tqdm import tqdm

load_dotenv()

def detectar_tipo_anexo(nome_anexo):
    palavras_nf = ["nf", "nota fiscal", "invoice", "faturado"]
    palavras_certificado = ["certificado", "certification", "material certification"]

    if any(palavra in nome_anexo.lower() for palavra in palavras_nf):
        return "Nota Fiscal"
    elif any(palavra in nome_anexo.lower() for palavra in palavras_certificado):
        return "Certificação de Material"
    else:
        return "Outro"

def extrair_metadados_anexo(attachment):
    return {"Nome": attachment.name, "Tamanho": attachment.size}

def solicitar_opcoes_filtro():
    msg = "Escolha as opções de filtro:"
    title = "Opções de Filtro"
    choices = ["Caixa de Entrada", "Rascunhos", "Itens Enviados", "Itens Excluídos"]
    selected_options = easygui.multchoicebox(msg, title, choices)

    # Filtro opcional
    filter_subject = None
    
    # Sempre forneça pelo menos duas opções, mesmo que algumas sejam nulas
    additional_options = easygui.multchoicebox("Selecione as opções de filtro adicional:", title,
                                               ["Email Recebido", "Filtrar por Data"])
    
    if additional_options is not None:
        if "Filtrar por Data" in additional_options:
            start_date = easygui.enterbox("Data de início para filtrar e-mails (formato: YYYY-MM-DD):")
            end_date = easygui.enterbox("Data de término para filtrar e-mails (formato: YYYY-MM-DD):")
        else:
            start_date = end_date = None
    else:
        start_date = end_date = None

    return selected_options, filter_subject, start_date, end_date

def obter_credenciais():
    while True:
        email_address = easygui.enterbox("Digite seu endereço de e-mail:")
        password = easygui.passwordbox("Digite a senha do e-mail:", title="Login")

        credentials = Credentials(email_address, password)
        try:
            account = Account(email_address, credentials=credentials, autodiscover=True, access_type=DELEGATE)
            break
        except:
            msg = "Senha incorreta. Tente novamente?"
            if not easygui.ynbox(msg, title="Senha Incorreta", choices=("Sim", "Não")):
                return None, None

    # Salva as credenciais no arquivo .env
    with open('.env', 'w') as env_file:
        env_file.write(f"EMAIL_ADDRESS={email_address}\n")

    return email_address, password

def encontrar_nome_arquivo_disponivel(pasta, nome_base):
    contador = 1
    nome_arquivo = f"{nome_base}.xlsx"

    while os.path.exists(os.path.join(pasta, nome_arquivo)):
        contador += 1
        nome_arquivo = f"{nome_base}_version_{contador}.xlsx"

    return nome_arquivo

def extrair_emails(account, filter_subject, progress_var, log_data):
    # Convertendo o QuerySet em lista para calcular o total de e-mails
    total_emails = len(list(account.inbox.filter(subject__contains=filter_subject or '')))

    for index, item in tqdm(enumerate(account.inbox.filter(subject__contains=filter_subject or ''), start=1), total=total_emails, desc="Extraindo e-mails"):
        sender = item.sender.email_address
        to_recipients = ", ".join(recipient.email_address for recipient in item.to_recipients)
        cc_recipients = ", ".join(recipient.email_address for recipient in item.cc_recipients)
        subject = item.subject
        body = item.text_body
        received_date = item.datetime_received

        if item.has_attachments:
            for i, attachment in enumerate(item.attachments, start=1):
                attachment_name = attachment.name
                attachment_size = attachment.size

                tipo_anexo = detectar_tipo_anexo(attachment_name)
                metadados_anexo = extrair_metadados_anexo(attachment)

                # Atualizar a barra de progresso
                progress_var.set((index / total_emails) * 100)

        log_data["Subject"].append(subject)
        log_data["Status"].append("Completo")

    logging.info(f"Total de {len(log_data['Subject'])} e-mails processados com sucesso.")

def extrair_emails_thread(account, filter_subject, progress_var, log_data):
    thread = Thread(target=extrair_emails, args=(account, filter_subject, progress_var, log_data))
    thread.start()

def criar_interface_grafica():
    root = Tk()
    root.title("Progresso da Extração")

    # Configurar a barra de progresso
    progress_var = DoubleVar()
    progress = ttk.Progressbar(root, length=300, mode="determinate", variable=progress_var)
    progress.grid(row=0, column=0, padx=10, pady=10)

    # Configurar o botão Concluído
    btn_concluido = Button(root, text="Concluído", command=root.destroy)
    btn_concluido.grid(row=1, column=0, padx=10, pady=10)

    return root, progress_var

def main():
    email_address, password = obter_credenciais()

    if email_address is None or password is None:
        print("Credenciais não fornecidas. Certifique-se de inserir um endereço de e-mail e senha válidos.")
        return

    credentials = Credentials(email_address, password)
    account = Account(email_address, credentials=credentials, autodiscover=True, access_type=DELEGATE)

    selected_options, filter_subject, start_date, end_date = solicitar_opcoes_filtro()

    # Solicitar o caminho para salvar os arquivos
    data_mail_folder = easygui.diropenbox("Escolha a pasta para salvar os arquivos:")
    
    # Nome do arquivo
    data_atual = datetime.now().strftime("%d%m%Y")
    nome_base_arquivo = f"Compilados_emails_{data_atual}"
    
    # Encontrar um nome de arquivo disponível
    nome_arquivo = encontrar_nome_arquivo_disponivel(data_mail_folder, nome_base_arquivo)

    wb = openpyxl.Workbook()
    main_sheet = wb.active
    main_sheet.title = "E-mails"

    log_sheet = wb.create_sheet("Log de Atividades")

    main_sheet.append(["DE (From)", "PARA (To)", "CC (CC)", "TÍTULO DO E-MAIL (Email Subject)", "ANEXO (Attachment)",
                       "TEXTO DO ANEXO", "TEXTO (Body Text)", "DATA DE RECEBIMENTO", "TIPO DE ANEXO", "METADADOS DO ANEXO"])

    logging.info("Script iniciado.")

    log_data = {"Subject": [], "Status": []}

    extrair_emails_thread(account, filter_subject, None, log_data)

    root, progress_var = criar_interface_grafica()

    # Atualizar a barra de progresso enquanto a extração está em andamento
    root.after(100, lambda: extrair_emails_thread(account, filter_subject, progress_var, log_data))
    root.mainloop()

    # Salvar o arquivo com o nome encontrado
    wb.save(os.path.join(data_mail_folder, nome_arquivo))

    logging.info("Script concluído.")

if __name__ == "__main__":
    main()
