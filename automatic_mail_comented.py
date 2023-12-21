import os # Importa o módulo 'os', que fornece uma maneira de usar funcionalidades dependentes do sistema
# operacional, como ler ou escrever no sistema de arquivos.

import logging # Importa o módulo 'logging', que oferece um framework flexível para emitir mensagens de log
# a partir de programas Python.

from exchangelib import Credentials, Account, DELEGATE # Importa as classes 'Credentials', 
# 'Account' e 'DELEGATE' do módulo 'exchangelib', que são usadas para autenticação e interação com uma conta 
# de e-mail do Exchange.

import openpyxl # Importa o módulo 'openpyxl', que permite a manipulação de arquivos Excel (xlsx).

from datetime import datetime # Importa a classe 'datetime' do módulo 'datetime', que fornece funcionalidades 
# para trabalhar com datas e horas.

from dotenv import load_dotenv # # Importa a função 'load_dotenv' do módulo 'dotenv', que é usada para carregar variáveis 
# de ambiente de um arquivo '.env'.

import argparse  # Importa o módulo 'argparse', que facilita a análise de argumentos da linha de comando.

import pandas as pd # Importa a biblioteca 'pandas' e a renomeia para 'pd', sendo amplamente utilizada para manipulação e análise de dados.

import easygui # Importa o módulo 'easygui', que fornece uma interface gráfica simples para solicitar entrada do usuário.

import keyring  # Adicionado para lidar com senhas de forma segura

# Função para detectar o tipo de anexo (Nota Fiscal, Certificação de Material ou Outro)
def detectar_tipo_anexo(nome_anexo):
    # Lista de palavras-chave associadas a Nota Fiscal
    palavras_nf = ["nf", "nota fiscal", "invoice"]
    # Lista de palavras-chave associadas a Certificação de Material
    palavras_certificado = ["certificado", "certification", "material certification"]

    # Verifica se alguma palavra-chave de Nota Fiscal está presente no nome do anexo
    if any(palavra in nome_anexo.lower() for palavra in palavras_nf):
        return "Nota Fiscal"
    # Verifica se alguma palavra-chave de Certificação de Material está presente no nome do anexo
    elif any(palavra in nome_anexo.lower() for palavra in palavras_certificado):
        return "Certificação de Material"
    # Se não encontrar nenhuma correspondência, classifica como "Outro"
    else:
        return "Outro"

# Função para extrair metadados do anexo
def extrair_metadados_anexo(attachment):
    # Adapte conforme necessário para extrair metadados específicos do anexo
    return {"Nome": attachment.name, "Tamanho": attachment.size}

# Função para solicitar opções de filtro ao usuário por meio de interface gráfica
def solicitar_opcoes_filtro():
    # Mensagem exibida na caixa de diálogo
    msg = "Escolha as opções de filtro:"
    # Título da caixa de diálogo
    title = "Opções de Filtro"
    # Opções disponíveis para seleção
    choices = ["Caixa de Entrada", "Rascunhos", "Itens Enviados", "Itens Excluídos"]
    # Caixa de seleção múltipla para escolher as opções desejadas
    selected_options = easygui.multchoicebox(msg, title, choices)

    # Solicitação para inserir uma palavra-chave para filtrar por assunto
    filter_subject = easygui.enterbox("Filtrar por subject (palavra-chave):")

    # Mensagem e opções adicionais para o filtro
    msg = "Selecione as opções de filtro adicional:"
    title = "Opções de Filtro Adicional"
    choices = ["Email Enviado", "Email Recebido", "Filtrar por Data"]
    # Caixa de seleção múltipla para escolher opções adicionais
    additional_options = easygui.multchoicebox(msg, title, choices)

    # Inicialização de variáveis para datas de início e término do filtro por data
    start_date = end_date = None

    # Verifica se a opção "Filtrar por Data" foi selecionada
    if "Filtrar por Data" in additional_options:
        # Solicitação para inserir a data de início para filtrar e-mails
        start_date = easygui.enterbox("Data de início para filtrar e-mails (formato: YYYY-MM-DD):")
        # Solicitação para inserir a data de término para filtrar e-mails
        end_date = easygui.enterbox("Data de término para filtrar e-mails (formato: YYYY-MM-DD):")

    # Retorna as opções selecionadas e as informações de filtro
    return selected_options, filter_subject, additional_options, start_date, end_date

# Configuração do log para registrar as atividades do script
logging.basicConfig(filename='script_log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

# Obtenção de credenciais por meio de uma interface gráfica
email_address = easygui.enterbox("Digite seu endereço de e-mail:")

# Tentar recuperar a senha salva
password = keyring.get_password("MeuApp", email_address)

# Se a senha não estiver salva, solicitar ao usuário e salvá-la usando keyring
if password is None:
    master_password = easygui.passwordbox("Digite a senha mestra:")
    keyring.set_password("MeuApp", email_address, master_password)
    password = keyring.get_password("MeuApp", email_address)

# Configuração do parser para aceitar argumentos de linha de comando
parser = argparse.ArgumentParser(description="Script para extrair dados de e-mails e salvar em uma planilha do Excel.")
args = parser.parse_args()

# Autenticação na conta de e-mail
credentials = Credentials(email_address, password)
account = Account(email_address, credentials=credentials, autodiscover=True, access_type=DELEGATE)

# Solicitar opções de filtro ao usuário
selected_options, filter_subject, additional_options, start_date, end_date = solicitar_opcoes_filtro()

# Criação da pasta "Data_mail" no diretório padrão "Meus Documentos"
data_mail_folder = os.path.join(os.path.expanduser("~"), "Documents", "Data_mail")
# Cria o diretório se ele não existir
os.makedirs(data_mail_folder, exist_ok=True)

# Inicialização da planilha do Excel
wb = openpyxl.Workbook()
main_sheet = wb.active
main_sheet.title = "E-mails"

# Criação de uma segunda planilha para log de atividades
log_sheet = wb.create_sheet("Log de Atividades")

# Cabeçalho da planilha principal
main_sheet.append(["DE (From)", "PARA (To)", "CC (CC)", "TÍTULO DO E-MAIL (Email Subject)", "ANEXO (Attachment)",
                   "TEXTO DO ANEXO", "TEXTO (Body Text)", "DATA DE RECEBIMENTO", "TIPO DE ANEXO", "METADADOS DO ANEXO"])

# Registro do início do script no log
logging.info("Script iniciado.")

# Dicionário para armazenar dados do log
log_data = {"Subject": [], "Status": []}

try:
    # Loop para processar e-mails de acordo com os filtros
    for item in account.inbox.filter(subject__contains=filter_subject, datetime_received__range=(start_date, end_date)):
        sender = item.sender.email_address
        to_recipients = ", ".join(recipient.email_address for recipient in item.to_recipients)
        cc_recipients = ", ".join(recipient.email_address for recipient in item.cc_recipients)
        subject = item.subject
        body = item.text_body
        received_date = item.datetime_received  # Adicionando a data de recebimento

        if item.has_attachments:
            for i, attachment in enumerate(item.attachments, start=1):
                attachment_name = attachment.name
                attachment_size = attachment.size

                # Detectar o tipo de anexo
                tipo_anexo = detectar_tipo_anexo(attachment_name)

                # Extrair metadados do anexo
                metadados_anexo = extrair_metadados_anexo(attachment)

                # Adiciona informações à planilha principal
                main_sheet.append([sender, to_recipients, cc_recipients, subject,
                                   f"{i} - {attachment_name}", f"{i} - {attachment.content}", body, received_date,
                                   tipo_anexo, metadados_anexo])
        else:
            # Adiciona informações à planilha principal
            main_sheet.append([sender, to_recipients, cc_recipients, subject, "", "", body, received_date, "", ""])

        # Atualiza dados do log
        log_data["Subject"].append(subject)
        log_data["Status"].append("Completo")

    # Melhoria: Registra no log o total de e-mails processados com sucesso
    logging.info(f"Total de {len(log_data['Subject'])} e-mails processados com sucesso.")

except Exception as e:
    # Melhoria: Em caso de erro, registra no log e marca os e-mails correspondentes como erro
    logging.error(f"Erro durante a execução do script: {str(e)}")
    for subject in log_data["Subject"]:
        log_data["Status"].append(f"Falha - {str(e)}")

# Converte dados do log para DataFrame e adiciona à planilha de log
log_df = pd.DataFrame(log_data)
log_sheet.append(["Subject", "Status"])
for index, row in log_df.iterrows():
    log_sheet.append([row["Subject"], row["Status"]])

# Obtém a data atual para nomear o arquivo da planilha
data_atual = datetime.now().strftime("%d%m%Y")
nome_arquivo = f"Compilados_emails_{data_atual}.xlsx"

# Salva a planilha no diretório especificado
wb.save(os.path.join(data_mail_folder, nome_arquivo))

# Encerra a sessão da conta de e-mail
account.logout()

# Melhoria: Registra no log o término do script
logging.info("Script concluído.")