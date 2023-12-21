Extração de Dados de E-mails para Planilha Excel

Objetivo: 
O objetivo geral do código é extrair dados específicos de e-mails de uma conta da Microsoft, 
permitindo que o usuário defina critérios de filtragem. Esses dados são organizados e registrados
em uma planilha Excel, proporcionando uma análise eficiente das informações contidas nos e-mails. 
O script prioriza a segurança ao solicitar credenciais de forma segura e automatiza a classificação 
de anexos, facilitando o acompanhamento e a compreensão dos dados extraídos. Permite ao usuário definir 
critérios de filtragem para extrair e organizar informações específicas dos e-mails.

Funcionalidades Principais:
Autenticação Segura:

O script solicita o endereço de e-mail do usuário, bem como sua senha, 
proporcionando uma autenticação segura por meio da biblioteca keyring.

Configuração de Filtros:
Utiliza uma interface gráfica simples, fornecida pela biblioteca easygui, 
para permitir que o usuário defina opções de filtragem, como caixa de entrada, 
rascunhos, filtro por assunto e filtro por data.

Extração e Classificação de Anexos:

Analisa os e-mails da caixa de entrada, extrai informações relevantes, como remetente,
 destinatários, assunto, corpo do e-mail e anexos.

Classifica automaticamente os anexos em categorias, como "Nota Fiscal," "
Certificação de Material" ou "Outro."

Registro de Atividades:
Registra as atividades do script em um arquivo de log (script_log.txt), incluindo o início, término
e o status (Completo ou Falha) de cada e-mail processado.

Geração de Planilha Excel:
Organiza os dados extraídos em uma planilha Excel (.xlsx) com colunas específicas, facilitando a análise
e o acompanhamento das informações.

Nomeação Dinâmica de Arquivos:
Os arquivos da planilha são nomeados automaticamente com a data atual, garantindo a rastreabilidade das extrações.

Como Usar:
Configuração Inicial:
Clone este repositório em seu ambiente local.
Certifique-se de ter as dependências necessárias instaladas usando pip install -r requirements.txt.

Execução do Script:
Execute o script Python e siga as instruções fornecidas na interface gráfica para autenticação e definição de filtros.

Resultados:

Após a execução bem-sucedida, uma planilha Excel será gerada no diretório "Meus Documentos/Data_mail," 
contendo os dados extraídos e organizados.
Este projeto oferece uma solução eficiente para a extração
 e organização de dados de e-mails, tornando a análise dessas 
 informações mais acessível e prática..