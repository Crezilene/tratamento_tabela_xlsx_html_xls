import pandas as pd
from datetime import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import shutil

email_from = 'e-mail@email.com'
email_password = '*********'
email_to = ['e-mail@email.com',
            'e-mail@email.com', 'e-mail@email.com', 'e-mail@email.com']
email_subject = 'Status do Processamento de Arquivo de Suspensão'

folder_path = 'caminho'

files = [os.path.join(folder_path, f) for f in os.listdir(
    folder_path) if f.endswith('.xls') or f.endswith('.xlsx')]

latest_file = max(files, key=os.path.getctime)

try:
    df = None
    if latest_file.endswith('.xls'):
        try:
            df = pd.read_excel(latest_file, engine='xlrd')
        except Exception as e:
            print(f"Erro ao ler arquivo .xls com xlrd: {
                  e}. Tentando ler como HTML...")
            try:
                # Lendo o arquivo HTML, ignorando a primeira linha e usando a segunda linha como cabeçalho
                # header=1 usa a segunda linha como cabeçalho
                dfs = pd.read_html(latest_file, header=1)
                if len(dfs) > 0:
                    df = dfs[0]  # Pega a primeira tabela encontrada
                else:
                    print("Nenhuma tabela encontrada no arquivo HTML.")
                    df = pd.DataFrame()  # Definir como DataFrame vazio para enviar notificação
            except ValueError as ve:
                print(f"Erro ao processar o arquivo como HTML: {ve}")
                df = pd.DataFrame()  # Definir como DataFrame vazio para enviar notificação
            except Exception as e:
                print(f"Erro ao ler arquivo HTML: {e}")
                df = pd.DataFrame()  # Definir como DataFrame vazio para enviar notificação
        if not df.empty:
            new_file = latest_file.replace('.xls', '.xlsx')
            df.to_excel(new_file, index=False, engine='openpyxl')
            print(f"Arquivo convertido para: {new_file}")
            latest_file = new_file  # Atualiza para o novo arquivo .xlsx
    else:
        df = pd.read_excel(latest_file, engine='openpyxl')

except Exception as e:
    print(f"Erro ao ler arquivo Excel: {e}")
    raise

if df is None or df.empty:
    print("O arquivo está vazio ou nenhuma tabela foi encontrada.")
    email_subject = 'Tabela Vazia - Arquivo de Suspensão'
    html_body = """
    <html>
    <head></head>
    <body style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p>Olá, prezados(as)</p><br>
        <p>O arquivo processado está vazio ou nenhuma tabela foi encontrada.</p>
        <p>Por favor, verifique o arquivo de origem ou entre em contato para mais informações.</p><br>
        <p>Atenciosamente,<br>
        Alguem</p>
    </body>
    </html>
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = ', '.join(email_to)
        msg['Subject'] = email_subject
        msg.attach(MIMEText(html_body, 'html'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_from, email_password)
        text = msg.as_string()
        server.sendmail(email_from, email_to, text)
        server.quit()

        print("Email de notificação enviado com sucesso!")

    except Exception as e:
        print(f"Falha ao enviar email de notificação: {e}")
else:
    print("Arquivo lido com sucesso!")

    print("Colunas disponíveis no arquivo:")
    print(df.columns)

    if 'Workflow: Curso (Nome)' in df.columns:
        df = df[df['Workflow: Curso (Nome)'] ==
                'CUMPRIMENTO PROVISÓRIO PADRÃO']
    else:
        print("A coluna 'Workflow: Curso (Nome)' não foi encontrada no arquivo.")

    # Inserindo a coluna "Status" após "Cód. Causa"
    df.insert(1, 'Status', '')

    # Usando .loc para evitar o SettingWithCopyWarning
    df.insert(2, 'Responsavel', '')

    # Substituindo '[NÃO INFORMADO]' por '[CONFERIR]' na coluna 'Conteúdo'
    df['Conteúdo'] = df['Conteúdo'].replace('[NÃO INFORMADO]', '[CONFERIR]')

    # Verificando duplicatas na coluna 'Cód. Causa' e marcando '[CONFERIR]' na coluna 'Conteúdo' onde houver duplicação
    if 'Cód. Causa' in df.columns:
        df.loc[df.duplicated(subset=['Cód. Causa'], keep=False),
               'Conteúdo'] = '[CONFERIR]'
    else:
        print("A coluna 'Cód. Causa' não foi encontrada no arquivo.")

    # Renomeando a coluna 'Conteúdo' para 'Conta'
    df.rename(columns={'Conteúdo': 'Conta'}, inplace=True)

    if 'Workflow: Término Prev.' in df.columns:
        df['Workflow: Término Prev.'] = pd.to_datetime(
            df['Workflow: Término Prev.'])
        df = df.sort_values(by='Workflow: Término Prev.')
    else:
        print("A coluna 'Workflow: Término Prev.' não foi encontrada no arquivo.")

    total_rows = len(df)
    third_rows = total_rows // 3

    df.loc[df.index[:third_rows], 'Responsavel'] = 'Fulana de Tal'
    df.loc[df.index[third_rows:2*third_rows],
           'Responsavel'] = 'Fulano'
    df.loc[df.index[2*third_rows:], 'Responsavel'] = 'Tal de Fulana'

    print("Itens da coluna 'Cód. Causa':")
    print(df['Cód. Causa'])

    data_atual = datetime.now().strftime('%Y-%m-%d')
    output_folder = 'caminho_saida'
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    output_file = os.path.join(output_folder, f'suspensao_{data_atual}.xlsx')

    # Removendo o arquivo existente, se houver
    if os.path.exists(output_file):
        os.remove(output_file)

    df.to_excel(output_file, index=False)

    print(f"Arquivo salvo como: {output_file}")

    try:
        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = ', '.join(email_to)
        msg['Subject'] = email_subject

        # Corpo do e-mail personalizado
        html_body = """
    <html>
    <head></head>
    <body style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p>Olá, prezados(as)</p><br>
        <p>O arquivo está anexado e pronto para inclusão do <strong>Status</strong>.</p>
        <p>Qualquer dúvida, estamos à disposição.</p><br>
        <p>Atenciosamente,<br>
        Alguem</p>
    </body>
    </html>
    """
        msg.attach(MIMEText(html_body, 'html'))

        # Anexando o arquivo
        with open(output_file, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition', f'attachment; filename={os.path.basename(output_file)}')
            msg.attach(part)

        # Configuração do servidor SMTP
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_from, email_password)
        text = msg.as_string()
        server.sendmail(email_from, email_to, text)
        server.quit()

        print("Email enviado com sucesso!")

        # Movendo o arquivo para a nova pasta
        shutil.move(output_file, os.path.join(
            output_folder, os.path.basename(output_file)))
        print(f"Arquivo movido para: {output_folder}")

    except Exception as e:
        print(f"Falha ao enviar email: {e}")
