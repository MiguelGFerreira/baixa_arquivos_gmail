import base64
import os
import json
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from dotenv import load_dotenv

load_dotenv()

empresa = ""

while empresa not in ['TCE', 'RKF']:
    empresa = input("Empresa (TCE ou RKF): ")
    empresa = empresa.upper()


API_KEY = os.getenv(f'API_KEY_{empresa}')
with open(f'credentials{empresa}.json') as credentials_file:
    parsed_json = json.load(credentials_file)


# Escopo necessário para acessar a API do Gmail
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def get_refresh_token():
    flow = InstalledAppFlow.from_client_secrets_file(
        f'credentials{empresa}.json', scopes=SCOPES)
    credentials = flow.run_console()
    return credentials.refresh_token


# Defina as credenciais da API do Gmail
creds = Credentials(
    API_KEY,
    refresh_token=get_refresh_token(),
    token_uri=parsed_json['installed']['token_uri'],
    client_id=parsed_json['installed']['client_id'],
    client_secret=parsed_json['installed']['client_secret']
)

# Crie um serviço do Gmail
service = build('gmail', 'v1', credentials=creds)

# Função para baixar os anexos
def download_attachments(message_id):
    message = service.users().messages().get(userId='me', id=message_id).execute()

    if message['internalDate'] > "1714878000":  # Pegando apenas emails a partir do dia que deu problema
        if 'parts' in message['payload']:
            for part in message['payload']['parts']:
                filename = part['filename']
                extension = filename.split('.')[-1]
                
                if filename and extension in ['pdf', 'xml']:
                    if 'data' in part['body']:
                        data = part['body']['data']
                    else:
                        att_id = part['body']['attachmentId']
                        att = service.users().messages().attachments().get(userId="me", messageId=message['id'],id=att_id).execute()
                        data = att['data']

                    file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                    path = f"{part['filename']}"

                    with open(path, 'wb') as f:
                        f.write(file_data)

                    print(f"Anexo {part['filename']} baixado com sucesso!")
        else:
            att_id = message['payload']['body']['attachmentId']
            att = service.users().messages().attachments().get(userId="me", messageId=message['id'],id=att_id).execute()
            data = att['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            path = f"{message['payload']['filename']}"

            with open(path, 'wb') as f:
                f.write(file_data)

            print(f"Anexo {message['payload']['filename']} baixado com sucesso!")

# Função para listar emails com anexos
def list_messages_with_attachments(query=''):
    response = service.users().messages().list(userId='me', q=query).execute()
    messages = []

    if 'messages' in response:
        messages.extend(response['messages'])

    for message in messages:
        download_attachments(message['id'])


if __name__ == "__main__":
    list_messages_with_attachments('has:attachment')
