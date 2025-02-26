import requests
import json
from io import BytesIO
from datetime import datetime
import pandas as pd
from dotenv import load_dotenv
import os

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

# Configurações do Microsoft Graph
client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
tenant_id = os.getenv('TENANT_ID')
scope = 'https://graph.microsoft.com/.default'
token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'

# Função para obter o token de acesso
def obter_token():
    response = requests.post(token_url, data={
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': scope
    })
    response.raise_for_status()
    return response.json()['access_token']

# Função para carregar dados da planilha do SharePoint
def carregar_dados(sharepoint_site, sharepoint_drive, sharepoint_file_path):
    token = obter_token()
    headers = {
        'Authorization': f'Bearer {token}'
    }
    url = f'https://graph.microsoft.com/v1.0/sites/{sharepoint_site}/drives/{sharepoint_drive}/root:/{sharepoint_file_path}:/content'
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    file_content = BytesIO(response.content)
    df = pd.read_excel(file_content, engine='openpyxl')
    return df

# Função para salvar dados na planilha do SharePoint
def salvar_dados(df, sharepoint_site, sharepoint_drive, sharepoint_file_path):
    token = obter_token()
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f'https://graph.microsoft.com/v1.0/sites/{sharepoint_site}/drives/{sharepoint_drive}/root:/{sharepoint_file_path}:/content'
    with BytesIO() as output:
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        response = requests.put(url, headers=headers, data=output)
    response.raise_for_status()