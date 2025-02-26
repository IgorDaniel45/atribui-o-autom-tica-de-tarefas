import requests
from Graph_API import obter_token, carregar_dados, salvar_dados
from atualizar_planilha import sharepoint_site_original, sharepoint_drive_original, sharepoint_file_path_original
from verificar_maquinas_preparadas import tecnicos, atualizar_responsavel

# Função para monitorar respostas de e-mail
def monitorar_respostas():
    token = obter_token()
    url = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    messages = response.json()['value']
    
    for msg in messages:
        if msg['from']['emailAddress']['address'] in tecnicos.keys() and 'Preparada' in msg['body']['content']:
            atualizar_planilha(tecnicos[msg['from']['emailAddress']['address']])
            concluir_tarefa_planner(msg['id'], token)

# Função para atualizar a planilha com o responsável designado
def atualizar_responsavel(colaborador, tecnico_email, task_id):
    df = carregar_dados(sharepoint_site_original, sharepoint_drive_original, sharepoint_file_path_original)
    tecnico_nome = tecnicos[tecnico_email]
    for index, row in df.iterrows():
        if row['colaborador'] == colaborador:
            df.at[index, 'responsavel'] = tecnico_nome
            df.at[index, 'task_id'] = task_id
            break
    salvar_dados(df, sharepoint_site_original, sharepoint_drive_original, sharepoint_file_path_original)

# Função para atualizar a planilha quando a máquina for preparada
def atualizar_planilha(tecnico_nome):
    df = carregar_dados(sharepoint_site_original, sharepoint_drive_original, sharepoint_file_path_original)
    for index, row in df.iterrows():
        if row['responsavel'] == tecnico_nome and not row['preparada']:
            df.at[index, 'preparada'] = True
            break
    salvar_dados(df, sharepoint_site_original, sharepoint_drive_original, sharepoint_file_path_original)

# Função para marcar tarefa como concluída no Planner
def concluir_tarefa_planner(task_id, token):
    url = f'https://graph.microsoft.com/v1.0/planner/tasks/{task_id}'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    body = {
        'percentComplete': 100
    }
    response = requests.patch(url, headers=headers, json=body)
    response.raise_for_status()

monitorar_respostas()