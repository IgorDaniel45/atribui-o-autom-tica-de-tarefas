import random
import requests
from Graph_API import obter_token, carregar_dados, salvar_dados
from atualizar_planilha import atualizar_planilha_com_dados_adicionais, sharepoint_site_original, sharepoint_drive_original, sharepoint_file_path_original, sharepoint_site_adicional, sharepoint_drive_adicional, sharepoint_file_path_adicional

tecnicos = {
    '',
    '',
    '',
    '',
    '',
}

planner_plan_id = ''

# Função para enviar alerta por e-mail
def enviar_alerta(maquina, tecnico_email, token):
    url = 'https://graph.microsoft.com/v1.0/me/sendMail'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    body = {
        'message': {
            'subject': 'Alerta: Verificação de Máquina',
            'body': {
                'contentType': 'Text',
                'content': f"A máquina para o colaborador {maquina['colaborador']} precisa ser preparada. Responda este e-mail com 'Preparada' quando a máquina estiver pronta.\n\nDetalhes do Colaborador:\nNome: {maquina['colaborador']}\nCargo: {maquina['cargo']}\nTelefone: {maquina['telefone']}\nEndereço: {maquina['endereco']}\nPerfil da Máquina: {maquina['perfil_maquina']}"
            },
            'toRecipients': [{'emailAddress': {'address': tecnico_email}}]
        }
    }
    response = requests.post(url, headers=headers, json=body)
    response.raise_for_status()

# Função para criar tarefa no Planner
def criar_tarefa_planner(maquina, tecnico_email, token, data_tarefa=None):
    url = f'https://graph.microsoft.com/v1.0/planner/plans/{planner_plan_id}/tasks'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    body = {
        'title': f"Preparar máquina para {maquina['colaborador']}",
        'assignments': {
            tecnico_email: {
                '@odata.type': 'microsoft.graph.plannerAssignment',
                'orderHint': ' !'
            }
        },
        'planId': planner_plan_id,
        'bucketId': 'SEU_BUCKET_ID',
        'details': {
            '@odata.type': 'microsoft.graph.plannerTaskDetails',
            'description': f"Detalhes do Colaborador:\nNome: {maquina['colaborador']}\nCargo: {maquina['cargo']}\nTelefone: {maquina['telefone']}\nEndereço: {maquina['endereco']}\nPerfil da Máquina: {maquina['perfil_maquina']}"
        }
    }
    if data_tarefa:
        body['dueDateTime'] = data_tarefa.isoformat()
    response = requests.post(url, headers=headers, json=body)
    response.raise_for_status()
    return response.json()['id']

# Função para contar quantas máquinas cada técnico está responsável atualmente
def contar_responsaveis(maquinas):
    responsaveis = {email: 0 for email in tecnicos.keys()}
    for maquina in maquinas:
        if not maquina['preparada'] and maquina['responsavel'] in responsaveis:
            responsaveis[maquina['responsavel']] += 1
    return responsaveis

# Função para escolher um técnico que não tenha mais de 2 máquinas
def escolher_tecnico(responsaveis):
    candidatos = [email for email, count in responsaveis.items() if count < 2]
    if candidatos:
        return random.choice(candidatos)
    else:
        print("Todos os técnicos já estão responsáveis por 2 máquinas.")
        return None

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

# Função para verificar se as máquinas foram preparadas
def verificar_maquinas():
    token = obter_token()
    maquinas = carregar_dados(sharepoint_site_original, sharepoint_drive_original, sharepoint_file_path_original)
    
    # Atualizar a planilha com dados adicionais antes de verificar as máquinas
    atualizar_planilha_com_dados_adicionais()
    
    responsaveis = contar_responsaveis(maquinas)
    maquinas_pendentes = [maquina for maquina in maquinas if not maquina['preparada']]
    
    for maquina in maquinas_pendentes:
        tecnico_email = escolher_tecnico(responsaveis)
        if tecnico_email:
            enviar_alerta(maquina, tecnico_email, token)
            task_id = criar_tarefa_planner(maquina, tecnico_email, token)
            atualizar_responsavel(maquina['colaborador'], tecnico_email, task_id)
            responsaveis[tecnico_email] += 1

verificar_maquinas()