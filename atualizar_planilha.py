from Graph_API import carregar_dados, salvar_dados
from datetime import datetime
import pandas as pd
import os

sharepoint_site_original = os.getenv('SHAREPOINT_SITE_ORIGINAL')
sharepoint_drive_original = os.getenv('SHAREPOINT_DRIVE_ORIGINAL')
sharepoint_file_path_original = os.getenv('SHAREPOINT_FILE_PATH_ORIGINAL')

planner_plan_id = os.getenv('PLANNER_PLAN_ID')
perfil_maquinas_file_path = os.getenv('PERFIL_MAQUINAS_FILE_PATH')

sharepoint_site_adicional = os.getenv('SHAREPOINT_SITE_ADICIONAL')
sharepoint_drive_adicional = os.getenv('SHAREPOINT_DRIVE_ADICIONAL')
sharepoint_file_path_adicional = os.getenv('SHAREPOINT_FILE_PATH_ADICIONAL')

# Função para atualizar a planilha original com dados da planilha adicional
def atualizar_planilha_com_dados_adicionais():
    # Carregar dados das planilhas
    df_original = carregar_dados(sharepoint_site_original, sharepoint_drive_original, sharepoint_file_path_original)
    df_adicional = carregar_dados(sharepoint_site_adicional, sharepoint_drive_adicional, sharepoint_file_path_adicional)
    
    # Filtrar dados da planilha adicional pela data atual
    data_atual = datetime.now().date()
    df_adicional['Data'] = pd.to_datetime(df_adicional['Data']).dt.date
    df_filtrada = df_adicional[df_adicional['Data'] == data_atual]
    
    # Atualizar a coluna 'Colaborador' da planilha original com a coluna 'Nome do contratado' da planilha adicional filtrada
    df_original['Colaborador'] = df_filtrada['Nome do contratado'].values
    
    # Atualizar a coluna 'Cargo' da planilha original com a coluna 'Cargo' da planilha adicional filtrada
    df_original['Cargo'] = df_filtrada['Cargo'].values
    
    # Atualizar a coluna 'Telefone para Contato' da planilha original com a coluna 'Telefone para Contato' da planilha adicional filtrada
    df_original['Telefone para Contato'] = df_filtrada['Telefone para Contato'].values
    
    # Atualizar a coluna 'Endereço' da planilha original com a coluna 'Endereço' da planilha adicional filtrada
    df_original['Endereço'] = df_filtrada['Endereço'].values
    
    # Obter o perfil da máquina com base no cargo do colaborador
    df_original['Perfil da Máquina'] = df_filtrada['Cargo'].apply(obter_perfil_maquina)
    
    # Salvar a planilha original atualizada
    salvar_dados(df_original, sharepoint_site_original, sharepoint_drive_original, sharepoint_file_path_original)

# Função para verificar o perfil da máquina com base no cargo
def obter_perfil_maquina(cargo):
    df_perfil_maquinas = pd.read_excel(perfil_maquinas_file_path)
    perfil_maquina = df_perfil_maquinas.loc[df_perfil_maquinas['Cargo'] == cargo, 'Perfil da Máquina'].values
    if len(perfil_maquina) > 0:
        return perfil_maquina[0]
    else:
        return "Perfil não encontrado"

atualizar_planilha_com_dados_adicionais()