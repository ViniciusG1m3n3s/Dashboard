import streamlit as st
import pandas as pd
import plotly.express as px
import os
from datetime import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from diario import diario  # Importa o diário de bordo

# Autenticação no Google Drive usando o client_secret.json
def authenticate_google_drive():
    gauth = GoogleAuth()
    
    # Use o arquivo baixado do Google Cloud
    gauth.LoadCredentialsFile('client_secret.json')

    if gauth.credentials is None:
        # Realiza a autenticação no navegador
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Atualiza o token expirado
        gauth.Refresh()
    else:
        # Carrega as credenciais
        gauth.Authorize()

    # Salva as credenciais em um arquivo para uso futuro
    gauth.SaveCredentialsFile('client_secret_623211291508-of9paiflonh9b9c4f0i44pg14e4apr51.apps.googleusercontent.com.json')

    # Conectando ao Google Drive
    drive = GoogleDrive(gauth)
    return drive

# Função para carregar o arquivo do Google Drive pelo nome
def load_data_from_drive(drive, usuario):
    file_name = f'dados_acumulados_{usuario}.xlsx'
    # Busca o arquivo pelo nome no Google Drive
    file_list = drive.ListFile({'q': f"title='{file_name}'"}).GetList()
    if file_list:
        # Baixa o arquivo
        file = file_list[0]
        file.GetContentFile(file_name)
        df_total = pd.read_excel(file_name, engine='openpyxl')
    else:
        df_total = pd.DataFrame(columns=['Protocolo', 'Usuário', 'Status', 'Tempo de Análise', 'Próximo'])
    return df_total

# Função para salvar o arquivo no Google Drive
def save_data_to_drive(drive, df, usuario):
    file_name = f'dados_acumulados_{usuario}.xlsx'
    df['Tempo de Análise'] = df['Tempo de Análise'].astype(str)
    df.to_excel(file_name, index=False)

    # Verifica se o arquivo já existe no Google Drive
    file_list = drive.ListFile({'q': f"title='{file_name}'"}).GetList()
    if file_list:
        # Atualiza o arquivo existente
        file = file_list[0]
        file.SetContentFile(file_name)
        file.Upload()
    else:
        # Cria um novo arquivo
        file = drive.CreateFile({'title': file_name})
        file.SetContentFile(file_name)
        file.Upload()

# Função para carregar os dados do Excel do usuário logado
def load_data(usuario):
    excel_file = f'dados_acumulados_{usuario}.xlsx'  # Nome do arquivo específico do usuário
    if os.path.exists(excel_file):
        df_total = pd.read_excel(excel_file, engine='openpyxl')
    else:
        df_total = pd.DataFrame(columns=['Protocolo', 'Usuário', 'Status', 'Tempo de Análise', 'Próximo'])
    return df_total

# Função para salvar os dados no Excel do usuário logado
def save_data(df, usuario):
    excel_file = f'dados_acumulados_{usuario}.xlsx'  # Nome do arquivo específico do usuário
    df['Tempo de Análise'] = df['Tempo de Análise'].astype(str)
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=False)

# Função para garantir que a coluna 'Tempo de Análise' esteja no formato timedelta temporariamente para cálculos
def convert_to_timedelta_for_calculations(df):
    df['Tempo de Análise'] = pd.to_timedelta(df['Tempo de Análise'], errors='coerce')
    return df

# Função para formatar o timedelta em minutos e segundos
def format_timedelta(td):
    if pd.isnull(td):
        return "0 min"
    total_seconds = int(td.total_seconds())
    minutes, seconds = divmod(total_seconds, 60)
    return f"{minutes} min {seconds} sec"

# Função para garantir que a coluna 'Próximo' esteja no formato de datetime
def convert_to_datetime_for_calculations(df):
    df['Próximo'] = pd.to_datetime(df['Próximo'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    return df

# Função para obter pontos de atenção
def get_points_of_attention(df):
    pontos_de_atencao = df[df['Tempo de Análise'] > pd.Timedelta(minutes=2)].copy()
    pontos_de_atencao['Protocolo'] = pontos_de_atencao['Protocolo'].astype(str).str.replace(',', '', regex=False)
    return pontos_de_atencao

# Função principal da dashboard
def dashboard():
    st.title("Dashboard de Produtividade")
    
    # Autenticar no Google Drive
    drive = authenticate_google_drive()
    
    # Carregar dados acumulados do arquivo Excel do usuário logado
    usuario_logado = st.session_state.usuario_logado  # Obtém o usuário logado
    df_total = load_data_from_drive(drive, usuario_logado)  # Carrega dados específicos do usuário

    st.sidebar.image("https://finchsolucoes.com.br/img/eb28739f-bef7-4366-9a17-6d629cf5e0d9.png", width=100)
    st.sidebar.text('')

    # Sidebar para navegação
    st.sidebar.header("Navegação")
    opcao_selecionada = st.sidebar.selectbox("Escolha uma visão", ["Visão Geral", "Métricas Individuais", "Diário de Bordo"])

    # Upload de planilha na sidebar
    uploaded_file = st.sidebar.file_uploader("Carregar nova planilha", type=["xlsx"])

    if uploaded_file is not None:
        df_new = pd.read_excel(uploaded_file)
        df_total = pd.concat([df_total, df_new], ignore_index=True)
        save_data(df_total, usuario_logado)  # Atualiza a planilha específica do usuário
        st.sidebar.success(f'Arquivo "{uploaded_file.name}" carregado e processado com sucesso!')

    # Converte para timedelta e datetime apenas para operações temporárias
    df_total = convert_to_timedelta_for_calculations(df_total)
    df_total = convert_to_datetime_for_calculations(df_total)

    custom_colors = ['#ff571c', '#7f2b0e', '#4c1908']

    # Função para calcular TMO por dia
    def calcular_tmo_por_dia(df):
        df['Dia'] = df['Próximo'].dt.date
        df_finalizados = df[df['Status'] == 'FINALIZADO'].copy()
        df_tmo = df_finalizados.groupby('Dia').agg(
            Tempo_Total=('Tempo de Análise', 'sum'),
            Total_Protocolos=('Tempo de Análise', 'count')
        ).reset_index()
        df_tmo['TMO'] = (df_tmo['Tempo_Total'] / pd.Timedelta(minutes=1)) / df_tmo['Total_Protocolos']
        return df_tmo[['Dia', 'TMO']]

    # Verifica qual opção foi escolhida no dropdown
    if opcao_selecionada == "Visão Geral":
        st.header("Visão Geral")

        # Adiciona filtros de datas 
        min_date = df_total['Próximo'].min().date() if not df_total.empty else datetime.today().date()
        max_date = df_total['Próximo'].max().date() if not df_total.empty else datetime.today().date()

        col1, col2 = st.columns(2)
        with col1:
            data_inicial = st.date_input("Data Inicial", min_date)
        with col2:
            data_final = st.date_input("Data Final", max_date)

        if data_inicial > data_final:
            st.sidebar.error("A data inicial não pode ser posterior à data final!")

        df_total = df_total[(df_total['Próximo'].dt.date >= data_inicial) & (df_total['Próximo'].dt.date <= data_final)]

        total_finalizados = len(df_total[df_total['Status'] == 'FINALIZADO'])
        total_reclass = len(df_total[df_total['Status'] == 'RECLASSIFICADO'])
        total_andamento = len(df_total[df_total['Status'] == 'ANDAMENTO_PRE'])
        tempo_medio = df_total[df_total['Status'] == 'FINALIZADO']['Tempo de Análise'].mean()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total de Cadastros", total_finalizados)
        col2.metric("Reclassificações", total_reclass)
        col3.metric("Andamentos", total_andamento)
        col4.metric("Tempo Médio por Cadastro", format_timedelta(tempo_medio))

        # Gráfico de pizza para o status
        st.subheader("Distribuição de Status")
        fig_status = px.pie(
            names=['Finalizado', 'Reclassificado', 'Andamento'],
            values=[total_finalizados, total_reclass, total_andamento],
            title='Distribuição de Status',
            color_discrete_sequence=custom_colors
        )
        st.plotly_chart(fig_status)