import streamlit as st
import pandas as pd
import plotly.express as px
import os
from datetime import datetime

# Nome do arquivo Excel que será usado para armazenar os dados
excel_file = 'dados_acumulados.xlsx'

# Função para carregar os dados do Excel
def load_data():
    if os.path.exists(excel_file):
        df_total = pd.read_excel(excel_file, engine='openpyxl')
    else:
        df_total = pd.DataFrame(columns=['Protocolo', 'Usuário', 'Status', 'Tempo de Análise', 'Próximo'])
    return df_total

# Função para salvar os dados no Excel
def save_data(df):
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

# Função para garantir que a coluna 'PRÓXIMO' esteja no formato de datetime
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
    
    # Carregar dados acumulados do arquivo Excel
    df_total = load_data()

    st.sidebar.image("https://finchsolucoes.com.br/img/eb28739f-bef7-4366-9a17-6d629cf5e0d9.png", width=100)
    st.sidebar.text('')

    # Sidebar para navegação
    st.sidebar.header("Navegação")
    opcao_selecionada = st.sidebar.selectbox("Escolha uma visão", ["Visão Geral", "Métricas Individuais"])

    # Upload de planilha na sidebar
    uploaded_file = st.sidebar.file_uploader("Carregar nova planilha", type=["xlsx"])

    if uploaded_file is not None:
        df_new = pd.read_excel(uploaded_file)
        df_total = pd.concat([df_total, df_new], ignore_index=True)
        save_data(df_total)
        st.sidebar.success(f'Arquivo "{uploaded_file.name}" carregado e processado com sucesso!')

    # Converte para timedelta e datetime apenas para operações temporárias
    df_total = convert_to_timedelta_for_calculations(df_total)
    df_total = convert_to_datetime_for_calculations(df_total)

    # Adiciona filtros de datas na sidebar
    st.sidebar.subheader("Filtro por Data")
    min_date = df_total['Próximo'].min().date() if not df_total.empty else datetime.today().date()
    max_date = df_total['Próximo'].max().date() if not df_total.empty else datetime.today().date()

    data_inicial = st.sidebar.date_input("Data Inicial", min_date)
    data_final = st.sidebar.date_input("Data Final", max_date)

    if data_inicial > data_final:
        st.sidebar.error("A data inicial não pode ser posterior à data final!")

    df_total = df_total[(df_total['Próximo'].dt.date >= data_inicial) & (df_total['Próximo'].dt.date <= data_final)]

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
        total_finalizados = len(df_total[df_total['Status'] == 'FINALIZADO'])
        total_reclass = len(df_total[df_total['Status'] == 'RECLASSIFICADO'])
        total_andamento = len(df_total[df_total['Status'] == 'ANDAMENTO_PRE'])
        tempo_medio = df_total[df_total['Status'] == 'FINALIZADO']['Tempo de Análise'].mean()

        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Cadastros", total_finalizados)
        col2.metric("Tempo Médio por Cadastro", format_timedelta(tempo_medio))
        col3.metric("Reclassificações", total_reclass)

        # Gráfico de pizza para o status
        st.subheader("Distribuição de Status")
        fig_status = px.pie(
            names=['Finalizado', 'Reclassificado', 'Andamento'],
            values=[total_finalizados, total_reclass, total_andamento],
            title='Distribuição de Status',
            color_discrete_sequence=custom_colors
        )
        st.plotly_chart(fig_status)

        # Gráfico de TMO por dia
        st.subheader("Tempo Médio de Operação (TMO) por Dia")
        df_tmo = calcular_tmo_por_dia(df_total)
        fig_tmo = px.bar(
            df_tmo,
            x='Dia',
            y='TMO',
            title='TMO por Dia (em minutos)',
            labels={'TMO': 'TMO (min)', 'Dia': 'Data'},
            color_discrete_sequence=custom_colors
        )
        st.plotly_chart(fig_tmo)

    elif opcao_selecionada == "Métricas Individuais":
        st.header("Análise por Analista")
        analista_selecionado = st.selectbox('Selecione o analista', df_total['Usuário'].unique())
        df_analista = df_total[df_total['Usuário'] == analista_selecionado].copy()

        total_finalizados_analista = len(df_analista[df_analista['Status'] == 'FINALIZADO'])
        total_reclass_analista = len(df_analista[df_analista['Status'] == 'RECLASSIFICADO'])
        total_andamento_analista = len(df_analista[df_analista['Status'] == 'ANDAMENTO_PRE'])
        tempo_medio_analista = df_analista[df_analista['Status'] == 'FINALIZADO']['Tempo de Análise'].mean()

        # st.write("---") - DIVISOR

        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Cadastros", total_finalizados_analista)
        col2.metric("Tempo Médio por Cadastro", format_timedelta(tempo_medio_analista))
        col3.metric("Reclassificações", total_reclass_analista)

        st.subheader(f"Carteiras Cadastradas por {analista_selecionado}")
        carteiras_analista = pd.DataFrame(df_analista['Carteira'].dropna().unique(), columns=['Carteiras'])
        st.write(carteiras_analista.to_html(index=False, justify='left', border=0), unsafe_allow_html=True)
        st.markdown("<style>table {width: 100%;}</style>", unsafe_allow_html=True)

        # Gráfico de pizza para o status do analista selecionado
        st.subheader(f"Distribuição de Status de {analista_selecionado}")
        fig_status_analista = px.pie(
            names=['Finalizado', 'Reclassificado', 'Andamento'],
            values=[total_finalizados_analista, total_reclass_analista, total_andamento_analista],
            title=f'Distribuição de Status - {analista_selecionado}',
            color_discrete_sequence=custom_colors
        )
        st.plotly_chart(fig_status_analista)

        # Gráfico de TMO por dia para o analista selecionado
        st.subheader(f"Tempo Médio de Operação (TMO) por Dia de {analista_selecionado}")
        df_tmo_analista = calcular_tmo_por_dia(df_analista)
        fig_tmo_analista = px.bar(
            df_tmo_analista,
            x='Dia',
            y='TMO',
            title=f'TMO por Dia - {analista_selecionado} (em minutos)',
            labels={'TMO': 'TMO (min)', 'Dia': 'Data'},
            color_discrete_sequence=custom_colors
        )
        st.plotly_chart(fig_tmo_analista)

        # Adicionar pontos de atenção do analista específico
        st.subheader(f"Pontos de Atenção de {analista_selecionado}")
        pontos_atencao_analista = get_points_of_attention(df_analista)
        
        if not pontos_atencao_analista.empty:
            st.table(pontos_atencao_analista[['Protocolo', 'Tempo de Análise']].assign(
                **{'Tempo de Análise': pontos_atencao_analista['Tempo de Análise'].apply(format_timedelta)}
            ))
        else:
            st.write("Nenhum ponto de atenção identificado para este analista.")

    # Botão para salvar a planilha atualizada
    if st.sidebar.button("Salvar Dados"):
        save_data(df_total)
        st.sidebar.success("Dados salvos com sucesso!")
