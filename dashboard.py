import streamlit as st
import pandas as pd
import plotly.express as px
import os

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
    # Converte 'Tempo de Análise' para string antes de salvar
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
    # Filtra protocolos com tempo de análise acima de 2 minutos
    pontos_de_atencao = df[df['Tempo de Análise'] > pd.Timedelta(minutes=2)].copy()
    # Remove vírgulas dos protocolos
    pontos_de_atencao['Protocolo'] = pontos_de_atencao['Protocolo'].astype(str).str.replace(',', '', regex=False)
    return pontos_de_atencao

# Interface do Streamlit
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
    
    # Adiciona novos dados ao DataFrame total
    df_total = pd.concat([df_total, df_new], ignore_index=True)
    
    # Salvar o DataFrame atualizado no Excel
    save_data(df_total)
    st.sidebar.success(f'Arquivo "{uploaded_file.name}" carregado e processado com sucesso!')

# Converte para timedelta e datetime apenas para operações temporárias
df_total = convert_to_timedelta_for_calculations(df_total)
df_total = convert_to_datetime_for_calculations(df_total)

# Definindo esquema de cores personalizado
custom_colors = ['#ff571c', '#7f2b0e', '#4c1908']

# Função para calcular TMO por dia
def calcular_tmo_por_dia(df):
    # Cria uma nova coluna com apenas o dia
    df['Dia'] = df['Próximo'].dt.date
    
    # Filtra apenas os finalizados
    df_finalizados = df[df['Status'] == 'FINALIZADO'].copy()  # Cria uma cópia para evitar o aviso
    
    # Agrupa os dados por dia, calcula a soma do tempo de análise e conta o número de protocolos
    df_tmo = df_finalizados.groupby('Dia').agg(
        Tempo_Total=('Tempo de Análise', 'sum'),
        Total_Protocolos=('Tempo de Análise', 'count')
    ).reset_index()
    
    # Calcula o TMO em minutos dividindo o Tempo Total pelo número de Protocolos
    df_tmo['TMO'] = (df_tmo['Tempo_Total'] / pd.Timedelta(minutes=1)) / df_tmo['Total_Protocolos']
    
    return df_tmo[['Dia', 'TMO']]

# Verifica qual opção foi escolhida no dropdown
if opcao_selecionada == "Visão Geral":
    st.header("Visão Geral")

    # Cálculos das métricas gerais
    total_finalizados = len(df_total[df_total['Status'] == 'FINALIZADO'])
    total_reclass = len(df_total[df_total['Status'] == 'RECLASSIFICADO'])
    total_andamento = len(df_total[df_total['Status'] == 'ANDAMENTO_PRE'])
    
    # Cálculo do tempo médio apenas para os finalizados
    tempo_medio = df_total[df_total['Status'] == 'FINALIZADO']['Tempo de Análise'].mean()
    
    # Exibir métricas
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
        color_discrete_sequence=custom_colors  # Aplicando cores personalizadas
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
        color_discrete_sequence=custom_colors  # Aplicando cores personalizadas
    )
    st.plotly_chart(fig_tmo)

    # Tabela de pontos de atenção
    st.subheader("Pontos de Atenção: Protocolos com Tempo Acima de 2 Minutos")
    df_pontos_de_atencao = get_points_of_attention(df_total)
    if not df_pontos_de_atencao.empty:
        # Formatar o tempo na tabela em minutos e segundos
        df_pontos_de_atencao['Tempo de Análise'] = df_pontos_de_atencao['Tempo de Análise'].apply(format_timedelta)
        st.table(df_pontos_de_atencao[['Protocolo', 'Usuário', 'Tempo de Análise']])
    else:
        st.write("Nenhum protocolo com tempo acima de 2 minutos.")

elif opcao_selecionada == "Métricas Individuais":
    st.header("Análise por Analista")
    
    # Selecionar o analista
    analista_selecionado = st.selectbox('Selecione o analista', df_total['Usuário'].unique())
    
    # Filtra o DataFrame pelo analista selecionado
    df_analista = df_total[df_total['Usuário'] == analista_selecionado].copy()  # Cria uma cópia para evitar o aviso
    
    # Cálculos para o analista específico
    total_finalizados_analista = len(df_analista[df_analista['Status'] == 'FINALIZADO'])
    total_reclass_analista = len(df_analista[df_analista['Status'] == 'RECLASSIFICADO'])
    total_andamento_analista = len(df_analista[df_analista['Status'] == 'ANDAMENTO_PRE'])
    
    # Verifica se há registros para o analista selecionado
    if total_finalizados_analista > 0:
        tempo_medio_analista = df_analista[df_analista['Status'] == 'FINALIZADO']['Tempo de Análise'].mean()
    else:
        tempo_medio_analista = pd.NaT  # Se não houver finalizados, define como NaT
    
    # Exibir métricas do analista
    st.subheader(f"Métricas de {analista_selecionado}")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Cadastros", total_finalizados_analista)
    col2.metric("Tempo Médio por Cadastro", format_timedelta(tempo_medio_analista))
    col3.metric("Reclassificações", total_reclass_analista)
    
    # Gráfico de pizza para o status do analista
    st.subheader(f"Distribuição de Status de {analista_selecionado}")
    fig_status_analista = px.pie(
        names=['Finalizado', 'Reclassificado', 'Andamento'],
        values=[total_finalizados_analista, total_reclass_analista, total_andamento_analista],
        title=f'Distribuição de Status - {analista_selecionado}',
        color_discrete_sequence=custom_colors  # Aplicando cores personalizadas
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
        color_discrete_sequence=custom_colors  # Aplicando cores personalizadas
    )
    st.plotly_chart(fig_tmo_analista)

# Botão para salvar a planilha atualizada
if st.sidebar.button("Salvar Dados"):
    save_data(df_total)
    st.sidebar.success("Dados salvos com sucesso!")
