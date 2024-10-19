import streamlit as st
from login import login
from dashboard import dashboard

# Verifica se o usuário está logado
if 'logado' not in st.session_state:
    st.session_state.logado = False

# Se não estiver logado, mostra o formulário de login
if not st.session_state.logado:
    if login():
        st.rerun()  # Reinicia a aplicação para carregar a dashboard
else:
    # Se estiver logado, mostra a dashboard
    dashboard()
