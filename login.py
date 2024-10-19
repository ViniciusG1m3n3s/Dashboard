import streamlit as st

# Dicionário com usuários e senhas
usuarios = {"usuario1": "senha1", "usuario2": "senha2"}

# Função para autenticar usuário
def autenticar(usuario, senha):
    return usuario in usuarios and usuarios[usuario] == senha

def login():
    st.sidebar.image("https://finchsolucoes.com.br/img/eb28739f-bef7-4366-9a17-6d629cf5e0d9.png", width=100)
    st.sidebar.header("Login")
    usuario = st.sidebar.text_input("Usuário")
    senha = st.sidebar.text_input("Senha", type="password")

    if st.sidebar.button("Entrar"):
        if autenticar(usuario, senha):
            st.session_state.logado = True  # Atualiza o estado para logado
            st.sidebar.success("Login bem-sucedido!")
            return True  # Login bem-sucedido
        else:
            st.sidebar.error("Usuário ou senha incorretos.")
    return False  # Login falhou
