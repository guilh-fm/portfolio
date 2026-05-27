import streamlit as st


def inicializar_estado():
    if "chat_sessions" not in st.session_state:
        st.session_state.chat_sessions = {"Chat 1": []}

    if "chat_atual" not in st.session_state:
        st.session_state.chat_atual = "Chat 1"

    if "retriever" not in st.session_state:
        st.session_state.retriever = None

    if "modelo" not in st.session_state:
        st.session_state.modelo = "groq"

    if "nivel" not in st.session_state:
        st.session_state.nivel = ""

    if "disciplina" not in st.session_state:
        st.session_state.disciplina = ""

    if "assunto" not in st.session_state:
        st.session_state.assunto = ""

    if "persona" not in st.session_state:
        st.session_state.persona = "Professor"

    if "pdf_carregado" not in st.session_state:
        st.session_state.pdf_carregado = ""

    if "contador_pergunta" not in st.session_state:
        st.session_state.contador_pergunta = 1
