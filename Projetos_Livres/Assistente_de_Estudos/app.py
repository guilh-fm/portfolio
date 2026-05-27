import streamlit as st

from codigo.assistente import gerar_titulo_chat, responder
from codigo.estado import inicializar_estado
from codigo.prompt_estudo import montar_prompt
from codigo.rag_pdf import criar_recuperador_pdf


st.set_page_config(page_title="Assistente de Estudos", layout="wide")


def aplicar_estilos():
    st.markdown(
        """
        <style>
            .block-container {
                max-width: 980px;
                padding-top: 2rem;
                padding-bottom: 4rem;
            }

            [data-testid="stSidebar"] {
                border-right: 1px solid rgba(125, 125, 125, 0.22);
            }

            .titulo-sidebar {
                font-size: 1.25rem;
                font-weight: 700;
                margin: 0.35rem 0 1.2rem 0;
                padding-bottom: 0.75rem;
                border-bottom: 1px solid rgba(125, 125, 125, 0.18);
            }

            .inicio {
                min-height: 54vh;
                display: flex;
                flex-direction: column;
                justify-content: center;
                gap: 1.1rem;
            }

            .inicio h1 {
                font-size: 2.15rem;
                line-height: 1.15;
                margin: 0;
                letter-spacing: 0;
            }

            .inicio p {
                max-width: 760px;
                font-size: 1.02rem;
                line-height: 1.65;
                color: inherit;
                margin: 0;
            }

            .inicio .destaque {
                max-width: 800px;
                border: 1px solid rgba(125, 125, 125, 0.22);
                border-radius: 18px;
                padding: 1rem 1.15rem;
                background: rgba(125, 125, 125, 0.08);
            }

            div[data-testid="stPopover"] button,
            div[data-testid="stButton"] button {
                min-height: 2.6rem;
            }

            div[data-testid="stChatMessage"] div[data-testid="stButton"] button {
                width: 2.25rem;
                min-width: 2.25rem;
                height: 2.25rem;
                min-height: 2.25rem;
                padding: 0;
                border-radius: 999px;
                font-weight: 800;
            }

            div[data-testid="stChatMessage"] div[data-testid="stButton"] button p {
                width: 100%;
                text-align: center;
                line-height: 1;
            }

            div[data-testid="stVerticalBlockBorderWrapper"] {
                border: 1px solid rgba(125, 125, 125, 0.20);
                border-radius: 22px;
                padding: 0.75rem 0.95rem 0.9rem 0.95rem;
                background: rgba(125, 125, 125, 0.08);
            }

            div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stHorizontalBlock"] {
                align-items: end;
            }

            div[data-testid="stChatMessage"] {
                width: 100%;
            }

            @media (max-width: 720px) {
                .block-container {
                    padding-left: 1rem;
                    padding-right: 1rem;
                }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def criar_nova_conversa():
    novo_nome = f"Chat {len(st.session_state.chat_sessions) + 1}"
    st.session_state.chat_sessions[novo_nome] = []
    st.session_state.chat_atual = novo_nome


def mostrar_sidebar():
    st.sidebar.markdown('<div class="titulo-sidebar">Assistente de estudos</div>', unsafe_allow_html=True)

    if st.sidebar.button("Nova conversa", use_container_width=True):
        criar_nova_conversa()
        st.rerun()

    st.sidebar.markdown("**Conversas**")

    for nome in list(st.session_state.chat_sessions.keys()):
        titulo = nome if len(nome) <= 34 else nome[:31] + "..."
        tipo_botao = "primary" if nome == st.session_state.chat_atual else "secondary"

        if st.sidebar.button(titulo, key=f"chat_{nome}", use_container_width=True, type=tipo_botao):
            st.session_state.chat_atual = nome
            st.rerun()


def mostrar_inicio():
    st.markdown(
        """
        <section class="inicio">
            <div>
                <h1>Assistente de Estudos</h1>
                <p>Sou um assistente de estudos com Inteligência Artificial.</p>
            </div>
            <div class="destaque">
                <p>
                    Recomendo que a utilize a função da engrenagem ao lado do campo da mensagem. Pois com ela o seu prompt enviado à LLM será mais completo.<br><br>
                    Com ele você terá ajustes de nível de estudos (fundamental, médio, superior, ...), disciplina, assunto e persona que responderá, influenciando a linguagem utilizada na resposta.<br><br>
                    Além disso, tenho também a função de receber um arquivo para que eu possa consultá-lo e te responder utilizando ele como fonte de conhecimento.<br><br>
                    Com isso, terá respostas mais eficientes e adequadas ao que você espera.<br>
                    Bons Estudos!
                </p>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def carregar_pdf_se_precisar(arquivo):
    if arquivo is None:
        return

    identificador_pdf = f"{arquivo.name}-{arquivo.size}"

    if st.session_state.pdf_carregado == identificador_pdf:
        return

    with open("temp.pdf", "wb") as pdf:
        pdf.write(arquivo.read())

    try:
        with st.spinner("Lendo PDF..."):
            st.session_state.retriever = criar_recuperador_pdf("temp.pdf")

        st.session_state.pdf_carregado = identificador_pdf
        st.success("PDF carregado com sucesso.")
    except Exception as erro:
        st.session_state.retriever = None
        st.session_state.pdf_carregado = ""
        st.error(f"Não consegui carregar o PDF: {erro}")


def mostrar_configuracoes_prompt():
    st.session_state.nivel = st.text_input("Nível de Estudo", value=st.session_state.nivel)
    st.session_state.disciplina = st.text_input("Disciplina", value=st.session_state.disciplina)
    st.session_state.assunto = st.text_input("Assunto", value=st.session_state.assunto)
    st.session_state.persona = st.selectbox(
        "Quem responde?",
        ["Professor", "Monitor", "Especialista"],
        index=["Professor", "Monitor", "Especialista"].index(st.session_state.persona),
    )

    arquivo = st.file_uploader("Carregar PDF para RAG", type="pdf")
    carregar_pdf_se_precisar(arquivo)

    if st.session_state.retriever is not None:
        st.caption("RAG ativo: o PDF carregado será usado como contexto.")
    else:
        st.caption("RAG inativo: carregue um PDF para usar contexto externo.")


def calcular_colunas_usuario(texto):
    tamanho = len(str(texto))

    if tamanho <= 40:
        return [0.64, 0.36]

    if tamanho <= 120:
        return [0.50, 0.50]

    if tamanho <= 280:
        return [0.35, 0.65]

    return [0.24, 0.76]


def valor_resumo(valor, vazio):
    texto = str(valor).strip()

    if texto == "":
        return vazio

    return texto


def montar_resumo_prompt(pergunta):
    persona = valor_resumo(st.session_state.persona, "Professor")
    nivel = valor_resumo(st.session_state.nivel, "não informado")
    disciplina = valor_resumo(st.session_state.disciplina, "não informada")
    assunto = valor_resumo(st.session_state.assunto, "não informado")

    return (
        "**Prompt Enviado:**\n\n"
        f"Você é um {persona} de estudos.\n\n"
        f"Nível: {nivel}\n"
        f"Disciplina: {disciplina}\n"
        f"Assunto: {assunto}\n\n"
        "Pergunta do aluno:\n"
        f"{pergunta}"
    )


def montar_tooltip_prompt(mensagem):
    if "prompt_resumo" in mensagem:
        return mensagem["prompt_resumo"]

    prompt = mensagem.get("prompt_enviado", "")
    prompt_limpo = str(prompt).strip()
    prompt_limpo = prompt_limpo.replace("$$ $$", "delimitadores em bloco de LaTeX")
    prompt_limpo = prompt_limpo.replace("$ $", "delimitadores inline de LaTeX")
    linhas = [linha.strip() for linha in prompt_limpo.splitlines() if linha.strip()]
    linhas_filtradas = []

    for linha in linhas:
        if linha.startswith("- use delimitadores"):
            continue
        if linha.startswith("- quando houver matematica"):
            continue

        linhas_filtradas.append(linha)

    resumo = "\n".join(linhas_filtradas[:8])
    return f"**Prompt Enviado:**\n\n{resumo}"


def mostrar_mensagem_usuario(mensagem, posicao):
    coluna_espaco, coluna_mensagem = st.columns(calcular_colunas_usuario(mensagem["content"]))

    with coluna_mensagem:
        with st.chat_message("user"):
            coluna_texto, coluna_info = st.columns([0.90, 0.10])

            with coluna_texto:
                st.markdown(mensagem["content"])

            with coluna_info:
                if "prompt_enviado" in mensagem:
                    st.button(
                        "i",
                        key=f"info_prompt_{posicao}",
                        help=montar_tooltip_prompt(mensagem),
                    )


def mostrar_mensagem_assistente(mensagem):
    coluna_mensagem, coluna_espaco = st.columns([0.78, 0.22])

    with coluna_mensagem:
        with st.chat_message("assistant"):
            st.markdown(mensagem["content"], unsafe_allow_html=True)


def mostrar_historico(historico):
    for posicao, mensagem in enumerate(historico):
        if mensagem["role"] == "user":
            mostrar_mensagem_usuario(mensagem, posicao)
        else:
            mostrar_mensagem_assistente(mensagem)


def trocar_titulo_primeira_mensagem(modelo, pergunta, historico):
    if len(historico) > 0:
        return historico

    titulo = gerar_titulo_chat(modelo, pergunta)

    if titulo in st.session_state.chat_sessions:
        titulo = f"{titulo} ({len(st.session_state.chat_sessions)})"

    st.session_state.chat_sessions[titulo] = st.session_state.chat_sessions.pop(
        st.session_state.chat_atual
    )
    st.session_state.chat_atual = titulo

    return st.session_state.chat_sessions[titulo]


def processar_pergunta(pergunta, historico):
    prompt_final = montar_prompt(
        nivel=st.session_state.nivel,
        disciplina=st.session_state.disciplina,
        assunto=st.session_state.assunto,
        persona=st.session_state.persona,
        pergunta=pergunta,
    )

    with st.spinner("Pensando na resposta..."):
        resposta = responder(
            modelo=st.session_state.modelo,
            prompt=prompt_final,
            historico=historico,
            recuperador_pdf=st.session_state.retriever,
        )

    historico = trocar_titulo_primeira_mensagem(st.session_state.modelo, pergunta, historico)

    historico.append({
        "role": "user",
        "content": pergunta,
        "prompt_enviado": prompt_final,
        "prompt_resumo": montar_resumo_prompt(pergunta),
    })

    historico.append({
        "role": "assistant",
        "content": resposta,
    })

    st.session_state.contador_pergunta += 1
    st.rerun()


def mostrar_composer(historico):
    with st.container(border=True):
        coluna_pergunta, coluna_modelo, coluna_config = st.columns([6.25, 1.35, 1.75])

        with coluna_pergunta:
            pergunta = st.chat_input("Pergunte alguma coisa")

        with coluna_modelo:
            st.selectbox(
                "Modelo",
                ["groq", "openai"],
                key="modelo",
                label_visibility="collapsed",
            )

        with coluna_config:
            with st.popover("Config", icon=":material/settings:", use_container_width=True):
                mostrar_configuracoes_prompt()

    if pergunta and pergunta.strip() != "":
        processar_pergunta(pergunta.strip(), historico)


def main():
    inicializar_estado()
    aplicar_estilos()
    mostrar_sidebar()

    historico = st.session_state.chat_sessions[st.session_state.chat_atual]

    if len(historico) == 0:
        mostrar_inicio()
    else:
        mostrar_historico(historico)

    mostrar_composer(historico)


if __name__ == "__main__":
    main()
