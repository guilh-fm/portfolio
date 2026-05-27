import os
import re

from dotenv import load_dotenv
from langchain_groq import ChatGroq
from langchain_openai import ChatOpenAI


load_dotenv()

MODELO_OPENAI_PADRAO = "gpt-4o-mini"
MODELO_GROQ_PADRAO = "llama-3.3-70b-versatile"


# Funções auxiliares

def buscar_variavel(nome_variavel):
    valor = os.getenv(nome_variavel)

    if valor is None or valor.strip() == "":
        raise ValueError(f"A variável de ambiente {nome_variavel} não foi configurada.")

    return valor


def criar_llm(modelo):
    if modelo == "openai":
        chave_openai = buscar_variavel("OPENAI_API_KEY")

        llm = ChatOpenAI(
            model=os.getenv("OPENAI_MODEL", MODELO_OPENAI_PADRAO),
            api_key=chave_openai,
            temperature=0.3,
        )

        return llm

    if modelo == "groq":
        chave_groq = buscar_variavel("GROQ_API_KEY")

        llm = ChatGroq(
            model=os.getenv("GROQ_MODEL", MODELO_GROQ_PADRAO),
            api_key=chave_groq,
            temperature=0.3,
        )

        return llm

    raise ValueError("Modelo inválido. Use groq ou openai.")


def formatar_historico(historico):
    texto = ""

    for mensagem in historico:
        papel = mensagem.get("role", "")
        conteudo = mensagem.get("content", "")
        texto += f"{papel}: {conteudo}\n"

    return texto


def buscar_contexto_pdf(recuperador, pergunta):
    if recuperador is None:
        return ""

    try:
        documentos = recuperador.invoke(pergunta)
    except AttributeError:
        documentos = recuperador.get_relevant_documents(pergunta)

    textos = []

    for documento in documentos:
        textos.append(documento.page_content)

    contexto = "\n\n".join(textos)
    return contexto


def montar_prompt_final(prompt, historico, contexto_pdf):
    historico_formatado = formatar_historico(historico)

    prompt_final = f"""
HISTORICO DA CONVERSA:
{historico_formatado}

CONTEXTO DO PDF:
{contexto_pdf}

PERGUNTA ATUAL:
{prompt}
"""

    return prompt_final


def limpar_titulo(texto):
    titulo = texto.strip()
    titulo = re.sub(r"[\n\r\t]+", " ", titulo)
    titulo = re.sub(r"[^\w\s-]", "", titulo, flags=re.UNICODE)
    titulo = re.sub(r"\s+", " ", titulo).strip()

    if titulo == "":
        titulo = "Nova conversa"

    palavras = titulo.split()
    titulo = " ".join(palavras[:6])

    return titulo


def titulo_por_fallback(pergunta):
    titulo = limpar_titulo(pergunta)
    return titulo


def traduzir_erro_api(erro):
    mensagem = str(erro)
    mensagem_minuscula = mensagem.lower()

    if "insufficient_quota" in mensagem_minuscula or "exceeded your current quota" in mensagem_minuscula:
        return (
            "A OpenAI recusou a chamada por falta de cota ou billing. "
            "Confira sua conta, limite e variável OPENAI_API_KEY."
        )

    if "model_not_found" in mensagem_minuscula or "does not exist or you do not have access" in mensagem_minuscula:
        return (
            "O modelo configurado não existe ou sua chave não tem acesso a ele. "
            "Na Groq, confira a variável GROQ_MODEL ou use llama-3.3-70b-versatile."
        )

    if "api key" in mensagem_minuscula or "variavel de ambiente" in mensagem_minuscula:
        return mensagem

    if "variável de ambiente" in mensagem_minuscula:
        return mensagem

    return f"Não consegui chamar o modelo agora. Detalhe técnico: {mensagem}"


# Funções principais do assistente

def gerar_titulo_chat(modelo, pergunta):
    try:
        llm = criar_llm(modelo)

        prompt = f"""
Resuma a pergunta abaixo em no máximo 6 palavras.
Não use pontuação.

Pergunta:
{pergunta}
"""

        resposta = llm.invoke(prompt)
        titulo = limpar_titulo(resposta.content)
    except Exception:
        titulo = titulo_por_fallback(pergunta)

    return titulo


def responder(modelo, prompt, historico, recuperador_pdf=None):
    try:
        llm = criar_llm(modelo)
        contexto_pdf = buscar_contexto_pdf(recuperador_pdf, prompt)
        prompt_final = montar_prompt_final(prompt, historico, contexto_pdf)

        resposta = llm.invoke(prompt_final)
        return resposta.content
    except Exception as erro:
        mensagem = traduzir_erro_api(erro)
        return mensagem
