from langchain_community.document_loaders import PyPDFLoader
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import FAISS
from langchain_text_splitters import RecursiveCharacterTextSplitter


# Funções auxiliares para preparar o PDF

def carregar_pdf(caminho_pdf):
    carregador = PyPDFLoader(caminho_pdf)
    documentos = carregador.load()
    return documentos


def dividir_documentos(documentos):
    divisor = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=200,
    )

    partes = divisor.split_documents(documentos)
    return partes


def criar_embeddings():
    embeddings = HuggingFaceEmbeddings(
        model_name="sentence-transformers/all-MiniLM-L6-v2",
    )

    return embeddings


# Função principal do RAG

def criar_recuperador_pdf(caminho_pdf):
    documentos = carregar_pdf(caminho_pdf)
    partes = dividir_documentos(documentos)
    embeddings = criar_embeddings()

    banco_vetorial = FAISS.from_documents(partes, embeddings)
    recuperador = banco_vetorial.as_retriever(search_kwargs={"k": 4})

    return recuperador
