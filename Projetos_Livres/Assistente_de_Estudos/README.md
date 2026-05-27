# Assistente de Estudos

Este projeto foi criado para estudar IA generativa na pratica usando Python.

A ideia foi construir um assistente de estudos e aplicar, em etapas, os conceitos que eu estava aprendendo: chamada de LLM via API, engenharia de prompt, LangChain e RAG com PDF.

## Objetivo

O assistente ajuda o usuario a estudar um assunto com respostas mais direcionadas.

O usuario pode informar:

- nivel de estudo;
- disciplina;
- assunto;
- persona da resposta;
- pergunta;
- PDF para servir como contexto.

Com essas informacoes, o programa monta um prompt mais organizado e envia para um modelo de linguagem.

## Escadinha de aprendizado

### 1. Chamada de LLM via API

A primeira etapa foi entender como chamar modelos de linguagem usando chaves de API.

Neste projeto, e possivel usar:

- OpenAI;
- Groq.

### 2. Engenharia de prompt

O prompt considera o nivel, a disciplina, o assunto e a persona escolhida.

Isso deixa a resposta mais orientada para o estudo, em vez de enviar apenas a pergunta solta para o modelo.

### 3. LangChain

O LangChain foi usado para padronizar a chamada dos modelos e para conectar o fluxo de RAG.

### 4. RAG com PDF

O fluxo de RAG funciona assim:

1. o usuario envia um PDF;
2. o programa le o arquivo;
3. o texto e dividido em partes menores;
4. os embeddings sao gerados localmente;
5. as partes sao salvas em um indice FAISS;
6. os trechos mais relevantes entram no prompt final.

### 5. Interface web

A interface foi feita com Streamlit.

Ela permite:

- criar conversas;
- alternar entre conversas;
- configurar o contexto de estudo pela engrenagem da area de pergunta;
- escolher o modelo pela area de pergunta;
- carregar PDF;
- visualizar o prompt enviado ao modelo;
- diferenciar mensagens do usuario e do assistente por lado e cor.

## Tecnologias usadas

- Python
- Streamlit
- LangChain
- OpenAI API
- Groq API
- FAISS
- sentence-transformers
- pypdf

## Estrutura

```text
Assistente_de_Estudos/
|-- app.py
|-- requirements.txt
|-- README.md
`-- codigo/
    |-- __init__.py
    |-- assistente.py
    |-- estado.py
    |-- prompt_estudo.py
    `-- rag_pdf.py
```

## Como instalar

No diretorio do projeto, instale as dependencias:

```bash
pip install -r requirements.txt
```

## Variaveis de ambiente

Para usar OpenAI:

```bash
setx OPENAI_API_KEY "sua_chave"
```

Para usar Groq:

```bash
setx GROQ_API_KEY "sua_chave"
```

Depois de configurar com `setx`, feche e abra o terminal novamente.

## Modelos padrao

O projeto usa estes modelos por padrao:

- OpenAI: `gpt-4o-mini`;
- Groq: `llama-3.3-70b-versatile`.

Se quiser trocar sem alterar o codigo, configure:

```bash
setx OPENAI_MODEL "nome_do_modelo"
setx GROQ_MODEL "nome_do_modelo"
```

## Como executar

```bash
python -m streamlit run app.py
```

## Arquivos gerados

Durante o uso, o programa pode gerar:

- `temp.pdf`, com o PDF carregado temporariamente.

Esse arquivo nao deve ser versionado.

## Observacoes sobre erros comuns

Se aparecer erro de cota da OpenAI, o problema esta na conta, billing ou limite da chave configurada.

Se aparecer erro de modelo da Groq, confira se a variavel `GROQ_MODEL` esta apontando para um modelo disponivel na sua conta.

## O que aprendi com esse projeto

Esse projeto me ajudou a entender como uma aplicacao com IA generativa e montada por partes.

Comecei com uma chamada simples de API, depois trabalhei engenharia de prompt, usei LangChain para organizar os modelos e adicionei RAG para buscar contexto em PDF.
