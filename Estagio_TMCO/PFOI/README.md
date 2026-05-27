# PFOI - Preenche Formulários de Interconexão OI

Este projeto foi desenvolvido durante meu estágio no setor de Interconexão em uma consultoria de telecomunicações.

O objetivo foi automatizar o preenchimento de formulários da operadora OI a partir de uma planilha de pedido. Antes disso, o preenchimento era feito manualmente e podia levar bastante tempo dependendo da quantidade de áreas locais solicitadas.

## O que o programa faz

- abre uma planilha Excel com os dados do pedido;
- lê as informações de cada área local;
- abre um formulário modelo em Excel;
- substitui marcadores como `{{AREA_LOCAL}}`, `{{PREFIXO}}` e `{{CEP}}`;
- gera um novo arquivo `.xlsx` preenchido para cada linha da planilha.

## Tecnologias usadas

- Python 3.9 ou superior;
- openpyxl;
- customtkinter;
- threading;
- tkinter.

## Estrutura

```text
PFOI/
├── pfoi.py
├── requirements.txt
├── README.md
└── exemplos/
    ├── Formulário OI.xlsx
    └── Pedido ITX OI.xlsx
```

## Como instalar

No diretório do projeto, instale as dependências:

```bash
pip install -r requirements.txt
```

## Como executar

```bash
python pfoi.py
```

Depois disso, a interface gráfica será aberta.

## Como usar

1. Selecione a planilha Excel com os dados do pedido.
2. Selecione o formulário modelo em Excel.
3. Informe o prefixo usado no nome dos arquivos gerados.
4. Escolha a pasta de saída.
5. Clique em `Gerar Documentos Excel`.

O programa cria uma pasta chamada `FORMULARIOS_PREENCHIDOS_XLSX` dentro do diretório escolhido.

## Padrão esperado da planilha

A planilha precisa seguir o mesmo padrão de colunas usado pela base de pedido da OI. O programa considera os dados a partir da terceira linha.

Alguns campos usados no preenchimento:

- cidade;
- estado;
- CN;
- área local;
- sigla;
- prefixo;
- inicial;
- final;
- endereço;
- CEP;
- latitude;
- longitude.

## O que aprendi com esse projeto

Esse projeto me ajudou a praticar automação com Excel, criação de interface gráfica simples e organização de um fluxo que antes era manual. Também foi importante para entender melhor como transformar uma necessidade operacional em uma ferramenta prática.

## Observação

Os arquivos da pasta `exemplos` devem ser usados apenas para teste e demonstração. Antes de publicar qualquer base real, é necessário remover dados sensíveis.
