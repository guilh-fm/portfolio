# PFICA - Preenche Formulários de Interconexão CLARO/ALGAR

Este projeto foi desenvolvido durante meu estágio no setor de Interconexão em uma consultoria de telecomunicações.

O objetivo foi automatizar o preenchimento de formulários de interconexão das operadoras Claro e Algar. A ideia era reduzir o tempo gasto com preenchimento manual e padronizar a geração dos documentos.

## O que o programa faz

- abre uma planilha Excel com os dados das áreas locais;
- lê cada linha da base de dados;
- abre um modelo Word da operadora selecionada;
- substitui marcadores como `{{SIGLA}}`, `{{AREA_LOCAL}}`, `{{PREFIXO}}` e `{{CEP}}`;
- gera arquivos `.docx` preenchidos;
- converte os documentos gerados para `.pdf`.

## Tecnologias usadas

- Python 3.9 ou superior;
- openpyxl;
- python-docx;
- pywin32;
- customtkinter;
- threading;
- tkinter.

## Requisitos do sistema

Para a conversão de Word para PDF, é necessário:

- Windows;
- Microsoft Word instalado.

Essa etapa usa automação COM pelo pacote `pywin32`.

## Estrutura

```text
PFICA/
├── pfica.py
├── requirements.txt
├── README.md
└── exemplos/
    ├── Base Exemplo.xlsx
    ├── Formulário ALGAR - Base.docx
    └── Formulário CLARO STFC - Base.docx
```

## Como instalar

No diretório do projeto, instale as dependências:

```bash
pip install -r requirements.txt
```

## Como executar

```bash
python pfica.py
```

Depois disso, a interface gráfica será aberta.

## Como usar

1. Selecione a operadora do modelo: `CLARO` ou `ALGAR`.
2. Selecione a planilha Excel com os dados.
3. Selecione o modelo Word correspondente.
4. Informe o prefixo usado no nome dos arquivos gerados.
5. Escolha a pasta de saída.
6. Clique em `Gerar Documentos`.
7. Após a geração dos DOCX, clique em `Converter para PDF`.

O programa cria duas pastas dentro do diretório de saída:

- `DOCX`, com os documentos Word preenchidos;
- `PDF`, com os documentos convertidos.

## Padrão esperado da planilha

A planilha deve seguir este padrão de colunas:

| UF | CN | AREA_LOCAL | SIGLA | MUNICIPIO | ENDERECO | CEP | LATITUDE | LONGITUDE | EOT | PREFIXO | INICIAL | FINAL |
|----|----|------------|-------|-----------|----------|-----|----------|-----------|-----|---------|---------|-------|

## O que aprendi com esse projeto

Esse projeto me ajudou a praticar automação de documentos, leitura de planilhas, manipulação de arquivos Word e criação de interface gráfica. Também foi um exercício importante de transformar uma tarefa repetitiva do estágio em uma ferramenta mais rápida e menos sujeita a erro manual.

## Observação

Os arquivos da pasta `exemplos` devem ser usados apenas para teste e demonstração. Antes de publicar qualquer documento real, é necessário remover dados sensíveis.
