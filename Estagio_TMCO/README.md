# Projetos de Estágio - TMCO

Esta pasta reúne projetos desenvolvidos durante meu estágio no setor de Interconexão em uma consultoria de telecomunicações.

Os projetos tinham um objetivo prático: reduzir o tempo gasto no preenchimento manual de formulários de interconexão. Em vez de copiar dados linha por linha, os programas leem uma base em Excel e geram os documentos preenchidos automaticamente.

## Projetos

### PFICA

Automatiza formulários de interconexão das operadoras Claro e Algar.

- lê uma planilha Excel com os dados;
- preenche modelos Word;
- gera arquivos DOCX;
- converte os documentos para PDF usando o Microsoft Word.

### PFOI

Automatiza formulários de interconexão da operadora OI.

- lê uma planilha de pedido em Excel;
- preenche um formulário modelo em Excel;
- gera um arquivo XLSX preenchido para cada área local.

## Tecnologias usadas

- Python
- openpyxl
- python-docx
- pywin32
- customtkinter

## Observação

Os exemplos incluídos nesta pasta servem apenas para demonstração. Antes de publicar qualquer arquivo real, é importante remover dados internos, nomes sensíveis ou informações de clientes.
