## PFICA - Preenche Formulários de Interconexão CLARO/ALGAR

Esta aplicação foi desenvolvida durante meu estágio no setor de Interconexão em uma consultoria de telecomunicações.

### 1. Objetivo

A aplicações foi desenvolvidas para acelerar o processo de preenchimento do formulário de Interconexão da operadora OI.

### 2. Requisitos

- Python 3.9 ou superior  

### 3. Instalação das Dependências

Antes de executar o programa, instale as bibliotecas necessárias utilizando o `pip`:

``` bash
pip install openpyxl customtkinter
```

### 4. Bibliotecas Utilizadas
- **openpyxl:** usada para leitura de arquivos .xlsx

- **customtkinter:** usada para interface gráfica (GUI)

- **threading:** usada para execução de funções em segundo plano, evitando o travamento da GUI durante o processamento

### 5. Como Executar o Programa
No diretório da aplicação, execute:
``` bash
python pfoi.py
```
A interface gráfica será aberta automaticamente.

### 6. Preparação dos Arquivos
**Planilha Excel**: O programa está preparado para utilizar a planilha de pedido de interconexão, que é solicitada pela OI junto com os formulários. Ela segue esse padrão:
|        |        |        |        |        |        |        |        | **PLANO DE NUMERAÇÃO** |        |        | **ÁREA** | **NÚMERO PEDIDO ITX** |        |        |        |        |        |        |        |
|--------|--------|--------|--------|--------|--------|--------|--------|----------------------|--------|--------|--------|----------------------|--------|--------|--------|--------|--------|--------|--------|
| CIDADE | UF | ESTADO | CN | ÁREA LOCAL | SIGLA | CNL | SETOR | PREFIXO | INICIAL | FINAL | TARIFÁRIA | LOCAL | ÁREA LOCAL | ENDEREÇO DO POI / PPI OPERADORA | CEP | LATITUDE | LONGITUDE |
|--------|----|--------|----|------------|-------|-----|-------|---------|---------|-------|-----------|-------|------------|---------------------------------|-----|----------|-----------|

### 7. Passo a Passo para utilização

**1. Selecionar a base de dados em Excel** 
Selecione a planilha de pedido da OI

**2. Selecionar o formulário**
Selecione o formulário base preenchido com as informações e com as marcações entre "{{}}"

**3. Informar o prefixo dos nomes dos arquivos**
Exemplo de padrão de saída:
```
ITX-OPERADORA-OI - SIGLA DA ÁREA LOCAL.xlsx
```

**4. Escolher o diretório de saída**

**6. Clicar em "Gerar Documentos Excel"**
Para criar todos os formulários preenchidos

### 7. Licença
Este projeto é distribuído sob a licença **MIT**.  
Para mais informações, consulte o arquivo `LICENSE`.