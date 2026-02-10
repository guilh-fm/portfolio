## PFICA - Preenche Formulários de Interconexão CLARO/ALGAR

Esta aplicação foi desenvolvida durante meu estágio no setor de Interconexão em uma consultoria de telecomunicações.

### 1. Objetivo

A aplicações foi desenvolvidas para acelerar o processo de preenchimento de formulários de interconexão das operadoras Claro e Algar.

### 2. Requisitos

- Python 3.9 ou superior  
- Sistema operacional Windows (necessário para conversão Word → PDF)  
- Microsoft Word instalado  

### 3. Instalação das Dependências

Antes de executar o programa, instale as bibliotecas necessárias utilizando o `pip`:

``` bash
pip install openpyxl python-docx pywin32 customtkinter
```

### 4. Bibliotecas Utilizadas
- **openpyxl:** usada para leitura de arquivos .xlsx

- **python-docx:** usada para leitura e edição de documentos Word

- **pywin32:** usada para conversão em PDF do Microsoft Word

- **customtkinter:** usada para interface gráfica (GUI)

### 5. Como Executar o Programa
No diretório da aplicação, execute:
``` bash
python pfica.py
```
A interface gráfica será aberta automaticamente.

### 6. Preparação dos Arquivos
**Planilha Excel**: A planilha deve seguir esse padrão de colunas para o funcionamento do programa.
| UF | CN | AREA_LOCAL | SIGLA | MUNICIPIO | ENDERECO | CEP | LATITUDE | LONGITUDE | EOT | PREFIXO | INICIAL | FINAL |
|----|----|------------|-------|-----------|----------|-----|----------|-----------|-----|---------|---------|-------|


**Formulário Word**: O formulário precisa ser preenchido com as informações:
- Da operadora
- Dos aspectos da Interconexão
- Do Responsável Operacional
- Do Representante Legal
- Campos a serem preenchidos com {{}}

### 7. Passo a Passo para utilização
**1. Selecionar a operadora**
Essa escolha influencia apenas no nome de saída dos arquivos gerados.

**2. Selecionar a base de dados em Excel** 

**3. Selecionar o formulário Word**
Correspondente à mesma operadora escolhida anteriormente, garantindo coerência com o nome final do arquivo gerado.

**4. Informar o prefixo dos nomes dos arquivos**
Exemplo de padrão de saída:
```
ITX-OPERADORA-CLARO - SIGLA DA ÁREA LOCAL.pdf
```

**5. Escolher o diretório de saída**
O programa criará automaticamente as pastas para arquivos DOCX e PDF.

**6. Clicar em "Gerar Documentos"**
Para criar todos os arquivos Word preenchidos.

**7. clicar em "Converter para PDF"**
Após a finalização da geração dos DOCX, esse botão será liberado.
Clique para converter automaticamente todos os documentos para PDF.

### 8. Licença
Este projeto é distribuído sob a licença **MIT**.  
Para mais informações, consulte o arquivo `LICENSE`.