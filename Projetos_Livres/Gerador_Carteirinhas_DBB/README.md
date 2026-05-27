# Gerador de Carteirinhas DBB

Projeto criado para gerar carteirinhas de membros da Diaspora Beninense do Brasil (DBB).

A ideia surgiu a partir de uma necessidade real: ajudar um amigo do Benin a gerar carteirinhas para compatriotas que vivem no Brasil, usando uma planilha com dados dos membros, fotos e assinaturas.

## Objetivo

O programa gera um PDF por pessoa, com a frente e o verso da carteirinha na mesma pagina.

Cada carteirinha usa:

- dados de uma planilha Excel;
- foto do membro;
- assinatura do membro;
- assinatura do presidente;
- logo da associacao;
- QR Code unico informado pelo usuario.

## O que o programa faz

- le uma planilha `.xlsx`;
- valida as colunas obrigatorias;
- aceita pequenas variacoes nos nomes das colunas;
- busca fotos pelo ID do membro;
- busca assinaturas pelo ID do membro;
- aplica marca d'agua, borda, logo, validade e status do membro;
- usa o codigo `DBB/BF/CRBD` no verso;
- exporta 1 PDF por pessoa.

## Tecnologias usadas

- Python
- CustomTkinter
- Pillow
- pandas
- openpyxl

## Estrutura

```text
Gerador_Carteirinhas_DBB/
|-- gerador_carteirinhas.py
|-- requirements.txt
|-- README.md
|-- assets/
|   |-- icon.ico
|   |-- logo_dbb.png
|   |-- assinatura_presidente.png
|   `-- fontes/
|       |-- Aptos-Bold.ttf
|       `-- Aptos-Light.ttf
`-- exemplos/
    |-- dados_teste.xlsx
    |-- fotos/
    `-- assinaturas/
```

## Como instalar

No diretorio do projeto, instale as dependencias:

```bash
pip install -r requirements.txt
```

## Como executar

```bash
python gerador_carteirinhas.py
```

## Como testar com os arquivos do projeto

O programa ja abre com alguns caminhos de exemplo preenchidos:

- planilha: `exemplos/dados_teste.xlsx`;
- logo: `assets/logo_dbb.png`;
- pasta de fotos: `exemplos/fotos`;
- pasta de assinaturas: `exemplos/assinaturas`;
- assinatura do presidente: `assets/assinatura_presidente.png`.

Para testar:

1. Abra o programa.
2. Confira se os caminhos de exemplo estao preenchidos.
3. Selecione uma imagem no campo `QR Code Unico`.
4. Escolha uma pasta de saida.
5. Informe a data de emissao.
6. Informe os anos de validade.
7. Clique em `GERAR AGORA`.
8. Confira os PDFs criados na pasta de saida.

Observacao: o QR Code unico e obrigatorio nesta versao. A pasta de exemplos nao possui um QR dedicado, entao para testar a geracao tecnica voce pode selecionar qualquer imagem quadrada. Para uso real, selecione o QR Code oficial da associacao.

## Padrao da planilha

A planilha deve ter uma pessoa por linha.

Colunas recomendadas:

| id | nome_completo | profissao | rnm_rg | data_nascimento |
|----|---------------|-----------|--------|-----------------|

Regras:

- o ID deve ser unico para cada pessoa;
- o ID deve ser igual ao nome do arquivo da foto;
- o ID deve ser igual ao nome do arquivo da assinatura;
- a data de nascimento deve estar no formato `DD/MM/AAAA`;
- recomenda-se formatar a coluna `id` como texto no Excel, para preservar zeros a esquerda.

O programa tambem aceita alguns apelidos nos cabecalhos, como `nome`, `rnm`, `rg` e `nascimento`.

## Padrao dos arquivos

### Fotos

Cada foto deve ter o ID do membro como nome:

```text
1001.jpg
1002.jpg
```

Tambem sao aceitos arquivos `.png` e `.jpeg`.

### Assinaturas

Cada assinatura deve ter o ID do membro como nome:

```text
1001.png
1002.png
```

Arquivos PNG com fundo transparente costumam dar melhor resultado.

### QR Code unico

O QR Code deve ser selecionado no programa como arquivo `.png`, `.jpg` ou `.jpeg`.

Ele e usado em todas as carteirinhas geradas. O programa preserva a proporcao da imagem dentro da area reservada no verso.

## Resultado gerado

- o programa gera 1 PDF por pessoa;
- cada PDF tem 1 pagina;
- a frente da carteirinha fica a esquerda;
- o verso da carteirinha fica a direita;
- o status `MEMBRO ATIVO` aparece em verde;
- a validade aparece em vermelho;
- o nome do PDF usa os dois primeiros nomes da pessoa;
- se houver nomes repetidos, o programa acrescenta um numero ao final.

## Como gerar executavel

O projeto pode ser compilado com PyInstaller.

Exemplo:

```bash
pyinstaller --windowed --onedir --icon assets/icon.ico gerador_carteirinhas.py
```

## O que aprendi com esse projeto

Esse projeto me ajudou a praticar geracao de documentos, manipulacao de imagens, leitura de planilhas e criacao de interface grafica em Python.

Tambem foi um exercicio interessante por ter uma necessidade real: transformar uma lista de membros, fotos e assinaturas em carteirinhas prontas para impressao.
