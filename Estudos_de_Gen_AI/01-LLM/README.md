## 1 - Uso de LLM com Python via API

### 1. Objetivo
O objetivo deste código é demonstrar o uso de um LLM (Large Language Model) em Python, através de uma API.
Ao utilizá-lo, um usuário é capaz de enviar um prompt e obter uma resposta de um modelo de LLM.

### 2. Idéia do Projeto
O projeto consiste em uma pequena trilha de estudo inicial e prático de Gen AI através construção de um assistente de estudos.
Está dividido em 5 etapas:
  + 1 - Integração com LLM via API
  + 2 - Engenharia de Prompt
  + 3 - Lanchain
  + 4 - RAG
  + 5 - Agentes
  + 6 - Assistente de Estudos - FINAL

### 3. Fluxo de Funcionamento
```
Usuário executa o script e digita um prompt
                  ↓
O programa envia o prompt para API
                  ↓
O LLM processa o prompt e gera uma resposta
                  ↓
O programa recebe e exibe a resposta do LLM em texto
```
### 4. Como usar?
Para utilizar a API do LLM, é necessário ter uma chave de autenticação e por segurança ela não foi escrita diretamente no código.  
Portanto, é necessário criar uma API Key no [site da OpenAI](https://openai.com/pt-BR/api/) e depois configurar a API Key como Variável de ambiente

  **1. Para definir a variável no Windows (PowerShell)**
  ```
  setx OPENAI_API_KEY "sua_api_key"
  ```

 **2. Para verificar se a variável foi configurada**
 ```
 echo $Env:OPENAI_API_KEY
  ```


