from llm_openai import perguntarOpenAI
from llm_groq import perguntarGroq

def selecionar():
    modelo = input("Deseja utilizar um modelo OpenAI ou Groq? ").strip().lower()
    
    if modelo == "openai":
        return perguntarOpenAI
    elif modelo == "groq":
        return perguntarGroq
    else:
        print("Modelo invalido! Usando o Groq como padrao.")
        return perguntarGroq

if __name__ == "__main__":
    perguntar = selecionar()
    prompt = input("Digite sua pergunta: ")
    resposta = perguntar(prompt)
    print("\nResposta:", resposta)