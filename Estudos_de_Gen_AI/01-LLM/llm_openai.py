from openai import OpenAI
from openai import RateLimitError

client = OpenAI()

def perguntarOpenAI(pergunta: str) -> str:
    try:
        resposta = client.responses.create(
            model = "gpt-5-nano",
            input = pergunta,
            store = True,
        )

        return resposta.output_text
    
    except RateLimitError:
        return "Resposta nao gerada por falta de creditos"

if __name__ == "__main__":
    pergunta = input("Digite sua pergunta: ")
    resposta = perguntarOpenAI(pergunta)
    print("\nResposta do LLM (GPT-5-Nano):", resposta)