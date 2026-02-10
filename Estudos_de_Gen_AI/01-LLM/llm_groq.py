from groq import Groq
from groq import RateLimitError
import os

client = Groq(
    api_key=os.getenv("GROQ_API_KEY")
)

def perguntarGroq(pergunta: str) -> str:
    try:
        resposta = client.chat.completions.create(
            model="meta-llama/llama-4-maverick-17b-128e-instruct",
            messages=[
                {"role": "user", "content": pergunta}
            ]
        )

        return resposta.choices[0].message.content

    except RateLimitError:
        return "Resposta nao gerada por falta de creditos"

    except Exception as e:
        return f"Erro inesperado: {e}"

if __name__ == "__main__":
    pergunta = input("Digite sua pergunta: ")
    resposta = perguntarGroq(pergunta)
    print("\nResposta do LLM (LLaMA 4 - Groq:")
    print(resposta)
