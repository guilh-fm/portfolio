def montar_prompt(nivel, disciplina, assunto, persona, pergunta):
    nivel = nivel.strip() or "não informado"
    disciplina = disciplina.strip() or "não informada"
    assunto = assunto.strip() or "não informado"
    persona = persona.strip() or "Professor"

    prompt = f"""
Você é um {persona} de estudos.

Contexto:
- Nível de estudo: {nivel}
- Disciplina: {disciplina}
- Assunto: {assunto}

Responda de forma didática, clara e organizada.
Adapte a linguagem ao contexto informado.
Use LaTeX apenas quando houver matemática.

Pergunta do aluno:
{pergunta}
""".strip()

    return prompt
