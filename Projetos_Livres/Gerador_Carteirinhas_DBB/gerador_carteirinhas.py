import calendar
import os
import re
import sys
import threading
import unicodedata
from pathlib import Path
from tkinter import LEFT, SOLID, Label, Toplevel, filedialog, messagebox

import customtkinter as ctk
import pandas as pd
from PIL import Image, ImageDraw, ImageFont


# Configurações principais

DPI = 300
LARGURA_CARTAO = 1016
ALTURA_CARTAO = 638
LARGURA_A4 = 2480
ALTURA_A4 = 3508
ESCALA_PDF = 2
QUALIDADE_PDF = 100
ESCALA_IMPRESSAO_CARTAO = 1.5
OPACIDADE_MARCA_DAGUA = 0.30
ALTURA_MAXIMA_LOGO_TOPO = int(122 * 1.30)
LARGURA_BORDA_CARTAO = 2
TAMANHO_QR_UNICO = 188

TITULO_ASSOCIACAO = "Diáspora Beninense do Brasil (DBB)"
TITULO_CARTEIRA = "CARTEIRA DE MEMBRO"
CODIGO_ASSOCIACAO = "DBB/BF/CRBD"

COLUNAS_OBRIGATORIAS = ("id", "nome_completo", "profissao", "rnm", "data_nascimento")
ORDEM_COLUNAS = "id, nome_completo, profissao, rnm_rg, data_nascimento"

APELIDOS_COLUNAS = {
    "id": {"id", "codigo", "cod", "matricula", "numero", "numero_membro"},
    "nome_completo": {"nome_completo", "nome", "nome_do_membro", "membro"},
    "profissao": {"profissao", "ocupacao", "cargo"},
    "rnm": {"rnm", "rnm_rg", "rnm/rg", "rg", "rnm_rr", "rnm/rr", "rr", "rne", "documento", "documento_migratorio"},
    "data_nascimento": {"data_nascimento", "nascimento", "data_de_nascimento", "dt_nascimento"},
}


# Funções auxiliares

def caminho_recurso(caminho_relativo):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, caminho_relativo)

    pasta_base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(pasta_base, caminho_relativo)


def caminho_padrao(*partes):
    caminho = Path(__file__).resolve().parent.joinpath(*partes)
    return str(caminho)


def detectar_fontes():
    fontes_possiveis = [
        ("assets/fontes/Aptos-Bold.ttf", "assets/fontes/Aptos-Light.ttf"),
        ("Aptos-Bold.ttf", "Aptos-Light.ttf"),
        ("arialbd.ttf", "arial.ttf"),
    ]

    for fonte_negrito, fonte_leve in fontes_possiveis:
        caminho_negrito = caminho_recurso(fonte_negrito)
        caminho_leve = caminho_recurso(fonte_leve)

        try:
            ImageFont.truetype(caminho_negrito, 10)
            ImageFont.truetype(caminho_leve, 10)
            return caminho_negrito, caminho_leve
        except Exception:
            pass

    return None, None


def carregar_fonte(caminho_fonte, tamanho):
    if caminho_fonte is None:
        return ImageFont.load_default()

    try:
        return ImageFont.truetype(caminho_fonte, max(1, int(tamanho)))
    except Exception:
        return ImageFont.load_default()


def normalizar_nome_coluna(nome_coluna):
    texto = str(nome_coluna).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(caractere for caractere in texto if not unicodedata.combining(caractere))
    texto = re.sub(r"[^a-z0-9]+", "_", texto).strip("_")
    return texto


def buscar_arquivo_por_id(pasta, identificador, extensoes):
    if pasta == "" or identificador == "":
        return None

    for extensao in extensoes:
        caminho = Path(pasta) / f"{identificador}{extensao}"

        if caminho.exists():
            return str(caminho)

    return None


def calcular_validade(data_emissao, anos_validade):
    data_emissao = str(data_emissao).strip()
    anos_validade = str(anos_validade).strip()

    if data_emissao == "" or not anos_validade.isdigit():
        return "DD/MM/20AA"

    try:
        dia, mes, ano = data_emissao.split("/")
        novo_ano = int(ano) + int(anos_validade)

        if dia == "29" and mes == "02" and not calendar.isleap(novo_ano):
            dia = "28"

        return f"{dia}/{mes}/{novo_ano}"
    except Exception:
        return "DD/MM/20AA"


def redimensionar_corte_central(imagem, tamanho):
    largura_final, altura_final = tamanho
    proporcao_imagem = imagem.width / imagem.height
    proporcao_destino = largura_final / altura_final

    if proporcao_imagem > proporcao_destino:
        nova_altura = altura_final
        nova_largura = int(nova_altura * proporcao_imagem)
        imagem = imagem.resize((nova_largura, nova_altura), Image.Resampling.LANCZOS)
        x_inicial = (nova_largura - largura_final) // 2
        return imagem.crop((x_inicial, 0, x_inicial + largura_final, nova_altura))

    nova_largura = largura_final
    nova_altura = int(nova_largura / proporcao_imagem)
    imagem = imagem.resize((nova_largura, nova_altura), Image.Resampling.LANCZOS)
    y_inicial = (nova_altura - altura_final) // 2
    return imagem.crop((0, y_inicial, nova_largura, y_inicial + altura_final))


def montar_nome_pdf(nome_pessoa, nomes_usados):
    partes_nome = [parte for parte in str(nome_pessoa).split() if parte.strip()]
    nome_base = " ".join(partes_nome[:2]) or "membro"
    nome_base = re.sub(r'[<>:"/\\|?*]+', "", nome_base)
    nome_base = re.sub(r"\s+", " ", nome_base).strip().rstrip(".") or "membro"

    chave = nome_base.lower()
    nomes_usados[chave] = nomes_usados.get(chave, 0) + 1

    if nomes_usados[chave] > 1:
        nome_base = f"{nome_base} {nomes_usados[chave]}"

    return f"{nome_base}.pdf"


def preparar_planilha(caminho_planilha):
    planilha = pd.read_excel(caminho_planilha, dtype=str).dropna(how="all")

    if planilha.empty:
        raise ValueError("A planilha não possui membros para processar.")

    colunas_finais = []
    contagem_colunas = {}

    for coluna in planilha.columns:
        coluna_normalizada = normalizar_nome_coluna(coluna)
        coluna_final = coluna_normalizada

        for coluna_padrao, apelidos in APELIDOS_COLUNAS.items():
            if coluna_normalizada in apelidos:
                coluna_final = coluna_padrao
                break

        if coluna_final in contagem_colunas:
            contagem_colunas[coluna_final] += 1
            coluna_final = f"{coluna_final}_{contagem_colunas[coluna_final]}"
        else:
            contagem_colunas[coluna_final] = 0

        colunas_finais.append(coluna_final)

    planilha.columns = colunas_finais

    colunas_faltando = []

    for coluna in COLUNAS_OBRIGATORIAS:
        if coluna not in planilha.columns:
            colunas_faltando.append(coluna)

    if colunas_faltando:
        raise ValueError(
            "Planilha inválida.\n\n"
            f"Colunas obrigatórias na ordem recomendada: {ORDEM_COLUNAS}\n"
            f"Colunas faltando: {', '.join(colunas_faltando)}"
        )

    erros_vazios = []

    for coluna in COLUNAS_OBRIGATORIAS:
        linhas_vazias = planilha[coluna].fillna("").astype(str).str.strip().eq("")

        if linhas_vazias.any():
            linhas = [str(indice + 2) for indice in planilha.index[linhas_vazias].tolist()]
            erros_vazios.append(f"{coluna}: linhas {', '.join(linhas[:10])}")

    if erros_vazios:
        raise ValueError(
            "Planilha inválida.\n\n"
            "As colunas obrigatórias não podem ter valores vazios:\n"
            + "\n".join(erros_vazios)
        )

    return planilha


FONTE_NEGRITO, FONTE_LEVE = detectar_fontes()


class DesenhoEscalado:
    def __init__(self, desenho, escala):
        self.desenho = desenho
        self.escala = escala

    def converter_valor(self, valor):
        return int(round(valor * self.escala))

    def converter_ponto(self, ponto):
        return self.converter_valor(ponto[0]), self.converter_valor(ponto[1])

    def converter_pontos(self, pontos):
        if isinstance(pontos, (list, tuple)) and pontos and isinstance(pontos[0], (list, tuple)):
            return [self.converter_ponto(ponto) for ponto in pontos]

        return self.converter_ponto(pontos)

    def rectangle(self, pontos, **kwargs):
        return self.desenho.rectangle(self.converter_pontos(pontos), **kwargs)

    def line(self, pontos, **kwargs):
        if "width" in kwargs:
            kwargs["width"] = max(1, self.converter_valor(kwargs["width"]))

        return self.desenho.line(self.converter_pontos(pontos), **kwargs)

    def text(self, ponto, texto, **kwargs):
        return self.desenho.text(self.converter_ponto(ponto), texto, **kwargs)

    def textbbox(self, ponto, texto, **kwargs):
        caixa = self.desenho.textbbox(self.converter_ponto(ponto), texto, **kwargs)
        return tuple(valor / self.escala for valor in caixa)


class Dica:
    def __init__(self, widget, texto):
        self.widget = widget
        self.texto = texto
        self.janela_dica = None

        self.widget.bind("<Enter>", self.mostrar)
        self.widget.bind("<Leave>", self.esconder)

    def mostrar(self, _evento=None):
        if self.janela_dica is not None or self.texto == "":
            return

        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25

        self.janela_dica = Toplevel(self.widget)
        self.janela_dica.wm_overrideredirect(True)
        self.janela_dica.wm_geometry(f"+{x}+{y}")

        label = Label(
            self.janela_dica,
            text=self.texto,
            justify=LEFT,
            background="#ffffe0",
            relief=SOLID,
            borderwidth=1,
            font=("tahoma", "9"),
        )
        label.pack(ipadx=5, ipady=5)

    def esconder(self, _evento=None):
        if self.janela_dica is not None:
            self.janela_dica.destroy()
            self.janela_dica = None


class GeradorCarteirinhas:
    def __init__(self, configuracao, callback_status=None):
        self.configuracao = configuracao
        self.callback_status = callback_status
        self.escala_render = 1

    def registrar_status(self, texto):
        if self.callback_status is not None:
            self.callback_status(texto)

        print(texto)

    def iniciar_renderizacao(self, modo_previa=False):
        if modo_previa:
            self.escala_render = 1
            return

        try:
            self.escala_render = max(1, int(self.configuracao.get("escala_pdf", ESCALA_PDF)))
        except Exception:
            self.escala_render = ESCALA_PDF

    def escala(self, valor):
        return int(round(valor * self.escala_render))

    def tamanho_escalado(self, tamanho):
        return self.escala(tamanho[0]), self.escala(tamanho[1])

    def posicao_escalada(self, posicao):
        return self.escala(posicao[0]), self.escala(posicao[1])

    def pegar_texto(self, dados_pessoa, coluna):
        valor = dados_pessoa.get(coluna)

        try:
            if valor is None or pd.isna(valor):
                return ""
        except Exception:
            pass

        return str(valor).strip()

    def fonte_negrito(self, tamanho):
        return carregar_fonte(FONTE_NEGRITO, self.escala(tamanho))

    def fonte_leve(self, tamanho):
        return carregar_fonte(FONTE_LEVE, self.escala(tamanho))

    def desenhar_texto_centralizado(self, desenho, y, segmentos):
        largura_total = 0

        for texto, fonte, _cor in segmentos:
            caixa = desenho.textbbox((0, 0), texto, font=fonte)
            largura_total += caixa[2] - caixa[0]

        x_atual = (LARGURA_CARTAO - largura_total) // 2

        for texto, fonte, cor in segmentos:
            desenho.text((x_atual, y), texto, font=fonte, fill=cor)
            caixa = desenho.textbbox((0, 0), texto, font=fonte)
            x_atual += caixa[2] - caixa[0]

    def desenhar_borda_cartao(self, desenho):
        desenho.rectangle(
            [(0, 0), (LARGURA_CARTAO - 1, ALTURA_CARTAO - 1)],
            outline="#000000",
            width=LARGURA_BORDA_CARTAO,
        )

    def aplicar_logo_marca_dagua(self, imagem_base, caminho_logo, area):
        if caminho_logo == "" or not os.path.exists(caminho_logo):
            return

        try:
            x_area, y_area, largura_area, altura_area = area
            logo = Image.open(caminho_logo).convert("RGBA")
            escala_logo = min(largura_area / logo.width, altura_area / logo.height)
            largura_logo = max(1, int(logo.width * escala_logo))
            altura_logo = max(1, int(logo.height * escala_logo))
            marca_dagua = logo.resize(self.tamanho_escalado((largura_logo, altura_logo)), Image.Resampling.LANCZOS)

            canal_alpha = marca_dagua.getchannel("A").point(lambda pixel: int(pixel * OPACIDADE_MARCA_DAGUA))
            marca_dagua.putalpha(canal_alpha)

            x_logo = x_area + (largura_area - largura_logo) // 2
            y_logo = y_area + (altura_area - altura_logo) // 2
            imagem_base.paste(marca_dagua, self.posicao_escalada((x_logo, y_logo)), marca_dagua)
        except Exception:
            pass

    def colar_imagem_transparente(self, imagem_base, imagem, posicao):
        imagem_base.paste(imagem, self.posicao_escalada(posicao), imagem)

    def carregar_logo(self):
        caminho_logo = self.configuracao.get("logo_path", "")

        if caminho_logo == "" or not os.path.exists(caminho_logo):
            return None

        try:
            return Image.open(caminho_logo).convert("RGBA")
        except Exception:
            return None

    def carregar_foto_membro(self, identificador):
        caminho_foto = buscar_arquivo_por_id(
            self.configuracao.get("fotos_dir", ""),
            identificador,
            [".jpg", ".jpeg", ".png"],
        )

        if caminho_foto is None:
            return None

        try:
            return Image.open(caminho_foto).convert("RGB")
        except Exception:
            return None

    def carregar_assinatura_membro(self, identificador):
        caminho_assinatura = buscar_arquivo_por_id(
            self.configuracao.get("assinaturas_dir", ""),
            identificador,
            [".png", ".jpg", ".jpeg"],
        )

        if caminho_assinatura is None:
            return None

        try:
            return Image.open(caminho_assinatura).convert("RGBA")
        except Exception:
            return None

    def carregar_assinatura_presidente(self):
        caminho_assinatura = self.configuracao.get("assinatura_presidente_path", "")

        if caminho_assinatura == "" or not os.path.exists(caminho_assinatura):
            return None

        try:
            return Image.open(caminho_assinatura).convert("RGBA")
        except Exception:
            return None

    def carregar_qr_unico(self):
        caminho_qr = self.configuracao.get("qr_path", "")

        if caminho_qr == "" or not os.path.exists(caminho_qr):
            return None

        try:
            qr_origem = Image.open(caminho_qr).convert("RGB")
            qr_origem.thumbnail(self.tamanho_escalado((TAMANHO_QR_UNICO, TAMANHO_QR_UNICO)), Image.Resampling.LANCZOS)
            qr_final = Image.new("RGB", self.tamanho_escalado((TAMANHO_QR_UNICO, TAMANHO_QR_UNICO)), "white")
            x_qr = (qr_final.width - qr_origem.width) // 2
            y_qr = (qr_final.height - qr_origem.height) // 2
            qr_final.paste(qr_origem, (x_qr, y_qr))
            return qr_final
        except Exception:
            return None

    def gerar_frente(self, dados_pessoa, modo_previa=False):
        self.iniciar_renderizacao(modo_previa)

        imagem = Image.new("RGB", self.tamanho_escalado((LARGURA_CARTAO, ALTURA_CARTAO)), "white")
        desenho = DesenhoEscalado(ImageDraw.Draw(imagem), self.escala_render)

        altura_cabecalho = 155
        altura_barra = 5
        altura_rodape = 50
        margem = 30

        logo = self.carregar_logo()
        caminho_logo = self.configuracao.get("logo_path", "")

        if logo is not None:
            altura_logo = ALTURA_MAXIMA_LOGO_TOPO
            largura_logo = int(logo.width * altura_logo / logo.height)
            logo = logo.resize(self.tamanho_escalado((largura_logo, altura_logo)), Image.Resampling.LANCZOS)
            self.colar_imagem_transparente(imagem, logo, (18, 2))

        fonte_titulo = self.fonte_negrito(30)
        fonte_leve_cabecalho = self.fonte_leve(16)
        fonte_negrito_cabecalho = self.fonte_negrito(16)

        caixa_titulo = desenho.textbbox((0, 0), TITULO_ASSOCIACAO, font=fonte_titulo)
        largura_titulo = caixa_titulo[2] - caixa_titulo[0]
        desenho.text(((LARGURA_CARTAO - largura_titulo) // 2, 17), TITULO_ASSOCIACAO, font=fonte_titulo, fill="#111111")

        self.desenhar_texto_centralizado(
            desenho,
            57,
            [
                ("Registro: ", fonte_negrito_cabecalho, "#444444"),
                ("RPJ n° 799.964 A-1021.   ", fonte_leve_cabecalho, "#666666"),
                ("Timbro digital: ", fonte_negrito_cabecalho, "#444444"),
                ("TJPB ALP 11808-MX09", fonte_leve_cabecalho, "#666666"),
            ],
        )

        self.desenhar_texto_centralizado(
            desenho,
            80,
            [
                ("Sede: ", fonte_negrito_cabecalho, "#444444"),
                ("João Pessoa, Brésil     ", fonte_leve_cabecalho, "#666666"),
                ("CNPJ: ", fonte_negrito_cabecalho, "#444444"),
                ("60.931.085/0001-53", fonte_leve_cabecalho, "#666666"),
            ],
        )

        fonte_carteira = self.fonte_negrito(20)
        caixa_carteira = desenho.textbbox((0, 0), TITULO_CARTEIRA, font=fonte_carteira)
        largura_carteira = caixa_carteira[2] - caixa_carteira[0]
        desenho.text(((LARGURA_CARTAO - largura_carteira) // 2, 117), TITULO_CARTEIRA, font=fonte_carteira, fill="#222222")

        y_barra = altura_cabecalho
        largura_barra = 600
        x_inicio_barra = (LARGURA_CARTAO - largura_barra) // 2
        largura_parte = largura_barra // 3

        desenho.rectangle([(x_inicio_barra, y_barra), (x_inicio_barra + largura_parte, y_barra + altura_barra)], fill=self.configuracao["color1"])
        desenho.rectangle([(x_inicio_barra + largura_parte, y_barra), (x_inicio_barra + 2 * largura_parte, y_barra + altura_barra)], fill=self.configuracao["color2"])
        desenho.rectangle([(x_inicio_barra + 2 * largura_parte, y_barra), (x_inicio_barra + largura_barra, y_barra + altura_barra)], fill=self.configuracao["color3"])

        y_conteudo = y_barra + altura_barra + 18
        y_rodape = ALTURA_CARTAO - altura_rodape

        largura_foto = 245
        altura_foto = y_rodape - y_conteudo - 12
        x_foto = margem
        y_foto = y_conteudo

        desenho.rectangle([(x_foto - 2, y_foto - 2), (x_foto + largura_foto + 2, y_foto + altura_foto + 2)], fill="#999999")
        desenho.rectangle([(x_foto, y_foto), (x_foto + largura_foto, y_foto + altura_foto)], fill="#DDDDDD")

        identificador = self.pegar_texto(dados_pessoa, "id")
        foto = self.carregar_foto_membro(identificador)

        if foto is not None:
            foto = redimensionar_corte_central(foto, self.tamanho_escalado((largura_foto, altura_foto)))
            imagem.paste(foto, self.posicao_escalada((x_foto, y_foto)))
        else:
            fonte_sem_foto = self.fonte_leve(22)
            desenho.text((x_foto + 55, y_foto + altura_foto // 2 - 12), "Sem Foto", font=fonte_sem_foto, fill="#888888")

        x_dados = x_foto + largura_foto + 32
        y_dados = y_conteudo + 2
        espacamento_linha = 58

        self.aplicar_logo_marca_dagua(imagem, caminho_logo, (x_dados + 60, y_conteudo - 30, 525, 525))

        fonte_label = self.fonte_negrito(26)
        fonte_valor = self.fonte_leve(26)

        campos = [
            ("Nome Completo:", "nome_completo"),
            ("Profissão:", "profissao"),
            ("RNM/RG:", "rnm"),
            ("Data de Nascimento:", "data_nascimento"),
            ("ID:", "id"),
        ]

        for label, coluna in campos:
            valor = self.pegar_texto(dados_pessoa, coluna)
            desenho.text((x_dados, y_dados), label, font=fonte_label, fill="#111111")

            caixa_label = desenho.textbbox((0, 0), label, font=fonte_label)
            largura_label = caixa_label[2] - caixa_label[0]
            desenho.text((x_dados + largura_label + 10, y_dados), valor, font=fonte_valor, fill="#333333")

            y_dados += espacamento_linha

        desenho.line([(margem, y_rodape), (LARGURA_CARTAO - margem, y_rodape)], fill="#CCCCCC", width=1)

        fonte_rodape = self.fonte_negrito(24)
        desenho.text((margem, y_rodape + 13), "MEMBRO ATIVO", font=fonte_rodape, fill="#008800")

        validade = calcular_validade(
            self.configuracao.get("data_emissao", ""),
            self.configuracao.get("anos_validade", "1"),
        )

        texto_validade = "VALIDADE: "
        caixa_label_validade = desenho.textbbox((0, 0), texto_validade, font=fonte_rodape)
        caixa_data_validade = desenho.textbbox((0, 0), validade, font=fonte_rodape)
        largura_validade = (caixa_label_validade[2] - caixa_label_validade[0]) + (caixa_data_validade[2] - caixa_data_validade[0])
        x_validade = LARGURA_CARTAO - margem - largura_validade

        desenho.text((x_validade, y_rodape + 13), texto_validade, font=fonte_rodape, fill="#111111")
        desenho.text((x_validade + (caixa_label_validade[2] - caixa_label_validade[0]), y_rodape + 13), validade, font=fonte_rodape, fill="#CC0000")

        self.desenhar_borda_cartao(desenho)
        return imagem

    def gerar_verso(self, dados_pessoa, modo_previa=False):
        self.iniciar_renderizacao(modo_previa)

        imagem = Image.new("RGB", self.tamanho_escalado((LARGURA_CARTAO, ALTURA_CARTAO)), "white")
        desenho = DesenhoEscalado(ImageDraw.Draw(imagem), self.escala_render)

        identificador = self.pegar_texto(dados_pessoa, "id")
        margem = 60

        fonte_label = self.fonte_negrito(28)
        fonte_valor = self.fonte_leve(28)
        data_emissao = self.configuracao.get("data_emissao", "") or "DD/MM/AAAA"

        desenho.text((margem, 30), "EMISSÃO:", font=fonte_label, fill="#111111")
        caixa_emissao = desenho.textbbox((0, 0), "EMISSÃO:", font=fonte_label)
        desenho.text((margem + (caixa_emissao[2] - caixa_emissao[0]) + 18, 30), data_emissao, font=fonte_valor, fill="#111111")

        caixa_codigo = desenho.textbbox((0, 0), CODIGO_ASSOCIACAO, font=fonte_label)
        largura_codigo = caixa_codigo[2] - caixa_codigo[0]
        desenho.text((LARGURA_CARTAO - margem - largura_codigo, 30), CODIGO_ASSOCIACAO, font=fonte_label, fill="#111111")

        y_assinatura = 110
        altura_assinatura = 160
        largura_assinatura = 305
        x_assinatura_membro = margem
        x_assinatura_presidente = LARGURA_CARTAO - margem - largura_assinatura

        assinatura_membro = self.carregar_assinatura_membro(identificador)

        if assinatura_membro is not None:
            assinatura_membro = assinatura_membro.resize(self.tamanho_escalado((largura_assinatura, altura_assinatura)), Image.Resampling.LANCZOS)
            self.colar_imagem_transparente(imagem, assinatura_membro, (x_assinatura_membro, y_assinatura))

        assinatura_presidente = self.carregar_assinatura_presidente()

        if assinatura_presidente is not None:
            assinatura_presidente = assinatura_presidente.resize(self.tamanho_escalado((largura_assinatura, altura_assinatura)), Image.Resampling.LANCZOS)
            self.colar_imagem_transparente(imagem, assinatura_presidente, (x_assinatura_presidente, y_assinatura))
        elif modo_previa:
            fonte_previa = self.fonte_leve(20)
            desenho.text((x_assinatura_presidente + 30, y_assinatura + 60), "[Assin. Presidente]", font=fonte_previa, fill="#AAAAAA")

        y_linha = y_assinatura + altura_assinatura + 10
        fonte_assinatura = self.fonte_negrito(18)

        for x_inicio, texto in [(x_assinatura_membro, "Assinatura do titular"), (x_assinatura_presidente, "Assinatura do Presidente da DBB")]:
            desenho.line([(x_inicio, y_linha), (x_inicio + largura_assinatura, y_linha)], fill="#333333", width=2)
            caixa_texto = desenho.textbbox((0, 0), texto, font=fonte_assinatura)
            largura_texto = caixa_texto[2] - caixa_texto[0]
            x_texto = x_inicio + (largura_assinatura - largura_texto) // 2
            desenho.text((x_texto, y_linha + 8), texto, font=fonte_assinatura, fill="#333333")

        y_texto_vermelho = y_linha + 60
        fonte_vermelha = self.fonte_leve(20)
        texto_vermelho = "Autentique esse documento pelo sítio eletrónico official da Associação:"
        caixa_vermelha = desenho.textbbox((0, 0), texto_vermelho, font=fonte_vermelha)
        largura_vermelha = caixa_vermelha[2] - caixa_vermelha[0]
        desenho.text(((LARGURA_CARTAO - largura_vermelha) // 2, y_texto_vermelho), texto_vermelho, font=fonte_vermelha, fill="#CC0000")

        tamanho_qr = TAMANHO_QR_UNICO
        x_qr = (LARGURA_CARTAO - tamanho_qr) // 2
        y_qr = y_texto_vermelho + 42
        qr_unico = self.carregar_qr_unico()

        if qr_unico is not None:
            imagem.paste(qr_unico, self.posicao_escalada((x_qr, y_qr)))
        elif modo_previa:
            desenho.rectangle([(x_qr, y_qr), (x_qr + tamanho_qr, y_qr + tamanho_qr)], fill="#DDDDDD", outline="#AAAAAA")
            fonte_qr = self.fonte_negrito(22)
            desenho.text((x_qr + 55, y_qr + 78), "QR CODE", font=fonte_qr, fill="#999999")

        self.desenhar_borda_cartao(desenho)
        return imagem

    def gerar_pdf_pessoa(self, frente, verso, pasta_saida, nome_pessoa, nomes_pdf_usados):
        nome_pdf = montar_nome_pdf(nome_pessoa, nomes_pdf_usados)
        caminho_pdf = Path(pasta_saida) / nome_pdf

        escala_pdf = max(1, int(self.configuracao.get("escala_pdf", ESCALA_PDF)))
        largura_pagina = ALTURA_A4 * escala_pdf
        altura_pagina = LARGURA_A4 * escala_pdf
        margem = 120 * escala_pdf
        espaco_cartoes = 120 * escala_pdf

        largura_frente_alvo = int(frente.width * ESCALA_IMPRESSAO_CARTAO)
        altura_frente_alvo = int(frente.height * ESCALA_IMPRESSAO_CARTAO)
        largura_verso_alvo = int(verso.width * ESCALA_IMPRESSAO_CARTAO)
        altura_verso_alvo = int(verso.height * ESCALA_IMPRESSAO_CARTAO)

        largura_total = largura_frente_alvo + largura_verso_alvo + espaco_cartoes
        escala = min(
            (largura_pagina - 2 * margem) / largura_total,
            (altura_pagina - 2 * margem) / max(altura_frente_alvo, altura_verso_alvo),
            1.0,
        )

        largura_frente = int(largura_frente_alvo * escala)
        altura_frente = int(altura_frente_alvo * escala)
        largura_verso = int(largura_verso_alvo * escala)
        altura_verso = int(altura_verso_alvo * escala)
        espaco_final = int(espaco_cartoes * escala)

        x_inicial = (largura_pagina - (largura_frente + largura_verso + espaco_final)) // 2
        y_inicial = (altura_pagina - max(altura_frente, altura_verso)) // 2

        if (largura_frente, altura_frente) != frente.size:
            frente = frente.resize((largura_frente, altura_frente), Image.Resampling.LANCZOS)

        if (largura_verso, altura_verso) != verso.size:
            verso = verso.resize((largura_verso, altura_verso), Image.Resampling.LANCZOS)

        pagina = Image.new("RGB", (largura_pagina, altura_pagina), "white")
        pagina.paste(frente, (x_inicial, y_inicial))
        pagina.paste(verso, (x_inicial + largura_frente + espaco_final, y_inicial))
        pagina.save(caminho_pdf, "PDF", resolution=DPI * escala_pdf, quality=QUALIDADE_PDF, subsampling=0)

        return str(caminho_pdf)

    def processar_todos(self):
        planilha = preparar_planilha(self.configuracao["excel_path"])
        pasta_saida = Path(self.configuracao["saida_dir"])
        pasta_saida.mkdir(parents=True, exist_ok=True)

        pdfs_gerados = []
        nomes_pdf_usados = {}
        total_membros = len(planilha)

        self.registrar_status("Processando...")

        for posicao, (_indice, dados_pessoa) in enumerate(planilha.iterrows(), start=1):
            nome = self.pegar_texto(dados_pessoa, "nome_completo") or "Sem nome"
            self.registrar_status(f"[{posicao}/{total_membros}] {nome}...")

            frente = self.gerar_frente(dados_pessoa)
            verso = self.gerar_verso(dados_pessoa)
            caminho_pdf = self.gerar_pdf_pessoa(frente, verso, pasta_saida, nome, nomes_pdf_usados)

            pdfs_gerados.append(caminho_pdf)
            self.registrar_status(f"PDF: {os.path.basename(caminho_pdf)}")

        self.registrar_status("CONCLUÍDO!")
        return pdfs_gerados


class AplicacaoCarteirinhas(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("System")

        self.title("Gerador de Carteirinhas - DBB")
        self.geometry("660x900")
        self.minsize(520, 480)
        self.resizable(True, True)

        try:
            self.iconbitmap(caminho_recurso("assets/icon.ico"))
        except Exception:
            pass

        self.configuracao = self.criar_configuracao_inicial()
        self.data_emissao_var = ctk.StringVar(value="")
        self.anos_validade_var = ctk.StringVar(value="1")
        self.modo_previa = "frente"
        self.job_previa = None
        self.entradas = {}

        self.data_emissao_var.trace("w", self.sincronizar_emissao)
        self.anos_validade_var.trace("w", self.sincronizar_emissao)

        self.criar_interface()
        self.after(500, self.atualizar_previa)

    def criar_configuracao_inicial(self):
        configuracao = {
            "excel_path": caminho_padrao("exemplos", "dados_teste.xlsx"),
            "logo_path": caminho_padrao("assets", "logo_dbb.png"),
            "fotos_dir": caminho_padrao("exemplos", "fotos"),
            "assinaturas_dir": caminho_padrao("exemplos", "assinaturas"),
            "assinatura_presidente_path": caminho_padrao("assets", "assinatura_presidente.png"),
            "qr_path": "",
            "saida_dir": "",
            "color1": "#009933",
            "color2": "#FFFF00",
            "color3": "#FF0000",
            "data_emissao": "",
            "anos_validade": "1",
            "escala_pdf": ESCALA_PDF,
        }

        return configuracao

    def criar_interface(self):
        self.frame_principal = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.frame_principal.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            self.frame_principal,
            text="Gerador de Carteirinhas - DBB",
            font=("Arial", 21, "bold"),
        ).pack(pady=(0, 12))

        frame_arquivos = ctk.CTkFrame(self.frame_principal)
        frame_arquivos.pack(fill="x", pady=5)

        self.criar_linha_arquivo(frame_arquivos, "Planilha Excel:", "excel_path", [("Excel", "*.xlsx")], mostrar_ajuda=True)
        self.criar_linha_arquivo(frame_arquivos, "Logo:", "logo_path", [("Imagem", "*.png *.jpg *.jpeg")])
        self.criar_linha_diretorio(frame_arquivos, "Pasta de Fotos:", "fotos_dir")
        self.criar_linha_diretorio(frame_arquivos, "Pasta Assinaturas:", "assinaturas_dir")
        self.criar_linha_arquivo(frame_arquivos, "Assin. Presidente:", "assinatura_presidente_path", [("Imagem", "*.png *.jpg *.jpeg")])
        self.criar_linha_arquivo(frame_arquivos, "QR Code Único:", "qr_path", [("Imagem", "*.png *.jpg *.jpeg")])
        self.criar_linha_diretorio(frame_arquivos, "Pasta de Saída:", "saida_dir")

        self.criar_bloco_configuracoes()
        self.criar_bloco_previa()
        self.criar_bloco_final()

    def criar_bloco_configuracoes(self):
        ctk.CTkLabel(self.frame_principal, text="──────────────────────────────────────", text_color="gray").pack(pady=5)
        ctk.CTkLabel(self.frame_principal, text="Configurações", font=("Arial", 13, "bold")).pack(pady=3)

        frame = ctk.CTkFrame(self.frame_principal, fg_color="transparent")
        frame.pack(fill="x", padx=10, pady=4)

        ctk.CTkLabel(frame, text="Data de Emissão:", width=110, anchor="w", font=("Arial", 11)).pack(side="left")
        ctk.CTkEntry(frame, textvariable=self.data_emissao_var, placeholder_text="DD/MM/AAAA", width=110, font=("Arial", 11)).pack(side="left", padx=5)

        ctk.CTkLabel(frame, text="Validade (Anos):", width=100, anchor="w", font=("Arial", 11)).pack(side="left", padx=(15, 0))
        ctk.CTkEntry(frame, textvariable=self.anos_validade_var, width=50, font=("Arial", 11)).pack(side="left", padx=5)

    def criar_bloco_previa(self):
        ctk.CTkLabel(self.frame_principal, text="──────────────────────────────────────", text_color="gray").pack(pady=5)
        ctk.CTkLabel(self.frame_principal, text="Prévia do Resultado:", font=("Arial", 12, "bold")).pack(pady=(2, 3))

        frame_botoes = ctk.CTkFrame(self.frame_principal, fg_color="transparent")
        frame_botoes.pack(pady=(0, 4))

        self.botao_frente = ctk.CTkButton(frame_botoes, text="Frente", width=105, height=28, fg_color="#008800", hover_color="#006600", command=lambda: self.mudar_previa("frente"))
        self.botao_frente.pack(side="left", padx=6)

        self.botao_verso = ctk.CTkButton(frame_botoes, text="Verso", width=105, height=28, fg_color="#555555", hover_color="#333333", command=lambda: self.mudar_previa("verso"))
        self.botao_verso.pack(side="left", padx=6)

        frame_previa = ctk.CTkFrame(self.frame_principal, fg_color="#D8D8D8", corner_radius=10)
        frame_previa.pack(pady=2, fill="x")
        frame_previa.configure(height=250)
        frame_previa.pack_propagate(False)

        self.label_imagem = ctk.CTkLabel(frame_previa, text="")
        self.label_imagem.place(relx=0.5, rely=0.5, anchor="center")

    def criar_bloco_final(self):
        self.label_status = ctk.CTkLabel(self.frame_principal, text="Aguardando...", text_color="gray")
        self.label_status.pack(pady=(5, 0))

        self.botao_gerar = ctk.CTkButton(
            self.frame_principal,
            text="GERAR AGORA",
            height=50,
            fg_color="#008800",
            hover_color="#006600",
            font=("Arial", 16, "bold"),
            command=self.iniciar_thread,
        )
        self.botao_gerar.pack(fill="x", pady=10)

    def texto_ajuda(self):
        return (
            "PADRÃO DE ARQUIVOS\n\n"
            "PLANILHA EXCEL (.xlsx)\n"
            "Colunas obrigatórias: id, nome_completo, profissao, rnm_rg, data_nascimento.\n\n"
            "PASTA DE FOTOS\n"
            "Cada foto deve ter o ID do membro como nome.\n"
            "Exemplo: 1001.jpg ou 1001.png.\n\n"
            "PASTA DE ASSINATURAS\n"
            "Cada assinatura deve ter o ID do membro como nome.\n\n"
            "QR CODE ÚNICO\n"
            "Um único arquivo .png, .jpg ou .jpeg usado em todas as carteirinhas.\n\n"
            "PDF GERADO\n"
            "O programa gera 1 PDF por pessoa, com frente à esquerda e verso à direita."
        )

    def criar_linha_arquivo(self, frame_pai, texto, chave, tipos, mostrar_ajuda=False):
        frame = ctk.CTkFrame(frame_pai, fg_color="transparent")
        frame.pack(fill="x", pady=2, padx=10)

        ctk.CTkLabel(frame, text=texto, width=148, anchor="w", font=("Arial", 11)).pack(side="left")

        if mostrar_ajuda:
            botao_ajuda = ctk.CTkButton(frame, text="?", width=20, height=20, fg_color="#666", font=("Arial", 10, "bold"))
            botao_ajuda.pack(side="left", padx=(0, 5))
            Dica(botao_ajuda, self.texto_ajuda())

        entrada = ctk.CTkEntry(frame, height=28, font=("Arial", 11))
        entrada.pack(side="left", fill="x", expand=True, padx=2)
        entrada.insert(0, self.configuracao.get(chave, ""))
        entrada.configure(state="disabled")

        self.entradas[chave] = entrada

        botao = ctk.CTkButton(frame, text="Selecionar", width=78, height=28, command=lambda: self.selecionar_arquivo(chave, tipos))
        botao.pack(side="right", padx=(5, 0))

    def criar_linha_diretorio(self, frame_pai, texto, chave):
        frame = ctk.CTkFrame(frame_pai, fg_color="transparent")
        frame.pack(fill="x", pady=2, padx=10)

        ctk.CTkLabel(frame, text=texto, width=148, anchor="w", font=("Arial", 11)).pack(side="left")

        entrada = ctk.CTkEntry(frame, height=28, font=("Arial", 11))
        entrada.pack(side="left", fill="x", expand=True, padx=2)
        entrada.insert(0, self.configuracao.get(chave, ""))
        entrada.configure(state="disabled")

        self.entradas[chave] = entrada

        botao = ctk.CTkButton(frame, text="Selecionar", width=78, height=28, command=lambda: self.selecionar_diretorio(chave))
        botao.pack(side="right", padx=(5, 0))

    def atualizar_entrada(self, chave, valor):
        self.configuracao[chave] = valor
        entrada = self.entradas[chave]

        entrada.configure(state="normal")
        entrada.delete(0, "end")
        entrada.insert(0, valor)
        entrada.configure(state="disabled")

        self.atualizar_previa()

    def selecionar_arquivo(self, chave, tipos):
        caminho = filedialog.askopenfilename(filetypes=tipos)

        if caminho:
            self.atualizar_entrada(chave, caminho)

    def selecionar_diretorio(self, chave):
        caminho = filedialog.askdirectory()

        if caminho:
            self.atualizar_entrada(chave, caminho)

    def sincronizar_emissao(self, *_args):
        self.configuracao["data_emissao"] = self.data_emissao_var.get()
        self.configuracao["anos_validade"] = self.anos_validade_var.get()

        if self.job_previa is not None:
            try:
                self.after_cancel(self.job_previa)
            except Exception:
                pass

        self.job_previa = self.after(600, self.atualizar_previa)

    def mudar_previa(self, modo):
        self.modo_previa = modo
        self.botao_frente.configure(fg_color="#008800" if modo == "frente" else "#555555")
        self.botao_verso.configure(fg_color="#008800" if modo == "verso" else "#555555")
        self.atualizar_previa()

    def atualizar_previa(self):
        dados_demo = {
            "nome_completo": "João Silva",
            "profissao": "Motorista",
            "rnm": "123456-A",
            "data_nascimento": "01/01/1990",
            "id": "1001",
        }

        gerador = GeradorCarteirinhas(self.configuracao)

        if self.modo_previa == "verso":
            imagem = gerador.gerar_verso(dados_demo, modo_previa=True)
        else:
            imagem = gerador.gerar_frente(dados_demo, modo_previa=True)

        altura_previa = 220
        largura_previa = int(imagem.width * (altura_previa / imagem.height))
        imagem_tk = ctk.CTkImage(light_image=imagem, dark_image=imagem, size=(largura_previa, altura_previa))
        self.label_imagem.configure(image=imagem_tk)

    def validar_campos(self):
        campos_obrigatorios = {
            "Planilha Excel": self.configuracao["excel_path"],
            "Pasta de Fotos": self.configuracao["fotos_dir"],
            "QR Code Único": self.configuracao["qr_path"],
            "Pasta de Saída": self.configuracao["saida_dir"],
        }

        faltando = []

        for nome_campo, valor in campos_obrigatorios.items():
            if valor == "":
                faltando.append(nome_campo)

        if faltando:
            mensagem = "Configure os campos obrigatórios:\n" + "\n".join(f"- {campo}" for campo in faltando)
            messagebox.showerror("Erro", mensagem)
            return False

        return True

    def iniciar_thread(self):
        if not self.validar_campos():
            return

        self.botao_gerar.configure(state="disabled", text="PROCESSANDO...")

        thread = threading.Thread(target=self.executar_geracao)
        thread.daemon = True
        thread.start()

    def executar_geracao(self):
        try:
            gerador = GeradorCarteirinhas(self.configuracao, self.atualizar_status_seguro)
            pdfs_gerados = gerador.processar_todos()

            self.after(0, lambda: messagebox.showinfo("Sucesso", f"Processo finalizado.\n{len(pdfs_gerados)} PDFs gerados."))
        except Exception as erro:
            mensagem_erro = str(erro)
            self.atualizar_status_seguro(f"ERRO: {mensagem_erro}")
            self.after(0, lambda: messagebox.showerror("Erro", mensagem_erro))
        finally:
            self.after(0, lambda: self.botao_gerar.configure(state="normal", text="GERAR AGORA"))

    def atualizar_status_seguro(self, texto):
        cor = "green" if "CONCLUÍDO" in texto else "blue"
        self.after(0, lambda: self.label_status.configure(text=texto, text_color=cor))


def main():
    app = AplicacaoCarteirinhas()
    app.mainloop()


if __name__ == "__main__":
    main()
