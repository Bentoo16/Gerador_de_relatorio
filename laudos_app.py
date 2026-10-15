"""
Gerador do Instrumento para Análise de Qualidade em Mamografia
==============================================================
Versão desktop (Tkinter) com IA LOCAL via Ollama — sem Streamlit, sem internet,
sem chave de API.

Como usar:
    1) Instale o Ollama (https://ollama.com) e baixe um modelo, ex.:
           ollama pull gemma3:4b
    2) pip install python-docx
    3) python laudos_app.py

Para trocar de modelo sem mexer no código, defina a variável de ambiente
MODELO_LOCAL (ex.: set MODELO_LOCAL=gemma3:12b) ou edite a constante abaixo.
"""
import json
import os
import queue
import re
import subprocess
import sys
import threading
import unicodedata
import urllib.error
import urllib.request
import tkinter as tk
from io import BytesIO
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

# ---------------------------------------------------------------------------
# Configuração da IA local (Ollama)
# ---------------------------------------------------------------------------
OLLAMA_URL = os.environ.get("OLLAMA_URL", "http://localhost:11434")
MODELO_LOCAL = os.environ.get("MODELO_LOCAL", "gemma3:4b")
TEMPERATURA = 0.2            # baixa = menos "criatividade" = menos risco de inventar
TIMEOUT_IA_SEGUNDOS = 300    # a 1ª chamada pode demorar (o modelo é carregado na memória)

# ---------------------------------------------------------------------------
# PROMPTS DA IA — edite aqui se quiser ajustar o comportamento do modelo.
# Modelos pequenos tendem a acrescentar introdução ("Claro! Aqui está..."), a
# inventar conteúdo e a ignorar instruções longas; por isso as regras são curtas,
# numeradas e acompanhadas de um exemplo.
# ---------------------------------------------------------------------------
INSTRUCAO_SISTEMA = (
    "Você revisa textos de relatórios de avaliação de qualidade em mamografia.\n\n"
    "Regras:\n"
    "1. Reescreva as frases recebidas em um único parágrafo coeso, em português do Brasil, com linguagem técnica e formal.\n"
    "2. Use SOMENTE as informações recebidas. Não acrescente, omita nem altere nenhum achado, classificação BI-RADS® ou recomendação.\n"
    "3. Mantenha termos técnicos e siglas exatamente como estão (BI-RADS®, SISMAMA etc.).\n"
    '4. Frases iniciadas por "Detalhe adicional:" e o bloco "Observações do avaliador" são anotações do avaliador: '
    "incorpore-as ao texto sem repetir o rótulo.\n"
    "5. Responda apenas com o parágrafo final: sem introdução, sem comentário, sem título, sem lista, sem markdown.\n"
    "6. Não mencione o número do caso."
)

# Exemplo mostrado ao modelo (entrada -> saída esperada) antes do pedido real.
EXEMPLO_ENTRADA = (
    'Frases a unificar:\n"""\n'
    "No laudo deste exame não foi utilizado corretamente o léxico do BI-RADS® ou do SISMAMA. "
    "Detalhe adicional: a margem da lesão não foi descrita.\n"
    '"""'
)
EXEMPLO_SAIDA = (
    "No laudo deste exame não foi utilizado corretamente o léxico do BI-RADS® ou do SISMAMA. "
    "Além disso, a margem da lesão não foi descrita."
)

INSTRUCAO_SISTEMA_GERAL = (
    "Você resume relatórios de avaliação de qualidade em mamografia.\n\n"
    "Regras:\n"
    "1. Escreva um único parágrafo conciso, em português do Brasil, com linguagem técnica e formal, resumindo os achados gerais.\n"
    "2. Use SOMENTE as informações presentes nos relatórios. Não invente achados, números ou recomendações.\n"
    "3. Não cite números de casos. Se o mesmo achado aparecer em mais de um relatório, pode descrevê-lo como recorrente; "
    "não afirme isso se aparecer em apenas um.\n"
    "4. Se nenhum relatório apontar problemas, diga apenas que não foram identificadas considerações específicas nos casos avaliados.\n"
    "5. Responda apenas com o parágrafo: sem introdução, sem comentário, sem título, sem lista, sem markdown."
)


class ErroIA(Exception):
    pass


def gerar_texto(prompt, sistema=INSTRUCAO_SISTEMA, exemplo=None):
    """Envia o pedido ao Ollama local e devolve o texto gerado.
    `exemplo` é um par opcional (entrada, saída) mostrado ao modelo antes do pedido real."""
    mensagens = [{"role": "system", "content": sistema}]
    if exemplo:
        mensagens.append({"role": "user", "content": exemplo[0]})
        mensagens.append({"role": "assistant", "content": exemplo[1]})
    mensagens.append({"role": "user", "content": prompt})
    payload = {
        "model": MODELO_LOCAL,
        "messages": mensagens,
        "stream": False,
        "options": {"temperature": TEMPERATURA},
    }
    req = urllib.request.Request(
        f"{OLLAMA_URL}/api/chat",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT_IA_SEGUNDOS) as resp:
            dados = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        corpo = e.read().decode("utf-8", errors="ignore")
        if e.code == 404:
            raise ErroIA(
                f"O modelo '{MODELO_LOCAL}' não foi encontrado.\n"
                f"Abra o terminal e rode:  ollama pull {MODELO_LOCAL}"
            )
        raise ErroIA(f"Erro do Ollama ({e.code}): {corpo}")
    except urllib.error.URLError:
        raise ErroIA(
            "Não consegui conectar ao Ollama.\n"
            "Verifique se ele está instalado e rodando (abra o aplicativo Ollama "
            "ou rode 'ollama serve' no terminal)."
        )
    except TimeoutError:
        raise ErroIA("A IA demorou demais para responder. Tente novamente.")

    texto = dados.get("message", {}).get("content", "")
    # Alguns modelos (ex.: Qwen3) devolvem o "raciocínio" entre <think>...</think>.
    texto = re.sub(r"<think>.*?</think>", "", texto, flags=re.DOTALL).strip()
    if not texto:
        raise ErroIA("A IA devolveu uma resposta vazia.")
    return texto


# ---------------------------------------------------------------------------
# Funções auxiliares
# ---------------------------------------------------------------------------
def extrair_numero(nome):
    match = re.search(r"\d+", nome)
    return int(match.group()) if match else 0


def sanitizar_nome_arquivo(texto):
    """Transforma um texto livre (ex.: nome do serviço) em algo seguro para
    usar como nome de arquivo: sem acentos, sem caracteres especiais e com
    espaços trocados por underscore."""
    if not texto or not texto.strip():
        return "Servico"
    texto_sem_acento = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto_limpo = re.sub(r"[^\w\s-]", "", texto_sem_acento).strip()
    texto_final = re.sub(r"[-\s]+", "_", texto_limpo)
    return texto_final or "Servico"


def limpar_formatacao(texto):
    return texto.replace("**", "").replace("__", "").replace("#", "")


def abrir_arquivo(caminho):
    try:
        if sys.platform.startswith("win"):
            os.startfile(caminho)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", caminho])
        else:
            subprocess.Popen(["xdg-open", caminho])
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Biblioteca de perguntas (organizada por grupos)
# ---------------------------------------------------------------------------
GRUPO_LAUDO = "Avaliação dos Critérios de Laudos"

perguntas = {
    GRUPO_LAUDO: {
        "Resumo da história presente": {
            "opcoes": {"Sim": " ", "Não": "É importante que nos laudos conste a indicação do exame. Essa indicação deve conter uma história resumida da paciente (exame de rastreamento x diagnóstico / história familiar / antecedentes cirúrgicos e resultados de biópsias / sintomas e queixas da paciente ... )."},
        },
        "Utiliza corretamente o Léxico BI-RADS® ou SISMAMA": {
            "opcoes": {"Sim": " ", "Não": "No laudo deste exame não foi utilizado corretamente o léxico do BI-RADS® ou do SISMAMA."},
        },
        "Classifica corretamente o exame segundo o BI-RADS®": {
            "opcoes": {"Sim": " ", "Não": "O exame não foi classificado corretamente."},
        },
        "Recomendação correta segundo o BI-RADS®": {
            "opcoes": {"Sim": " ", "Não": " No laudo enviado para avaliação, o exame não foi classificado corretamente."},
        },
        "Interpretou corretamente todos os achados do exame": {
            "opcoes": {"Sim": "", "Não": " Os achados do exame não foram interpretados corretamente."},
        },
    },
}

# Perguntas adicionais (texto livre) - não entram no texto enviado à IA,
# servem apenas para compor uma tabela própria no documento .docx.
perguntas_adicionais_texto = [
    "Padrão de Mama segundo o serviço: (MARCAR: A, B, C ou D)",
    "Tipo de achado segundo o serviço: (MARCAR: MCF, CALC, NOD, DISTORC, ASSIM, etc)",
    "Classificação BI-RADS® do serviço: (MARCAR: 0, 1, 2, 3, 4, 5, 6)",
    "Classificação BI-RADS® dos avaliadores: (MARCAR: 0, 1, 2, 3, 4, 5, 6)",
]

# ---------------------------------------------------------------------------
# Combinações de respostas: quando um CONJUNTO de perguntas específicas
# recebe a mesma resposta (ex.: "Não" em duas perguntas ao mesmo tempo),
# usa-se UMA frase conjunta no lugar das frases individuais de cada
# pergunta. A frase aparece só uma vez.
# Para CRIAR uma nova combinação, copie um bloco e ajuste "perguntas",
# "resposta_gatilho" e "texto". Para REMOVER, apague o bloco.
# ---------------------------------------------------------------------------
COMBINACOES_RESPOSTA = [
    {
        "perguntas": [
            "Classifica corretamente o exame segundo o BI-RADS®",
            "Recomendação correta segundo o BI-RADS®",
        ],
        "resposta_gatilho": "Não",
        "texto": "O exame não foi classificado corretamente e, consequentemente, a recomendação de conduta também não está correta.",
    },
]

# ---------------------------------------------------------------------------
# Recomendações do documento
# Estrutura: {pergunta: {sub_opcao: texto}}.
# Perguntas sem sub_opcoes usam a chave "_default".
# ---------------------------------------------------------------------------
RECOMENDACOES = {
    "Utiliza corretamente o Léxico BI-RADS® ou SISMAMA": {
        "_default":
            "É recomendado aos médicos do serviço a realização do curso de reciclagem em diagnóstico mamário “Atualização em BI-RADS”, oferecido pelo Colégio Brasileiro de Radiologia. O curso é gratuito para os médicos dos serviços que participam do Programa de Qualidade em Mamografia do INCA. O detalhamento do processo para realização do curso indicado será encaminhado em um e-mail à parte, que tratará exclusivamente desse assunto."
    },
}

# Recomendações "por grupo": disparam se QUALQUER pergunta de um grupo (menos
# as excluídas) tiver a resposta indicada em "resposta_gatilho". Aparecem só
# UMA VEZ no documento. Campos: "grupo", "resposta_gatilho", "texto" e
# "perguntas_excluidas" OU "perguntas_incluidas" (esta última tem prioridade).
RECOMENDACOES_POR_GRUPO = [
    # Nenhuma ativa no momento.
]

TEXTO_LAUDO_SEM_CONSIDERACOES = "Sem considerações específicas sobre o laudo deste caso."
TEXTO_SEM_CONSIDERACOES = "Sem considerações específicas sobre o caso."
TOTAL_CASOS = 5
OPCOES_TIPO = ["Convencional", "Digital CR", "Digital DR", "DR retrofit"]

# ---------------------------------------------------------------------------
# Estado da sessão (substitui o st.session_state)
# ---------------------------------------------------------------------------
CABECALHO_VAZIO = {
    "mamografo_fabricante": "",
    "mamografo_modelo": "",
    "cnes": "",
    "qiid": "",
    "tipo_mamografo": OPCOES_TIPO[0],
    "instituicao": "",
    "cidade": "",
    "estado": "",
    "servico": "",
}


class Estado:
    def __init__(self):
        self.resetar()

    def resetar(self):
        self.dados_cabecalho = dict(CABECALHO_VAZIO)
        self.casos_salvos = {}
        self.relatorios_ia = {}
        self.relatorio_geral_salvo = None
        self.consideracoes_caso = {}
        self.consideracoes_gerais = ""
        self.escolhas_casos = {}
        self.identificacao_exames = {}
        self.dados_adicionais_casos = {}


# ---------------------------------------------------------------------------
# Lógica de montagem do texto de um caso (idêntica à versão Streamlit)
# ---------------------------------------------------------------------------
def _detalhe(obs):
    """Formata a observação de uma pergunta, garantindo ponto final para que a IA
    saiba onde ela termina (senão cola na frase seguinte)."""
    obs = obs.strip()
    if obs and obs[-1] not in ".!?;:":
        obs += "."
    return f"Detalhe adicional: {obs}"


def montar_caso(respostas_temporarias, consideracoes_caso):
    """Recebe a lista [{titulo, escolha, sub_escolha, obs}] e devolve
    (texto_bruto, texto_para_ia, escolhas)."""
    titulos_grupo_laudo = set(perguntas[GRUPO_LAUDO].keys())
    respostas_grupo_laudo = {
        item["titulo"]: item["escolha"]
        for item in respostas_temporarias
        if item["titulo"] in titulos_grupo_laudo
    }
    laudo_tudo_sim = (
        len(respostas_grupo_laudo) == len(titulos_grupo_laudo)
        and all(escolha == "Sim" for escolha in respostas_grupo_laudo.values())
    )

    respostas_por_titulo = {item["titulo"]: item["escolha"] for item in respostas_temporarias}
    combinacoes_disparadas = [
        combinacao
        for combinacao in COMBINACOES_RESPOSTA
        if all(
            respostas_por_titulo.get(titulo) == combinacao["resposta_gatilho"]
            for titulo in combinacao["perguntas"]
        )
    ]
    titulos_cobertos_por_combinacao = {
        titulo for combinacao in combinacoes_disparadas for titulo in combinacao["perguntas"]
    }

    respostas_finais = []
    laudo_texto_inserido = False
    combinacoes_inseridas = set()
    for item in respostas_temporarias:
        if item["titulo"] in titulos_grupo_laudo and laudo_tudo_sim:
            # Todas "Sim": em vez dos textos em branco, insere uma única frase fixa.
            if not laudo_texto_inserido:
                respostas_finais.append(TEXTO_LAUDO_SEM_CONSIDERACOES)
                laudo_texto_inserido = True
            if item["obs"]:
                respostas_finais.append(_detalhe(item["obs"]))
            continue
        if item["titulo"] in titulos_cobertos_por_combinacao:
            # Faz parte de uma combinação disparada: frase conjunta uma única vez.
            combinacao = next(c for c in combinacoes_disparadas if item["titulo"] in c["perguntas"])
            chave_combinacao = tuple(combinacao["perguntas"])
            if chave_combinacao not in combinacoes_inseridas:
                respostas_finais.append(combinacao["texto"])
                combinacoes_inseridas.add(chave_combinacao)
            if item["obs"]:
                respostas_finais.append(_detalhe(item["obs"]))
            continue
        for questoes in perguntas.values():
            if item["titulo"] in questoes:
                info_pergunta = questoes[item["titulo"]]
                break
        gatilho_pergunta = info_pergunta.get("gatilho_sub_opcoes", "Não")
        if item["escolha"] == gatilho_pergunta and item["sub_escolha"]:
            frase_base = " ".join(info_pergunta["sub_opcoes"][opcao] for opcao in item["sub_escolha"])
        else:
            frase_base = info_pergunta["opcoes"][item["escolha"]]
        if item["obs"]:
            frase_base += f" {_detalhe(item['obs'])}"
        respostas_finais.append(frase_base)

    texto_bruto = " ".join(respostas_finais)
    texto_para_ia = texto_bruto
    if consideracoes_caso.strip():
        texto_para_ia += f"\n\n{consideracoes_caso}"

    escolhas = {
        item["titulo"]: {"resposta": item["escolha"], "sub_opcao": item["sub_escolha"]}
        for item in respostas_temporarias
    }
    return texto_bruto, texto_para_ia, escolhas


def montar_pedido_caso(texto_bruto, consideracoes_caso):
    """Monta o pedido à IA para um caso: frases e observações em blocos separados."""
    partes = []
    frases = re.sub(r"\s+", " ", texto_bruto).strip()
    if frases:
        partes.append(f'Frases a unificar:\n"""\n{frases}\n"""')
    if consideracoes_caso.strip():
        partes.append(f'Observações do avaliador:\n"""\n{consideracoes_caso.strip()}\n"""')
    return {
        "prompt": "\n\n".join(partes),
        "sistema": INSTRUCAO_SISTEMA,
        "exemplo": (EXEMPLO_ENTRADA, EXEMPLO_SAIDA),
    }


def montar_pedido_geral(estado):
    """Monta o pedido do relatório geral a partir dos relatórios individuais (já revisados)."""
    casos = sorted(estado.relatorios_ia.keys(), key=extrair_numero)
    blocos = [f"[{c}]\n{estado.relatorios_ia[c].strip()}" for c in casos]
    prompt = "Relatórios individuais:\n\"\"\"\n" + "\n\n".join(blocos) + "\n\"\"\""
    if estado.consideracoes_gerais.strip():
        prompt += f'\n\nObservações gerais do avaliador:\n"""\n{estado.consideracoes_gerais.strip()}\n"""'
    return {"prompt": prompt, "sistema": INSTRUCAO_SISTEMA_GERAL, "exemplo": None}


# ---------------------------------------------------------------------------
# Geração do documento Word
# ---------------------------------------------------------------------------
COR_PRIMARIA = "1F4E79"    # azul escuro - cabeçalho principal das tabelas
COR_SECUNDARIA = "2E75B6"  # azul médio - linha "Caso" das tabelas
COR_ZEBRA = "F4F7FB"       # cinza-azulado bem claro - linhas alternadas


def set_cell_shading(cell, color):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color)
    shading.set(qn("w:val"), "clear")
    cell._tc.get_or_add_tcPr().append(shading)


def estilizar_documento(doc):
    """Define fonte base e cores dos títulos para todo o documento."""
    estilo_normal = doc.styles["Normal"]
    estilo_normal.font.name = "Calibri"
    estilo_normal.font.size = Pt(11)

    tamanhos_titulo = {0: 20, 1: 15, 2: 13}
    for nivel, tamanho in tamanhos_titulo.items():
        estilo = doc.styles[f"Heading {nivel}" if nivel > 0 else "Title"]
        estilo.font.name = "Calibri"
        estilo.font.size = Pt(tamanho)
        estilo.font.color.rgb = RGBColor.from_string(COR_PRIMARIA)
        estilo.font.bold = True


def formatar_celula(cell, texto=None, cor_fundo=None, cor_texto=None, negrito=False, centralizar=True):
    """Preenche uma célula de tabela com texto formatado (cor, negrito, fundo)."""
    if texto is not None:
        cell.text = ""
        paragrafo = cell.paragraphs[0]
        run = paragrafo.add_run(texto)
    else:
        paragrafo = cell.paragraphs[0]
        run = paragrafo.runs[0] if paragrafo.runs else paragrafo.add_run("")
    run.bold = negrito
    if cor_texto:
        run.font.color.rgb = RGBColor.from_string(cor_texto)
    if centralizar:
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if cor_fundo:
        set_cell_shading(cell, cor_fundo)


def criar_docx(estado):
    """Monta o documento Word e devolve os bytes do arquivo."""
    doc = Document()
    estilizar_documento(doc)

    doc.add_heading("Instrumento para a análise da qualidade da mamografia", level=0)
    doc.add_paragraph()

    cab = estado.dados_cabecalho
    p = doc.add_paragraph()
    p.add_run("Mamógrafo (fabricante e modelo): ").bold = True
    p.add_run(f"{cab['mamografo_fabricante']} - {cab['mamografo_modelo']}")

    p = doc.add_paragraph()
    p.add_run("CNES: ").bold = True
    p.add_run(cab["cnes"])
    p.add_run("     QIID: ").bold = True
    p.add_run(cab["qiid"])

    p = doc.add_paragraph()
    p.add_run("Tipo de mamógrafo: ").bold = True
    for opcao in OPCOES_TIPO:
        marcado = "X" if cab["tipo_mamografo"] == opcao else "-"
        p.add_run(f"  [{marcado}] {opcao}  ")

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Instituição: ").bold = True
    p.add_run(cab["instituicao"])

    p = doc.add_paragraph()
    p.add_run("Cidade: ").bold = True
    p.add_run(cab["cidade"])
    p.add_run("     Estado: ").bold = True
    p.add_run(cab["estado"])
    doc.add_paragraph()

    # Tabelas de respostas - uma por grupo
    casos_ord = sorted(estado.relatorios_ia.keys(), key=extrair_numero)
    num_casos = len(casos_ord)

    for nome_grupo, questoes in perguntas.items():
        doc.add_heading(f"Tabela de Respostas - {nome_grupo}", level=1)
        perguntas_ord = list(questoes.keys())
        total_colunas = 1 + num_casos * 2
        tabela = doc.add_table(rows=2 + len(perguntas_ord), cols=total_colunas)
        tabela.style = "Table Grid"

        tabela.cell(0, 0).merge(tabela.cell(1, 0))
        formatar_celula(tabela.cell(0, 0), "Pergunta", cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)

        for idx, caso in enumerate(casos_ord):
            col_inicio = 1 + idx * 2
            col_fim = col_inicio + 1
            tabela.cell(0, col_inicio).merge(tabela.cell(0, col_fim))
            formatar_celula(tabela.cell(0, col_inicio), caso, cor_fundo=COR_SECUNDARIA, cor_texto="FFFFFF", negrito=True)

        for idx in range(num_casos):
            col_sim = 1 + idx * 2
            col_nao = col_sim + 1
            formatar_celula(tabela.cell(1, col_sim), "Sim", cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)
            formatar_celula(tabela.cell(1, col_nao), "Não", cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)

        for i, pergunta in enumerate(perguntas_ord):
            linha_atual = i + 2
            cor_linha = COR_ZEBRA if i % 2 == 1 else None
            formatar_celula(tabela.cell(linha_atual, 0), pergunta, cor_fundo=cor_linha, centralizar=False)
            for j, caso in enumerate(casos_ord):
                item = estado.escolhas_casos.get(caso, {}).get(pergunta, {})
                resposta = item.get("resposta", "-") if isinstance(item, dict) else item
                col_sim = 1 + j * 2
                col_nao = col_sim + 1
                formatar_celula(tabela.cell(linha_atual, col_sim), "X" if resposta == "Sim" else "", cor_fundo=cor_linha, negrito=True)
                formatar_celula(tabela.cell(linha_atual, col_nao), "X" if resposta == "Não" else "", cor_fundo=cor_linha, negrito=True)

        doc.add_paragraph()

    doc.add_page_break()

    # Identificação dos exames
    doc.add_heading("Identificação dos Exames", level=2)
    mini_tabela = doc.add_table(rows=2, cols=num_casos)
    mini_tabela.style = "Table Grid"
    for j, caso in enumerate(casos_ord):
        formatar_celula(mini_tabela.rows[0].cells[j], caso, cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)
    for j, caso in enumerate(casos_ord):
        formatar_celula(mini_tabela.rows[1].cells[j], estado.identificacao_exames.get(caso, ""))
    doc.add_paragraph()

    # Dados adicionais dos casos (não entram no texto da IA)
    doc.add_heading("Dados Adicionais dos Casos", level=2)
    tabela_adicional = doc.add_table(rows=1 + len(perguntas_adicionais_texto), cols=1 + num_casos)
    tabela_adicional.style = "Table Grid"
    formatar_celula(tabela_adicional.cell(0, 0), "Pergunta", cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)
    for j, caso in enumerate(casos_ord):
        formatar_celula(tabela_adicional.cell(0, j + 1), caso, cor_fundo=COR_PRIMARIA, cor_texto="FFFFFF", negrito=True)
    for i, pergunta_extra in enumerate(perguntas_adicionais_texto):
        linha_atual = i + 1
        cor_linha = COR_ZEBRA if i % 2 == 1 else None
        formatar_celula(tabela_adicional.cell(linha_atual, 0), pergunta_extra, cor_fundo=cor_linha, centralizar=False)
        for j, caso in enumerate(casos_ord):
            valor = estado.dados_adicionais_casos.get(caso, {}).get(pergunta_extra, "")
            formatar_celula(tabela_adicional.cell(linha_atual, j + 1), valor, cor_fundo=cor_linha)
    doc.add_paragraph()

    # Anexo
    doc.add_heading("Anexo ao instrumento para análise da qualidade da mamografia", level=0)
    p = doc.add_paragraph()
    p.add_run("Serviço: ").bold = True
    p.add_run(cab["servico"])
    p = doc.add_paragraph()
    p.add_run("CNES: ").bold = True
    p.add_run(cab["cnes"])
    p = doc.add_paragraph()
    p.add_run("QIID: ").bold = True
    p.add_run(cab["qiid"])
    doc.add_paragraph()

    # Considerações específicas
    doc.add_heading("Considerações Específicas", level=0)
    for nome_caso in casos_ord:
        doc.add_heading(nome_caso, level=1)
        texto_ia = estado.relatorios_ia[nome_caso]
        for linha in texto_ia.strip().split("\n"):
            if linha.strip():
                doc.add_paragraph(limpar_formatacao(linha))
        if estado.consideracoes_caso.get(nome_caso, "").strip():
            doc.add_heading("Considerações Adicionais", level=2)
            doc.add_paragraph(limpar_formatacao(estado.consideracoes_caso[nome_caso]))
        doc.add_paragraph("-" * 30)

    # Recomendações
    def resposta_do_caso(caso, pergunta):
        item = estado.escolhas_casos.get(caso, {}).get(pergunta, {})
        return item.get("resposta", "") if isinstance(item, dict) else item

    def sub_opcao_do_caso(caso, pergunta):
        item = estado.escolhas_casos.get(caso, {}).get(pergunta, {})
        return item.get("sub_opcao") if isinstance(item, dict) else None

    def obter_gatilho(pergunta):
        for questoes in perguntas.values():
            if pergunta in questoes:
                return questoes[pergunta].get("gatilho_sub_opcoes", "Não")
        return "Não"

    def perguntas_da_config_grupo(config):
        questoes_grupo = perguntas.get(config["grupo"], {})
        incluidas = config.get("perguntas_incluidas") or []
        if incluidas:
            return incluidas
        excluidas = set(config.get("perguntas_excluidas") or [])
        return [p for p in questoes_grupo if p not in excluidas]

    def grupo_disparado(caso, config):
        resposta_gatilho = config.get("resposta_gatilho", "Não")
        return any(
            resposta_do_caso(caso, pergunta) == resposta_gatilho
            for pergunta in perguntas_da_config_grupo(config)
        )

    tem_recomendacao = any(
        resposta_do_caso(caso, pergunta) == obter_gatilho(pergunta)
        for caso in casos_ord
        for pergunta in RECOMENDACOES
    ) or any(
        grupo_disparado(caso, config)
        for caso in casos_ord
        for config in RECOMENDACOES_POR_GRUPO
    )
    if tem_recomendacao:
        doc.add_heading("Recomendações", level=0)
        textos_inseridos = set()
        for caso in casos_ord:
            for pergunta, textos_por_sub in RECOMENDACOES.items():
                if resposta_do_caso(caso, pergunta) == obter_gatilho(pergunta):
                    subs = sub_opcao_do_caso(caso, pergunta)
                    if not subs:
                        textos = [textos_por_sub.get("_default")]
                    else:
                        textos = [textos_por_sub.get(sub) or textos_por_sub.get("_default") for sub in subs]
                    for texto in textos:
                        if texto and texto not in textos_inseridos:
                            textos_inseridos.add(texto)
                            p = doc.add_paragraph(style="List Bullet")
                            p.add_run(texto)

        for caso in casos_ord:
            for config in RECOMENDACOES_POR_GRUPO:
                if grupo_disparado(caso, config):
                    texto = config.get("texto")
                    if texto and texto not in textos_inseridos:
                        textos_inseridos.add(texto)
                        p = doc.add_paragraph(style="List Bullet")
                        p.add_run(texto)

    if estado.relatorio_geral_salvo:
        doc.add_paragraph()
        p = doc.add_paragraph()
        p.add_run("Resumo dos casos:").bold = True
        for linha in estado.relatorio_geral_salvo.strip().split("\n"):
            if linha.strip():
                doc.add_paragraph(limpar_formatacao(linha))

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


# ---------------------------------------------------------------------------
# Interface gráfica (Tkinter)
# ---------------------------------------------------------------------------
class ScrollFrame(ttk.Frame):
    """Frame com barra de rolagem vertical. Coloque os widgets em .interior"""

    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.interior = ttk.Frame(self.canvas, padding=14)
        self._janela = self.canvas.create_window((0, 0), window=self.interior, anchor="nw")
        self.interior.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._janela, width=e.width))


class CasoForm:
    """Formulário de um caso (uma aba). Guarda as próprias variáveis, então
    trocar de aba não perde o que foi digitado."""

    def __init__(self, app, parent, numero):
        self.app = app
        self.numero = numero
        self.nome = f"Caso {numero}"
        self.scroll = ScrollFrame(parent)
        raiz = self.scroll.interior
        raiz.columnconfigure(0, weight=1)

        # --- Dados adicionais
        ttk.Label(raiz, text="Dados Adicionais do Caso", style="Secao.TLabel").pack(anchor="w")
        self.vars_extras = {}
        for pergunta_extra in perguntas_adicionais_texto:
            ttk.Label(raiz, text=pergunta_extra, wraplength=820, justify="left").pack(anchor="w", pady=(6, 0))
            var = tk.StringVar()
            ttk.Entry(raiz, textvariable=var).pack(fill="x")
            self.vars_extras[pergunta_extra] = var

        # --- Perguntas por grupo
        self.vars_perguntas = {}  # titulo -> dict(escolha, obs, subs, gatilho)
        for nome_grupo, questoes in perguntas.items():
            grupo = ttk.LabelFrame(raiz, text=f" {nome_grupo} ", padding=10)
            grupo.pack(fill="x", pady=(16, 0))
            for titulo, info in questoes.items():
                self._criar_pergunta(grupo, titulo, info)

        # --- Fechamento do caso
        ttk.Label(raiz, text="Identificação do Exame:", style="Secao.TLabel").pack(anchor="w", pady=(16, 0))
        self.var_id_exame = tk.StringVar()
        ttk.Entry(raiz, textvariable=self.var_id_exame).pack(fill="x")

        ttk.Label(raiz, text="Considerações adicionais para este caso (opcional):").pack(anchor="w", pady=(10, 0))
        self.txt_consideracoes = tk.Text(raiz, height=4, wrap="word", font=app.fonte_base, relief="solid", borderwidth=1)
        self.txt_consideracoes.pack(fill="x")

        ttk.Button(
            raiz, text=f"Analisar e Salvar {self.nome}", style="Primario.TButton",
            command=lambda: app.salvar_caso(numero),
        ).pack(fill="x", pady=(16, 4))

    def _criar_pergunta(self, parent, titulo, info):
        bloco = ttk.Frame(parent)
        bloco.pack(fill="x", pady=(0, 10))
        bloco.columnconfigure(0, weight=1)

        ttk.Label(bloco, text=titulo, style="Pergunta.TLabel", wraplength=800, justify="left").grid(row=0, column=0, sticky="w")

        escolha = tk.StringVar(value=list(info["opcoes"].keys())[0])
        linha_radio = ttk.Frame(bloco)
        linha_radio.grid(row=1, column=0, sticky="w")
        for opcao in info["opcoes"]:
            ttk.Radiobutton(linha_radio, text=opcao, value=opcao, variable=escolha).pack(side="left", padx=(0, 18))

        gatilho = info.get("gatilho_sub_opcoes", "Não")
        subs = {}
        if "sub_opcoes" in info:
            frame_subs = ttk.Frame(bloco)
            frame_subs.grid(row=2, column=0, sticky="w", padx=(18, 0))
            ttk.Label(frame_subs, text="Especifique:").pack(anchor="w")
            for opcao in info["sub_opcoes"]:
                subs[opcao] = tk.BooleanVar(value=False)
                ttk.Checkbutton(frame_subs, text=opcao, variable=subs[opcao]).pack(anchor="w")
            frame_subs.grid_remove()

            def alternar(*_args, f=frame_subs, v=escolha, g=gatilho):
                if v.get() == g:
                    f.grid()
                else:
                    f.grid_remove()

            escolha.trace_add("write", alternar)

        ttk.Label(bloco, text="Considerações adicionais (opcional):").grid(row=3, column=0, sticky="w", pady=(4, 0))
        obs = tk.StringVar()
        ttk.Entry(bloco, textvariable=obs).grid(row=4, column=0, sticky="ew")

        self.vars_perguntas[titulo] = {"escolha": escolha, "obs": obs, "subs": subs, "gatilho": gatilho, "info": info}

    # --- leitura / escrita dos campos
    def coletar_respostas(self):
        respostas = []
        for titulo, v in self.vars_perguntas.items():
            escolha = v["escolha"].get()
            sub_escolha = []
            if v["subs"] and escolha == v["gatilho"]:
                sub_escolha = [o for o, var in v["subs"].items() if var.get()]
            respostas.append({
                "titulo": titulo,
                "escolha": escolha,
                "sub_escolha": sub_escolha,
                "obs": v["obs"].get().strip(),
            })
        return respostas

    def coletar_adicionais(self):
        return {p: var.get() for p, var in self.vars_extras.items()}

    def consideracoes(self):
        return self.txt_consideracoes.get("1.0", "end").strip()

    def marcar_tudo_sim(self):
        for v in self.vars_perguntas.values():
            v["escolha"].set("Sim")

    def limpar(self):
        for var in self.vars_extras.values():
            var.set("")
        for v in self.vars_perguntas.values():
            v["escolha"].set(list(v["info"]["opcoes"].keys())[0])
            v["obs"].set("")
            for var in v["subs"].values():
                var.set(False)
        self.var_id_exame.set("")
        self.txt_consideracoes.delete("1.0", "end")


class RevisaoDialog(tk.Toplevel):
    """Janela para conferir e editar o texto gerado pela IA antes de salvar."""

    def __init__(self, app, titulo, origem, texto_ia):
        super().__init__(app)
        self.acao = "cancelar"
        self.texto = None
        self.title(titulo)
        self.geometry("800x620")
        self.minsize(620, 480)
        self.transient(app)

        corpo = ttk.Frame(self, padding=14)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text="Confira se a IA não alterou nenhum achado clínico antes de salvar.",
                  style="Secao.TLabel", wraplength=740, justify="left").pack(anchor="w")

        if origem.strip():
            ttk.Label(corpo, text="Texto de origem (suas respostas — somente leitura):").pack(anchor="w", pady=(10, 2))
            txt_origem = tk.Text(corpo, height=5, wrap="word", font=app.fonte_base, relief="solid",
                                 borderwidth=1, background=f"#{COR_ZEBRA}")
            txt_origem.insert("1.0", origem.strip())
            txt_origem.configure(state="disabled")
            txt_origem.pack(fill="x")

        ttk.Label(corpo, text="Texto gerado pela IA (você pode editar):").pack(anchor="w", pady=(10, 2))
        botoes = ttk.Frame(corpo)
        botoes.pack(side="bottom", fill="x", pady=(12, 0))
        quadro = ttk.Frame(corpo)
        quadro.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(quadro, orient="vertical")
        self.txt = tk.Text(quadro, wrap="word", font=app.fonte_base, relief="solid", borderwidth=1, yscrollcommand=sb.set)
        sb.configure(command=self.txt.yview)
        sb.pack(side="right", fill="y")
        self.txt.pack(side="left", fill="both", expand=True)
        self.txt.insert("1.0", texto_ia)

        ttk.Button(botoes, text="Salvar este texto", style="Primario.TButton", command=self.salvar).pack(side="left")
        ttk.Button(botoes, text="Gerar novamente", style="Secundario.TButton", command=self.regerar).pack(side="left", padx=8)
        ttk.Button(botoes, text="Cancelar", style="Secundario.TButton", command=self.cancelar).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self.cancelar)
        self.update_idletasks()
        x = app.winfo_rootx() + max(0, (app.winfo_width() - self.winfo_width()) // 2)
        y = app.winfo_rooty() + 40
        self.geometry(f"+{x}+{y}")
        self.wait_visibility()
        self.grab_set()
        self.txt.focus_set()

    def salvar(self):
        texto = self.txt.get("1.0", "end").strip()
        if not texto:
            messagebox.showwarning("Texto vazio", "O texto não pode ficar vazio. Edite-o ou use 'Gerar novamente'.", parent=self)
            return
        self.acao, self.texto = "salvar", texto
        self.destroy()

    def regerar(self):
        self.acao = "regerar"
        self.destroy()

    def cancelar(self):
        self.acao = "cancelar"
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador do Instrumento para Análise de Qualidade em Mamografia")
        self.geometry("1020x820")
        self.minsize(900, 640)
        self.estado = Estado()
        self._ia_ocupada = False
        self.forms = {}

        self.fonte_base = tkfont.nametofont("TkDefaultFont").copy()
        self.fonte_base.configure(size=10)
        self._configurar_estilo()
        self._criar_topo()
        self._criar_rodape()

        self.abas = ttk.Notebook(self)
        self.abas.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self._criar_aba_cabecalho()
        self._criar_aba_casos()
        self._criar_aba_resumo()
        self.atualizar_progresso()

        self.bind_all("<MouseWheel>", self._on_roda_mouse)
        self.bind_all("<Button-4>", self._on_roda_mouse)
        self.bind_all("<Button-5>", self._on_roda_mouse)

    # ------------------------------------------------------------------ estilo
    def _configurar_estilo(self):
        familia = self.fonte_base.actual("family")
        est = ttk.Style(self)
        try:
            est.theme_use("clam")
        except tk.TclError:
            pass
        est.configure(".", font=(familia, 10))
        est.configure("TNotebook.Tab", padding=(16, 7), font=(familia, 10, "bold"))
        est.configure("Titulo.TLabel", font=(familia, 16, "bold"), foreground=f"#{COR_PRIMARIA}")
        est.configure("Secao.TLabel", font=(familia, 11, "bold"), foreground=f"#{COR_PRIMARIA}")
        est.configure("Pergunta.TLabel", font=(familia, 10, "bold"))
        est.configure("TLabelframe.Label", font=(familia, 10, "bold"), foreground=f"#{COR_PRIMARIA}")
        est.configure("Primario.TButton", background=f"#{COR_PRIMARIA}", foreground="white",
                      font=(familia, 10, "bold"), padding=9, borderwidth=0)
        est.map("Primario.TButton", background=[("active", f"#{COR_SECUNDARIA}"), ("disabled", "#9aa8b8")])
        est.configure("Secundario.TButton", padding=8)
        est.configure("Treeview", rowheight=26)
        est.configure("Treeview.Heading", background=f"#{COR_PRIMARIA}", foreground="white", font=(familia, 10, "bold"))
        est.map("Treeview.Heading", background=[("active", f"#{COR_SECUNDARIA}")])

    def _criar_topo(self):
        topo = ttk.Frame(self, padding=(14, 10, 14, 6))
        topo.pack(fill="x")
        ttk.Label(topo, text="Gerador do Instrumento para Análise de Qualidade em Mamografia",
                  style="Titulo.TLabel").pack(anchor="w")
        linha = ttk.Frame(topo)
        linha.pack(fill="x", pady=(6, 0))
        self.lbl_progresso = ttk.Label(linha, text="")
        self.lbl_progresso.pack(side="left")
        self.barra = ttk.Progressbar(linha, maximum=TOTAL_CASOS, length=260)
        self.barra.pack(side="left", padx=10)
        ttk.Label(linha, text=f"IA local: {MODELO_LOCAL}", foreground="#666666").pack(side="right")

    def _criar_rodape(self):
        rodape = ttk.Frame(self, padding=(14, 0, 14, 8))
        rodape.pack(side="bottom", fill="x")
        self.var_status = tk.StringVar(value="")
        ttk.Label(rodape, textvariable=self.var_status, foreground=f"#{COR_PRIMARIA}").pack(side="left")
        self.indicador = ttk.Progressbar(rodape, mode="indeterminate", length=140)

    # ------------------------------------------------------- aba: cabeçalho
    def _criar_aba_cabecalho(self):
        aba = ttk.Frame(self.abas, padding=18)
        self.abas.add(aba, text="Cabeçalho")
        aba.columnconfigure(0, weight=1)
        aba.columnconfigure(1, weight=1)

        ttk.Label(aba, text="Dados do Cabeçalho", style="Secao.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self.vars_cab = {k: tk.StringVar(value=v) for k, v in CABECALHO_VAZIO.items()}

        def campo(rotulo, chave, linha, coluna=0, span=1):
            quadro = ttk.Frame(aba)
            quadro.grid(row=linha, column=coluna, columnspan=span, sticky="ew", padx=(0, 12 if coluna == 0 and span == 1 else 0), pady=(0, 10))
            quadro.columnconfigure(0, weight=1)
            ttk.Label(quadro, text=rotulo).grid(row=0, column=0, sticky="w")
            ttk.Entry(quadro, textvariable=self.vars_cab[chave]).grid(row=1, column=0, sticky="ew")

        campo("Mamógrafo - Fabricante:", "mamografo_fabricante", 1, 0)
        campo("Mamógrafo - Modelo:", "mamografo_modelo", 1, 1)
        campo("CNES:", "cnes", 2, 0, 2)
        campo("QIID:", "qiid", 3, 0, 2)

        ttk.Label(aba, text="Tipo de mamógrafo:").grid(row=4, column=0, columnspan=2, sticky="w")
        linha_tipo = ttk.Frame(aba)
        linha_tipo.grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 10))
        for opcao in OPCOES_TIPO:
            ttk.Radiobutton(linha_tipo, text=opcao, value=opcao, variable=self.vars_cab["tipo_mamografo"]).pack(side="left", padx=(0, 18))

        campo("Instituição:", "instituicao", 6, 0, 2)
        campo("Cidade:", "cidade", 7, 0)
        campo("Estado:", "estado", 7, 1)
        campo("Serviço:", "servico", 8, 0, 2)

    # ------------------------------------------------------------ aba: casos
    def _criar_aba_casos(self):
        aba = ttk.Frame(self.abas, padding=(10, 10, 10, 0))
        self.abas.add(aba, text="Casos")

        atalho = ttk.LabelFrame(aba, text=" Atalho Rápido ", padding=8)
        atalho.pack(fill="x", pady=(0, 8))
        ttk.Label(atalho, text="Se nenhum dos 5 casos tem consideração específica a registrar:").pack(side="left")
        ttk.Button(atalho, text="Marcar os 5 casos como 'Sim' em tudo e salvar",
                   style="Secundario.TButton", command=self.atalho_tudo_sim).pack(side="right")

        self.nb_casos = ttk.Notebook(aba)
        self.nb_casos.pack(fill="both", expand=True)
        for n in range(1, TOTAL_CASOS + 1):
            form = CasoForm(self, self.nb_casos, n)
            self.forms[n] = form
            self.nb_casos.add(form.scroll, text=f"Caso {n}")

    # ----------------------------------------------------- aba: resumo/export
    def _criar_aba_resumo(self):
        aba = ttk.Frame(self.abas, padding=14)
        self.abas.add(aba, text="Resumo e Exportação")

        ttk.Label(aba, text="Resumo dos casos", style="Secao.TLabel").pack(anchor="w")
        ttk.Label(aba, text="Observações que se aplicam a todos os casos:").pack(anchor="w", pady=(4, 0))
        self.txt_gerais = tk.Text(aba, height=4, wrap="word", font=self.fonte_base, relief="solid", borderwidth=1)
        self.txt_gerais.pack(fill="x")

        botoes = ttk.Frame(aba)
        botoes.pack(fill="x", pady=10)
        ttk.Button(botoes, text="Gerar Relatório Geral", style="Primario.TButton",
                   command=self.gerar_relatorio_geral).pack(side="left")
        ttk.Button(botoes, text="Gerar Documento Word (.docx)", style="Primario.TButton",
                   command=self.exportar_docx).pack(side="left", padx=8)
        ttk.Button(botoes, text="Limpar todos os dados", style="Secundario.TButton",
                   command=self.limpar_tudo).pack(side="right")

        ttk.Label(aba, text="Tabela de Respostas", style="Secao.TLabel").pack(anchor="w", pady=(4, 2))
        self.tabela = ttk.Treeview(aba, show="headings", height=6)
        self.tabela.pack(fill="x")

        ttk.Label(aba, text="Histórico da Sessão", style="Secao.TLabel").pack(anchor="w", pady=(10, 2))
        quadro_texto = ttk.Frame(aba)
        quadro_texto.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(quadro_texto, orient="vertical")
        self.txt_historico = tk.Text(quadro_texto, wrap="word", font=self.fonte_base, relief="solid",
                                     borderwidth=1, state="disabled", yscrollcommand=sb.set)
        sb.configure(command=self.txt_historico.yview)
        sb.pack(side="right", fill="y")
        self.txt_historico.pack(side="left", fill="both", expand=True)
        self.txt_historico.tag_configure("titulo", font=(self.fonte_base.actual("family"), 11, "bold"), foreground=f"#{COR_PRIMARIA}")
        self.txt_historico.tag_configure("rotulo", font=(self.fonte_base.actual("family"), 10, "bold"))

    # ------------------------------------------------------------ utilidades
    def _on_roda_mouse(self, event):
        if getattr(event, "num", None) == 4:
            passo = -1
        elif getattr(event, "num", None) == 5:
            passo = 1
        else:
            passo = -1 if event.delta > 0 else 1
        try:
            w = self.winfo_containing(event.x_root, event.y_root)
        except KeyError:
            return
        while w is not None:
            if isinstance(w, (tk.Text, ttk.Treeview)):
                return
            if isinstance(w, ScrollFrame):
                if w.canvas.yview() != (0.0, 1.0):
                    w.canvas.yview_scroll(passo, "units")
                return
            w = getattr(w, "master", None)

    def sincronizar_cabecalho(self):
        self.estado.dados_cabecalho = {k: v.get() for k, v in self.vars_cab.items()}

    def atualizar_progresso(self):
        feitos = len(self.estado.casos_salvos)
        self.barra["value"] = feitos
        self.lbl_progresso.configure(text=f"Progresso: {feitos} de {TOTAL_CASOS} casos analisados")
        for n in range(1, TOTAL_CASOS + 1):
            marca = " ✓" if f"Caso {n}" in self.estado.casos_salvos else ""
            self.nb_casos.tab(n - 1, text=f"Caso {n}{marca}")

    def mostrar_status(self, texto, limpar_em_ms=None):
        self.var_status.set(texto)
        if limpar_em_ms:
            self.after(limpar_em_ms, lambda: self.var_status.get() == texto and self.var_status.set(""))

    def _rodar_ia(self, pedido, ao_concluir, mensagem):
        """Chama a IA em segundo plano (a janela não trava) e, ao terminar,
        executa ao_concluir(texto) na thread principal."""
        if self._ia_ocupada:
            messagebox.showinfo("Aguarde", "A IA ainda está processando o pedido anterior.")
            return
        self._ia_ocupada = True
        self.var_status.set(mensagem)
        self.indicador.pack(side="right")
        self.indicador.start(12)
        fila = queue.Queue()

        def trabalho():
            try:
                fila.put(("ok", gerar_texto(**pedido)))
            except Exception as e:  # noqa: BLE001
                fila.put(("erro", str(e)))

        threading.Thread(target=trabalho, daemon=True).start()

        def checar():
            try:
                tipo, valor = fila.get_nowait()
            except queue.Empty:
                self.after(150, checar)
                return
            self._ia_ocupada = False
            self.indicador.stop()
            self.indicador.pack_forget()
            self.var_status.set("")
            if tipo == "ok":
                ao_concluir(valor)
            else:
                messagebox.showerror("Erro ao gerar relatório", valor)

        self.after(150, checar)

    # ----------------------------------------------------------------- ações
    def revisar_texto(self, titulo, origem, texto_ia):
        """Abre a caixa de revisão e devolve (acao, texto): acao = salvar | regerar | cancelar."""
        dlg = RevisaoDialog(self, titulo, origem, limpar_formatacao(texto_ia).strip())
        self.wait_window(dlg)
        return dlg.acao, dlg.texto

    def _gravar_caso(self, nome_caso, texto_bruto, consideracoes, id_exame, escolhas, relatorio):
        e = self.estado
        e.casos_salvos[nome_caso] = texto_bruto
        e.consideracoes_caso[nome_caso] = consideracoes
        e.identificacao_exames[nome_caso] = id_exame
        e.escolhas_casos[nome_caso] = escolhas
        e.relatorios_ia[nome_caso] = relatorio
        self.atualizar_progresso()
        self.atualizar_resumo()
        self.mostrar_status(f"{nome_caso} processado com sucesso!", 5000)

    def salvar_caso(self, numero):
        nome_caso = f"Caso {numero}"
        if nome_caso in self.estado.casos_salvos:
            if not messagebox.askyesno("Sobrescrever?", f"O {nome_caso} já foi salvo anteriormente.\n\nDeseja sobrescrever o relatório existente?"):
                return
        form = self.forms[numero]
        consideracoes = form.consideracoes()
        id_exame = form.var_id_exame.get()
        texto_bruto, texto_para_ia, escolhas = montar_caso(form.coletar_respostas(), consideracoes)
        self.estado.dados_adicionais_casos[nome_caso] = form.coletar_adicionais()

        if not texto_para_ia.strip():
            # Nada relevante a enviar: usa o texto padrão, sem chamar a IA.
            self._gravar_caso(nome_caso, texto_bruto, consideracoes, id_exame, escolhas, TEXTO_SEM_CONSIDERACOES)
        elif texto_para_ia.strip() == TEXTO_LAUDO_SEM_CONSIDERACOES:
            # Tudo "Sim" e sem observações: a frase já é fixa, não precisa de IA.
            self._gravar_caso(nome_caso, texto_bruto, consideracoes, id_exame, escolhas, TEXTO_LAUDO_SEM_CONSIDERACOES)
        else:
            pedido = montar_pedido_caso(texto_bruto, consideracoes)
            origem = re.sub(r"\s+", " ", texto_bruto).strip()
            if consideracoes.strip():
                origem += f"\n\nObservações do avaliador: {consideracoes.strip()}"

            def gerar():
                self._rodar_ia(pedido, revisar, "IA está formatando o relatório...")

            def revisar(texto_ia):
                acao, texto_final = self.revisar_texto(f"Revisar texto — {nome_caso}", origem, texto_ia)
                if acao == "salvar":
                    self._gravar_caso(nome_caso, texto_bruto, consideracoes, id_exame, escolhas, texto_final)
                elif acao == "regerar":
                    gerar()
                else:
                    self.mostrar_status(f"{nome_caso} não foi salvo (revisão cancelada).", 5000)

            gerar()

    def atalho_tudo_sim(self):
        nomes = [f"Caso {n}" for n in range(1, TOTAL_CASOS + 1)]
        ja_salvos = [n for n in nomes if n in self.estado.casos_salvos]
        if ja_salvos:
            if not messagebox.askyesno("Sobrescrever?", f"Isto vai sobrescrever os casos já salvos: {', '.join(ja_salvos)}.\n\nConfirma?"):
                return
        titulos = list(perguntas[GRUPO_LAUDO].keys())
        for n, nome in enumerate(nomes, start=1):
            self.forms[n].marcar_tudo_sim()
            escolhas = {t: {"resposta": "Sim", "sub_opcao": []} for t in titulos}
            self.estado.casos_salvos[nome] = TEXTO_LAUDO_SEM_CONSIDERACOES
            self.estado.consideracoes_caso[nome] = ""
            self.estado.identificacao_exames.setdefault(nome, "")
            self.estado.escolhas_casos[nome] = escolhas
            self.estado.relatorios_ia[nome] = TEXTO_LAUDO_SEM_CONSIDERACOES
        self.atualizar_progresso()
        self.atualizar_resumo()
        self.mostrar_status("Os 5 casos foram marcados como 'Sim' em tudo e salvos!", 5000)

    def gerar_relatorio_geral(self):
        if len(self.estado.casos_salvos) < 2:
            messagebox.showinfo("Relatório Geral", "Salve pelo menos 2 casos para gerar o relatório geral.")
            return
        self.estado.consideracoes_gerais = self.txt_gerais.get("1.0", "end").strip()
        pedido = montar_pedido_geral(self.estado)

        def gerar():
            self._rodar_ia(pedido, revisar, "IA está gerando o relatório geral...")

        def revisar(texto_ia):
            acao, texto_final = self.revisar_texto("Revisar relatório geral", "", texto_ia)
            if acao == "salvar":
                self.estado.relatorio_geral_salvo = texto_final
                self.atualizar_resumo()
                self.mostrar_status("Relatório geral salvo!", 5000)
            elif acao == "regerar":
                gerar()
            else:
                self.mostrar_status("Relatório geral não foi salvo (revisão cancelada).", 5000)

        gerar()

    def atualizar_resumo(self):
        e = self.estado
        casos = sorted(e.relatorios_ia.keys(), key=extrair_numero)

        # Tabela de respostas
        self.tabela.delete(*self.tabela.get_children())
        colunas = ["pergunta"] + [f"c{extrair_numero(c)}" for c in casos]
        self.tabela.configure(columns=colunas)
        self.tabela.heading("pergunta", text="Pergunta", anchor="w")
        self.tabela.column("pergunta", width=380, anchor="w", stretch=True)
        for c in casos:
            idc = f"c{extrair_numero(c)}"
            self.tabela.heading(idc, text=c)
            self.tabela.column(idc, width=80, anchor="center", stretch=False)
        self.tabela.tag_configure("zebra", background=f"#{COR_ZEBRA}")
        linha = 0
        for questoes in perguntas.values():
            for pergunta in questoes:
                valores = [pergunta]
                for c in casos:
                    item = e.escolhas_casos.get(c, {}).get(pergunta, {})
                    valores.append(item.get("resposta", "-") if isinstance(item, dict) else item)
                self.tabela.insert("", "end", values=valores, tags=("zebra",) if linha % 2 else ())
                linha += 1
        self.tabela.configure(height=max(3, min(linha, 8)))

        # Histórico
        t = self.txt_historico
        t.configure(state="normal")
        t.delete("1.0", "end")
        for c in casos:
            t.insert("end", f"{c}\n", "titulo")
            if e.identificacao_exames.get(c, "").strip():
                t.insert("end", "Identificação do Exame: ", "rotulo")
                t.insert("end", f"{e.identificacao_exames[c]}\n")
            t.insert("end", f"{e.relatorios_ia[c]}\n")
            if e.consideracoes_caso.get(c, "").strip():
                t.insert("end", "Considerações adicionais: ", "rotulo")
                t.insert("end", f"{e.consideracoes_caso[c]}\n")
            t.insert("end", "\n")
        if e.relatorio_geral_salvo:
            t.insert("end", "Relatório Geral\n", "titulo")
            t.insert("end", f"{e.relatorio_geral_salvo}\n")
        t.configure(state="disabled")

    def exportar_docx(self):
        if not self.estado.relatorios_ia:
            messagebox.showwarning("Exportar", "Salve pelo menos um caso antes de gerar o documento.")
            return
        self.sincronizar_cabecalho()
        self.estado.consideracoes_gerais = self.txt_gerais.get("1.0", "end").strip()
        nome_servico = sanitizar_nome_arquivo(self.estado.dados_cabecalho.get("servico", ""))
        caminho = filedialog.asksaveasfilename(
            title="Salvar documento Word",
            defaultextension=".docx",
            initialfile=f"relatório_{nome_servico}.docx",
            filetypes=[("Documento Word", "*.docx")],
        )
        if not caminho:
            return
        try:
            conteudo = criar_docx(self.estado)
            with open(caminho, "wb") as f:
                f.write(conteudo)
        except PermissionError:
            messagebox.showerror("Erro", "Não consegui salvar o arquivo. Se ele estiver aberto no Word, feche e tente de novo.")
            return
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Erro ao gerar o documento", str(e))
            return
        if messagebox.askyesno("Documento gerado", f"Documento salvo em:\n{caminho}\n\nDeseja abri-lo agora?"):
            abrir_arquivo(caminho)

    def limpar_tudo(self):
        if not messagebox.askyesno("Limpar tudo", "Isto apaga todos os casos, textos e dados do cabeçalho desta sessão.\n\nConfirma?"):
            return
        self.estado.resetar()
        for chave, var in self.vars_cab.items():
            var.set(CABECALHO_VAZIO[chave])
        for form in self.forms.values():
            form.limpar()
        self.txt_gerais.delete("1.0", "end")
        self.atualizar_progresso()
        self.atualizar_resumo()
        self.mostrar_status("Dados da sessão apagados.", 4000)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()