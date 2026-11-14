# -*- coding: utf-8 -*-
"""
Gerador de Projeto de Lei (PL) no padrão do Gabinete da Deputada Júlia Zanatta,
observando a técnica legislativa da Lei Complementar nº 95/1998 (LC 95/98).

Este módulo:
1. Recebe uma lista estruturada de "dispositivos" (Título, Capítulo, Seção,
   Subseção, Artigo, Parágrafo, Inciso, Alínea, Item) digitados pelo usuário.
2. Calcula automaticamente a numeração de cada dispositivo, seguindo o art. 10
   da LC 95/98 (numeração ordinal até o nono, cardinal a partir do décimo;
   incisos em algarismos romanos; alíneas em letras minúsculas; itens em
   algarismos arábicos).
3. Aplica pontuação automática (";", "; e", "; ou", ":", ".") de acordo com a
   posição do dispositivo dentro da enumeração.
4. Gera o XML (word/document.xml) reaproveitando integralmente os estilos,
   cabeçalho, rodapé e sectPr do modelo .docx do gabinete, e reempacota tudo
   num novo arquivo .docx válido.
"""

import re
import zipfile
import io
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

# Caminho relativo a este arquivo: funciona de qualquer pasta em que o app
# seja iniciado (antes dependia de rodar a partir da raiz do repositório).
TEMPLATE_PATH = str(Path(__file__).resolve().parent / "assets" / "template_gabinete.docx")

# ---------------------------------------------------------------------------
# Tipos de dispositivo e hierarquia (art. 10, V, LC 95/98)
# ---------------------------------------------------------------------------

TIPOS_ESTRUTURA = ["titulo", "capitulo", "secao", "subsecao"]
TIPOS_CONTEUDO = ["artigo", "paragrafo", "inciso", "alinea", "item"]
# Texto de lei alterada: o bloco entre aspas que vem depois de "passa a
# vigorar com a seguinte redação:". Não é numerado nem pontuado pelo sistema
# e não entra na contagem dos artigos do projeto.
TIPO_CITACAO = "citacao"
TIPOS_ENUMERACAO = ("inciso", "alinea", "item")

# "profundidade" relativa usada para decidir se um dispositivo abre uma
# sublista (e portanto deve terminar em ":") em relação ao próximo item.
PROFUNDIDADE = {
    "titulo": 0, "capitulo": 0, "secao": 0, "subsecao": 0,
    "artigo": 1, "paragrafo": 2, "inciso": 3, "alinea": 4, "item": 5,
    "citacao": 6,
}

ROTULOS = {
    "titulo": "Título", "capitulo": "Capítulo", "secao": "Seção",
    "subsecao": "Subseção", "artigo": "Artigo", "paragrafo": "Parágrafo",
    "inciso": "Inciso", "alinea": "Alínea", "item": "Item",
    "citacao": "Texto da lei alterada",
}


@dataclass
class Dispositivo:
    tipo: str                      # um de TIPOS_ESTRUTURA + TIPOS_CONTEUDO
    texto: str                     # conteúdo digitado pelo usuário
    titulo_estrutura: str = ""     # para titulo/capitulo/secao/subsecao: título do bloco (ex.: "DISPOSIÇÕES GERAIS")
    paragrafo_unico_forcado: Optional[bool] = None  # None = automático
    pontuacao_automatica: bool = True
    uid: str = ""


# ---------------------------------------------------------------------------
# Conversão para algarismos romanos
# ---------------------------------------------------------------------------

_ROMAN_MAP = [
    (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
    (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
    (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
]


def to_roman(n: int) -> str:
    if n <= 0:
        return str(n)
    result = ""
    for value, symbol in _ROMAN_MAP:
        while n >= value:
            result += symbol
            n -= value
    return result


def to_letra(n: int) -> str:
    """1 -> a, 2 -> b, ..., 26 -> z, 27 -> aa, ..."""
    letras = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letras = chr(97 + rem) + letras
    return letras


def numero_ordinal_cardinal(n: int, marcador: str) -> str:
    """Regra do art. 10, I e III, da LC 95/98: ordinal até o nono, cardinal
    a partir deste. `marcador` é 'º' para artigos/parágrafos."""
    if n <= 9:
        return f"{n}{marcador}"
    return f"{n}."


# ---------------------------------------------------------------------------
# Cálculo de numeração e pontuação
# ---------------------------------------------------------------------------

def _conector_sufixo(conector: str) -> str:
    return {"e": "; e", "ou": "; ou", "nenhum": ";"}.get(conector, "; e")


# Caracteres que o XML 1.0 não aceita (aparecem ao colar texto do Word ou de
# PDF) — um único deles deixava o .docx impossível de abrir.
_XML_INVALIDO = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")


def limpar_texto(texto: str, manter_quebras: bool = False) -> str:
    """Remove caracteres de controle e normaliza espaços. Quebras de linha
    viram espaço, a não ser em blocos de texto citado (manter_quebras)."""
    t = _XML_INVALIDO.sub(" ", (texto or "").replace("\r\n", "\n").replace("\r", "\n"))
    if manter_quebras:
        linhas = [re.sub(r"[ \t\u00a0]+", " ", l).strip() for l in t.split("\n")]
        return "\n".join(l for l in linhas if l)
    return re.sub(r"\s+", " ", t).strip()


def _ja_pontuado(texto: str) -> bool:
    """True se o usuário já encerrou o texto e o sistema não deve acrescentar
    nada: pontuação final, "; e" / "; ou" digitados, aspas de fechamento ou a
    marca "(NR)" de dispositivo com nova redação."""
    t = texto.rstrip()
    if re.search(r"[.;:,?!]$", t):
        return True
    if re.search(r"[;,]\s*(e|ou)$", t, re.IGNORECASE):
        return True
    if re.search(r"\(NR\)$", t, re.IGNORECASE):
        return True
    return t.endswith(("”", '"'))


def _abre_sublista(dispositivos: List["Dispositivo"], idx: int) -> bool:
    """O próximo dispositivo é um item de enumeração mais profundo: o texto
    atual anuncia uma lista e termina em dois-pontos."""
    if idx + 1 >= len(dispositivos):
        return False
    prox = dispositivos[idx + 1]
    return prox.tipo in TIPOS_ENUMERACAO and PROFUNDIDADE[prox.tipo] > PROFUNDIDADE[dispositivos[idx].tipo]


def _proximo_irmao(dispositivos: List["Dispositivo"], idx: int) -> Optional[int]:
    """Próximo item da mesma enumeração (mesmo nível, mesmo "pai"), pulando
    os subitens que estiverem no meio. None se a enumeração termina aqui."""
    prof = PROFUNDIDADE[dispositivos[idx].tipo]
    for j in range(idx + 1, len(dispositivos)):
        t = dispositivos[j].tipo
        if t in TIPOS_ESTRUTURA or t == TIPO_CITACAO or PROFUNDIDADE[t] < prof:
            return None
        if PROFUNDIDADE[t] == prof:
            return j
    return None


def _pai(dispositivos: List["Dispositivo"], idx: int) -> Optional[int]:
    """Dispositivo ao qual este está subordinado (o anterior mais raso)."""
    prof = PROFUNDIDADE[dispositivos[idx].tipo]
    for j in range(idx - 1, -1, -1):
        t = dispositivos[j].tipo
        if t in TIPOS_ESTRUTURA or t == TIPO_CITACAO:
            return None
        if PROFUNDIDADE[t] < prof:
            return j
    return None


def _fim_de_enumeracao(dispositivos: List["Dispositivo"], idx: int, conector: str) -> str:
    """
    Pontuação final de um inciso, alínea ou item que não abre sublista:
      - ";" no meio da enumeração;
      - "; e" / "; ou" no penúltimo;
      - no último, depende de onde a enumeração está: "." se ela encerra o
        artigo ou parágrafo; se está dentro de outra enumeração que continua,
        herda a pontuação que o item de cima teria. Assim, a última alínea de
        um inciso que não é o último termina em ";", e não em ponto final:
            I – os servidores:
            a) ativos; e
            b) inativos;
            II – os militares.
    """
    irmao = _proximo_irmao(dispositivos, idx)
    if irmao is None:
        pai = _pai(dispositivos, idx)
        if pai is not None and dispositivos[pai].tipo in TIPOS_ENUMERACAO:
            return _fim_de_enumeracao(dispositivos, pai, conector)
        return "."
    if _proximo_irmao(dispositivos, irmao) is None and conector != "nenhum":
        return _conector_sufixo(conector)
    return ";"


def processar_dispositivos(dispositivos: List[Dispositivo], conector: str = "e"):
    """Retorna uma lista de dicts prontos para renderização, já com o
    prefixo numérico calculado e o texto final (com pontuação)."""

    contadores = {
        "titulo": 0, "capitulo": 0, "secao": 0, "subsecao": 0,
        "artigo": 0, "paragrafo": 0, "inciso": 0, "alinea": 0, "item": 0,
    }

    # Para saber se um parágrafo é "único", precisamos olhar à frente:
    # contamos quantos parágrafos existem consecutivamente sob o mesmo artigo.
    resultado = []
    n = len(dispositivos)

    for idx, d in enumerate(dispositivos):
        tipo = d.tipo

        # --- resets hierárquicos ---
        if tipo == "titulo":
            contadores["titulo"] += 1
            contadores["capitulo"] = 0
        elif tipo == "capitulo":
            contadores["capitulo"] += 1
            contadores["secao"] = 0
        elif tipo == "secao":
            contadores["secao"] += 1
            contadores["subsecao"] = 0
        elif tipo == "subsecao":
            contadores["subsecao"] += 1
        elif tipo == "artigo":
            contadores["artigo"] += 1
            contadores["paragrafo"] = 0
            contadores["inciso"] = 0
            contadores["alinea"] = 0
            contadores["item"] = 0
        elif tipo == "paragrafo":
            contadores["paragrafo"] += 1
            contadores["inciso"] = 0
            contadores["alinea"] = 0
            contadores["item"] = 0
        elif tipo == "inciso":
            contadores["inciso"] += 1
            contadores["alinea"] = 0
            contadores["item"] = 0
        elif tipo == "alinea":
            contadores["alinea"] += 1
            contadores["item"] = 0
        elif tipo == "item":
            contadores["item"] += 1

        # --- prefixo numérico / rótulo ---
        prefixo = ""
        modo = "estrutura"  # 'estrutura' | 'artigo' | 'paragrafo' | 'enum'

        if tipo == "titulo":
            prefixo = f"TÍTULO {to_roman(contadores['titulo'])}"
        elif tipo == "capitulo":
            prefixo = f"CAPÍTULO {to_roman(contadores['capitulo'])}"
        elif tipo == "secao":
            prefixo = f"Seção {to_roman(contadores['secao'])}"
        elif tipo == "subsecao":
            prefixo = f"Subseção {to_roman(contadores['subsecao'])}"
        elif tipo == "artigo":
            prefixo = f"Art. {numero_ordinal_cardinal(contadores['artigo'], 'º')}"
            modo = "artigo"
        elif tipo == "paragrafo":
            # é único? olhar se o próximo dispositivo de mesmo nível dentro
            # do mesmo artigo também é parágrafo
            eh_unico = d.paragrafo_unico_forcado
            if eh_unico is None:
                # conta quantos parágrafos existem no total sob este artigo
                total_paragrafos = 1
                for j in range(idx + 1, n):
                    if dispositivos[j].tipo == "artigo":
                        break
                    if dispositivos[j].tipo == "paragrafo":
                        total_paragrafos += 1
                # também precisamos saber se já vieram parágrafos antes
                # (contadores['paragrafo'] já foi incrementado acima)
                eh_unico = (contadores["paragrafo"] == 1 and total_paragrafos == 1)
            if eh_unico:
                prefixo = "Parágrafo único."
            else:
                prefixo = f"§ {numero_ordinal_cardinal(contadores['paragrafo'], 'º')}"
            modo = "paragrafo"
        elif tipo == "inciso":
            prefixo = f"{to_roman(contadores['inciso'])} –"
            modo = "enum"
        elif tipo == "alinea":
            prefixo = f"{to_letra(contadores['alinea'])})"
            modo = "enum"
        elif tipo == "item":
            prefixo = f"{contadores['item']}."
            modo = "enum"

        # --- pontuação automática ---
        if tipo == TIPO_CITACAO:
            modo = "citacao"
            texto = limpar_texto(d.texto, manter_quebras=True)
        else:
            texto = limpar_texto(d.texto)
            if d.pontuacao_automatica and texto and not _ja_pontuado(texto):
                prox = dispositivos[idx + 1] if idx + 1 < n else None
                if prox is not None and prox.tipo == TIPO_CITACAO:
                    texto += ":"      # "…passa a vigorar com a seguinte redação:"
                elif _abre_sublista(dispositivos, idx):
                    texto += ":"
                elif modo == "enum":
                    texto += _fim_de_enumeracao(dispositivos, idx, conector)
                else:
                    # artigo, parágrafo (sem sublista): frase normal -> ponto final
                    texto += "."

        resultado.append({
            "tipo": tipo,
            "modo": modo,
            "prefixo": prefixo,
            "texto": texto,
            "titulo_estrutura": (limpar_texto(d.titulo_estrutura).upper() if tipo in ("titulo", "capitulo")
                                 else limpar_texto(d.titulo_estrutura)),
        })

    return resultado


# ---------------------------------------------------------------------------
# Verificação da estrutura (avisos — não impedem gerar o documento)
# ---------------------------------------------------------------------------

_PAIS_VALIDOS = {
    "paragrafo": ("artigo",),
    "inciso": ("artigo", "paragrafo"),
    "alinea": ("inciso",),
    "item": ("alinea",),
}
_NOMES = {"artigo": "um Artigo", "paragrafo": "um Parágrafo", "inciso": "um Inciso", "alinea": "uma Alínea"}


def verificar_estrutura(dispositivos: List[Dispositivo], conector: str = "e") -> List[Dict[str, str]]:
    """
    Aponta problemas de técnica legislativa na sequência de dispositivos.
    Devolve [{"uid": ..., "rotulo": "Art. 2º", "aviso": "..."}], na ordem do
    texto; "uid" vazio quando o aviso é sobre o projeto como um todo.
    """
    avisos: List[Dict[str, str]] = []
    if not dispositivos:
        return avisos
    processados = processar_dispositivos(dispositivos, conector)

    def _rotulo(i: int) -> str:
        return processados[i]["prefixo"].rstrip(" –.") or ROTULOS[dispositivos[i].tipo]

    def _aviso(i: int, texto: str) -> None:
        avisos.append({"uid": dispositivos[i].uid, "rotulo": _rotulo(i), "aviso": texto})

    for i, d in enumerate(dispositivos):
        conteudo = d.titulo_estrutura if d.tipo in TIPOS_ESTRUTURA else d.texto
        if not limpar_texto(conteudo):
            _aviso(i, "está sem texto.")

        if d.tipo in _PAIS_VALIDOS:
            pai = _pai(dispositivos, i)
            aceitos = _PAIS_VALIDOS[d.tipo]
            if pai is None or dispositivos[pai].tipo not in aceitos:
                esperado = " ou ".join(_NOMES[t] for t in aceitos)
                _aviso(i, f"precisa vir depois de {esperado} — do jeito que está, não pertence a nenhum.")

        if d.tipo == TIPO_CITACAO and (i == 0 or dispositivos[i - 1].tipo in TIPOS_ESTRUTURA):
            _aviso(i, "o texto da lei alterada precisa vir depois do Artigo que anuncia a alteração.")

        if d.tipo in TIPOS_ENUMERACAO:
            anterior_e_irmao = any(_proximo_irmao(dispositivos, j) == i for j in range(i))
            if not anterior_e_irmao and _proximo_irmao(dispositivos, i) is None:
                _aviso(i, "é o único item da lista. Uma enumeração com um item só costuma caber no próprio "
                          "texto do dispositivo de cima.")

        if d.tipo in ("artigo", "paragrafo") + TIPOS_ENUMERACAO:
            texto = limpar_texto(d.texto)
            seguinte = dispositivos[i + 1] if i + 1 < len(dispositivos) else None
            anuncia = texto.endswith(":")
            tem_lista = seguinte is not None and (seguinte.tipo == TIPO_CITACAO or _abre_sublista(dispositivos, i))
            if anuncia and not tem_lista:
                _aviso(i, "termina em dois-pontos, mas não há lista nem texto citado logo depois.")

    artigos = [d for d in dispositivos if d.tipo == "artigo"]
    if not artigos:
        avisos.append({"uid": "", "rotulo": "Projeto", "aviso": "ainda não tem nenhum Artigo."})
    elif not any(re.search(r"entra(m|rá|rão)? em vigor", limpar_texto(d.texto), re.IGNORECASE) for d in artigos):
        avisos.append({"uid": "", "rotulo": "Projeto",
                       "aviso": "não tem cláusula de vigência (ex.: \"Esta Lei entra em vigor na data de sua "
                                "publicação\"), que costuma ser o último Artigo."})
    return avisos


# ---------------------------------------------------------------------------
# Construção de XML (OOXML) — reaproveita estilos do template do gabinete
# ---------------------------------------------------------------------------

def esc(s: str) -> str:
    s = _XML_INVALIDO.sub(" ", s or "")
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_runs(text: str, force_bold: bool = False) -> str:
    """Converte **negrito** em runs <w:r> com <w:b/>."""
    if not text:
        return ""
    parts = re.split(r'(\*\*.*?\*\*)', text)
    runs = ""
    for part in parts:
        if not part:
            continue
        bold = force_bold
        t = part
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            bold = True
            t = part[2:-2]
        if t == "":
            continue
        rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
        runs += f'<w:r>{rpr}<w:t xml:space="preserve">{esc(t)}</w:t></w:r>'
    return runs


def p_blank() -> str:
    return '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:p>'


def p_center_bold(text: str) -> str:
    return (
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    )


def p_center(text: str) -> str:
    return (
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    )


def p_ementa(text: str) -> str:
    return (
        '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/>'
        '<w:ind w:left="4252"/><w:jc w:val="both"/></w:pPr>'
        f'{build_runs(text)}</w:p>'
    )


def p_decreta(text: str) -> str:
    return (
        '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/>'
        '<w:ind w:firstLine="709"/><w:jc w:val="both"/></w:pPr>'
        f'{build_runs(text)}</w:p>'
    )


def p_heading(text: str) -> str:
    """Título/Capítulo/Seção/Subseção: centralizado, negrito, Arial 12."""
    return (
        '<w:p><w:pPr><w:pStyle w:val="SemEspaamento"/>'
        '<w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="center"/>'
        '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="24"/></w:rPr>'
        '</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/>'
        f'<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'
    )


def p_conteudo(prefixo: str, texto: str) -> str:
    """Artigo / Parágrafo / Inciso / Alínea / Item: mesmo estilo de parágrafo
    (firstLine=709, linha 1.5, justificado) — igual ao observado no modelo."""
    corpo = f"{prefixo} {texto}" if prefixo else texto
    return (
        '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/>'
        '<w:ind w:firstLine="709"/><w:jc w:val="both"/></w:pPr>'
        f'{build_runs(corpo)}</w:p>'
    )


def p_justificacao_header() -> str:
    return (
        '<w:p><w:pPr><w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>'
        '<w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:b/></w:rPr><w:t>JUSTIFICAÇÃO</w:t></w:r></w:p>'
    )


def p_justificacao_paragrafo(text: str) -> str:
    return (
        '<w:p><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/>'
        '<w:ind w:firstLine="1418"/><w:jc w:val="both"/></w:pPr>'
        f'{build_runs(text)}</w:p>'
    )


# ---------------------------------------------------------------------------
# Montagem completa do documento
# ---------------------------------------------------------------------------

def paragrafos_justificacao(justificativa: str) -> List[str]:
    """Parágrafos da justificação: separados por linha em branco; quebras
    simples dentro de um parágrafo viram espaço. Usado pelo Word e pela
    pré-visualização, para os dois mostrarem a mesma coisa."""
    texto = (justificativa or "").replace("\r\n", "\n").replace("\r", "\n")
    return [limpar_texto(p) for p in re.split(r"\n\s*\n", texto) if p.strip()]


def montar_paragrafos_corpo(dados: dict) -> str:
    xml = []

    ano = dados.get("ano", "2026")
    xml.append(p_center_bold(f"PROJETO DE LEI Nº _____, DE {ano}"))
    xml.append(p_center(f"({dados['autor_prefixo']} {dados['autor_nome']})"))
    xml.append(p_blank())
    xml.append(p_blank())
    xml.append(p_blank())
    xml.append(p_ementa(limpar_texto(dados["ementa"])))
    xml.append(p_blank())
    xml.append(p_blank())
    xml.append(p_decreta("O Congresso Nacional decreta:"))
    xml.append(p_blank())

    processados = processar_dispositivos(dados["dispositivos"], dados.get("conector", "e"))

    for item in processados:
        tipo = item["tipo"]
        if tipo in TIPOS_ESTRUTURA:
            xml.append(p_heading(item["prefixo"]))
            if item["titulo_estrutura"]:
                xml.append(p_heading(item["titulo_estrutura"]))
            xml.append(p_blank())
        elif tipo == TIPO_CITACAO:
            # texto de lei alterada: cada linha digitada vira um parágrafo
            for linha in item["texto"].split("\n"):
                xml.append(p_conteudo("", linha))
        else:
            xml.append(p_conteudo(item["prefixo"], item["texto"]))

    # bloco final de vigência já vem como dispositivos (Art. N); não é
    # necessário tratamento especial.

    xml.append(p_blank())
    xml.append(p_justificacao_header())

    for par in paragrafos_justificacao(dados.get("justificativa", "")):
        xml.append(p_justificacao_paragrafo(par))

    xml.append(p_blank())
    xml.append(p_center(dados.get("local_data", "Sala das Sessões, na data de sua assinatura.")))
    xml.append(p_blank())
    xml.append(p_center_bold(f"{dados['autor_prefixo_assinatura']} {dados['autor_nome'].upper()}"))
    xml.append(p_center(dados["autor_partido_uf"]))

    return "".join(xml)


def _extrair_prefixo_sufixo(document_xml: str):
    """Separa o document.xml do template em (prefixo até <w:body>,
    sectPr final, sufixo após </w:sectPr>)."""
    body_start = document_xml.index("<w:body>") + len("<w:body>")
    prefixo = document_xml[:body_start]

    # último sectPr do documento = sectPr de nível de corpo (seção final)
    sect_matches = list(re.finditer(r'<w:sectPr[ >].*?</w:sectPr>', document_xml, re.DOTALL))
    if not sect_matches:
        raise ValueError("sectPr não encontrado no template")
    last_sect = sect_matches[-1]
    sect_xml = last_sect.group(0)

    sufixo = document_xml[last_sect.end():]  # deve ser só </w:body></w:document>

    return prefixo, sect_xml, sufixo


def _empacotar(corpo_xml: str, template_path: str = TEMPLATE_PATH) -> bytes:
    """Reaproveita estilos/cabeçalho/rodapé/sectPr do template do gabinete,
    substituindo apenas o corpo do documento."""
    with open(template_path, "rb") as f:
        template_bytes = f.read()

    with zipfile.ZipFile(io.BytesIO(template_bytes), "r") as zin:
        document_xml = zin.read("word/document.xml").decode("utf-8")
        prefixo, sect_xml, sufixo = _extrair_prefixo_sufixo(document_xml)

        novo_document_xml = prefixo + corpo_xml + sect_xml + sufixo

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = novo_document_xml.encode("utf-8")
                zout.writestr(item, data)

    return buffer.getvalue()


def gerar_docx(dados: dict, template_path: str = TEMPLATE_PATH) -> bytes:
    """Gera o .docx de um Projeto de Lei."""
    corpo = montar_paragrafos_corpo(dados)
    return _empacotar(corpo, template_path)


# ---------------------------------------------------------------------------
# Módulo: Requerimento de Urgência (art. 154/155 do RICD)
# ---------------------------------------------------------------------------

# Siglas de proposições de gênero feminino ("da PEC", "da Medida Provisória")
_SIGLAS_FEMININAS = {"PEC", "MPV", "MP", "INC", "MSC", "EMC", "EMP"}


def artigo_da_sigla(sigla: str) -> str:
    """ "do" ou "da", conforme a proposição: "do PL", "da PEC", "da MPV"."""
    return "da" if (sigla or "").strip().upper() in _SIGLAS_FEMININAS else "do"


def limpar_ementa_citada(ementa: str) -> str:
    """Ementa colada pelo usuário, pronta para ficar entre aspas: sem as aspas
    e sem o ponto final (o ponto vem depois das aspas de fechamento)."""
    t = limpar_texto(ementa).strip(' "“”\'')
    return t.rstrip(". ").strip()


def ementa_requerimento_urgencia(artigo_ricd: str, sigla_materia: str, numero_materia: str) -> str:
    """Ementa padrão, preenchida automaticamente a partir dos campos do
    formulário — mesmo texto que aparece no bloco de ementa, no estilo PL."""
    numero = limpar_texto(numero_materia) or "____/____"
    sigla = limpar_texto(sigla_materia).upper() or "____"
    return (
        f"Requer, nos termos do art. {artigo_ricd} do Regimento Interno da "
        f"Câmara dos Deputados, urgência para apreciação {artigo_da_sigla(sigla)} {sigla} {numero}."
    )


def texto_requerimento_urgencia(artigo_ricd: str, sigla_materia: str, numero_materia: str,
                                ementa_materia: str) -> str:
    """Corpo do requerimento. Usado pelo Word e pela pré-visualização."""
    numero = limpar_texto(numero_materia) or "____/____"
    sigla = limpar_texto(sigla_materia).upper() or "____"
    ementa = limpar_ementa_citada(ementa_materia) or "(a ementa da matéria aparecerá aqui)"
    return (
        f"Requeremos, nos termos do art. {artigo_ricd} do Regimento Interno da Câmara "
        f"dos Deputados, urgência para apreciação {artigo_da_sigla(sigla)} {sigla} {numero}, que "
        f'"{ementa}".'
    )


def montar_paragrafos_corpo_requerimento(dados: dict) -> str:
    xml = []

    ano = dados.get("ano", "2026")
    artigo = dados.get("artigo_ricd", "155")
    sigla = dados.get("sigla_materia", "PL")
    numero = dados.get("numero_materia", "").strip()
    ementa_materia = dados.get("ementa_materia", "").strip()

    xml.append(p_center_bold(f"REQUERIMENTO Nº _____ DE {ano}"))
    xml.append(p_blank())
    xml.append(p_blank())
    xml.append(p_ementa(ementa_requerimento_urgencia(artigo, sigla, numero)))
    xml.append(p_blank())
    xml.append(p_blank())
    xml.append(p_conteudo("", "Senhor Presidente,"))
    xml.append(p_blank())

    xml.append(p_conteudo("", texto_requerimento_urgencia(artigo, sigla, numero, ementa_materia)))

    xml.append(p_blank())
    xml.append(p_blank())
    xml.append(p_blank())

    xml.append(p_center(dados.get("fecho", "Sala das Sessões, na data de sua assinatura")))
    xml.append(p_blank())
    xml.append(p_center_bold(f"{dados['autor_prefixo_assinatura']} {dados['autor_nome']}"))
    xml.append(p_center(f"({dados['autor_partido_uf']})"))

    return "".join(xml)


def gerar_docx_requerimento(dados: dict, template_path: str = TEMPLATE_PATH) -> bytes:
    """Gera o .docx de um Requerimento de Urgência (art. 154/155 do RICD)."""
    corpo = montar_paragrafos_corpo_requerimento(dados)
    return _empacotar(corpo, template_path)
