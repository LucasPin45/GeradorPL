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
import copy
from dataclasses import dataclass, field
from typing import List, Optional

TEMPLATE_PATH = "assets/template_gabinete.docx"

# ---------------------------------------------------------------------------
# Tipos de dispositivo e hierarquia (art. 10, V, LC 95/98)
# ---------------------------------------------------------------------------

TIPOS_ESTRUTURA = ["titulo", "capitulo", "secao", "subsecao"]
TIPOS_CONTEUDO = ["artigo", "paragrafo", "inciso", "alinea", "item"]

# "profundidade" relativa usada para decidir se um dispositivo abre uma
# sublista (e portanto deve terminar em ":") em relação ao próximo item.
PROFUNDIDADE = {
    "titulo": 0, "capitulo": 0, "secao": 0, "subsecao": 0,
    "artigo": 1, "paragrafo": 2, "inciso": 3, "alinea": 4, "item": 5,
}

ROTULOS = {
    "titulo": "Título", "capitulo": "Capítulo", "secao": "Seção",
    "subsecao": "Subseção", "artigo": "Artigo", "paragrafo": "Parágrafo",
    "inciso": "Inciso", "alinea": "Alínea", "item": "Item",
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
        texto = d.texto.rstrip()
        if d.pontuacao_automatica and texto:
            prof_atual = PROFUNDIDADE[tipo]
            prox = dispositivos[idx + 1] if idx + 1 < n else None
            # ":" só cabe quando o próximo dispositivo é uma enumeração
            # (inciso/alínea/item) mais profunda — um parágrafo após um
            # artigo NÃO é enumeração, então o caput termina em "."
            abre_sublista = (
                prox is not None
                and prox.tipo in ("inciso", "alinea", "item")
                and PROFUNDIDADE[prox.tipo] > prof_atual
            )

            # já termina com pontuação explícita? respeita se usuário colocou
            ja_pontuado = bool(re.search(r'[.;:,]$', texto)) or texto.endswith(("º", "”", '"'))

            if not ja_pontuado:
                if abre_sublista:
                    texto += ":"
                elif modo == "enum":
                    # é item de enumeração (inciso/alínea/item): decide ; / ; e / .
                    eh_ultimo_da_serie = (
                        prox is None or PROFUNDIDADE[prox.tipo] != prof_atual
                    )
                    eh_penultimo_da_serie = False
                    if not eh_ultimo_da_serie:
                        prox2 = dispositivos[idx + 2] if idx + 2 < n else None
                        eh_penultimo_da_serie = (
                            prox2 is None or PROFUNDIDADE[prox2.tipo] != prof_atual
                        )
                    if eh_ultimo_da_serie:
                        texto += "."
                    elif eh_penultimo_da_serie and conector != "nenhum":
                        texto += _conector_sufixo(conector)
                    else:
                        texto += ";"
                else:
                    # artigo, parágrafo (sem sublista): frase normal -> ponto final
                    texto += "."

        resultado.append({
            "tipo": tipo,
            "modo": modo,
            "prefixo": prefixo,
            "texto": texto,
            "titulo_estrutura": d.titulo_estrutura.strip().upper() if tipo in ("titulo", "capitulo") else d.titulo_estrutura.strip(),
        })

    return resultado


# ---------------------------------------------------------------------------
# Construção de XML (OOXML) — reaproveita estilos do template do gabinete
# ---------------------------------------------------------------------------

def esc(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


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

def montar_paragrafos_corpo(dados: dict) -> str:
    xml = []

    ano = dados.get("ano", "2026")
    xml.append(p_center_bold(f"PROJETO DE LEI Nº _____, DE {ano}"))
    xml.append(p_center(f"({dados['autor_prefixo']} {dados['autor_nome']})"))
    xml.append(p_blank())
    xml.append(p_blank())
    xml.append(p_blank())
    xml.append(p_ementa(dados["ementa"]))
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
        else:
            xml.append(p_conteudo(item["prefixo"], item["texto"]))

    # bloco final de vigência já vem como dispositivos (Art. N); não é
    # necessário tratamento especial.

    xml.append(p_blank())
    xml.append(p_justificacao_header())

    justificativa = dados.get("justificativa", "").strip()
    if justificativa:
        paragrafos = [p.strip() for p in re.split(r'\n\s*\n', justificativa) if p.strip()]
        for par in paragrafos:
            par_limpo = re.sub(r'\s*\n\s*', ' ', par)
            xml.append(p_justificacao_paragrafo(par_limpo))

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


def gerar_docx(dados: dict, template_path: str = TEMPLATE_PATH) -> bytes:
    with open(template_path, "rb") as f:
        template_bytes = f.read()

    with zipfile.ZipFile(io.BytesIO(template_bytes), "r") as zin:
        document_xml = zin.read("word/document.xml").decode("utf-8")
        prefixo, sect_xml, sufixo = _extrair_prefixo_sufixo(document_xml)

        corpo = montar_paragrafos_corpo(dados)
        novo_document_xml = prefixo + corpo + sect_xml + sufixo

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = novo_document_xml.encode("utf-8")
                zout.writestr(item, data)

    return buffer.getvalue()
