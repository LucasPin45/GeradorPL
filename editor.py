# -*- coding: utf-8 -*-
"""
Operações sobre a lista de dispositivos em edição e sobre o rascunho.

Tudo aqui é função pura (sem Streamlit): recebe e devolve listas de dicts no
formato usado pelo app:

    {"uid": "a1b2c3d4", "tipo": "artigo", "texto": "...", "titulo_estrutura": "",
     "paragrafo_unico_forcado": None, "pontuacao_automatica": True}

Um dispositivo é movido ou excluído junto com os que dependem dele: o Artigo
leva seus parágrafos, incisos e o texto de lei alterada; o Inciso leva suas
alíneas. Movê-los um a um deixava parágrafos "pendurados" no artigo errado.
"""
import json
import uuid
from typing import Any, Dict, List, Optional, Tuple

from gerador import PROFUNDIDADE, TIPOS_CONTEUDO, TIPOS_ESTRUTURA, TIPO_CITACAO

TIPOS_VALIDOS = TIPOS_ESTRUTURA + TIPOS_CONTEUDO + [TIPO_CITACAO]
VERSAO_RASCUNHO = 1

Disp = Dict[str, Any]


def novo_uid() -> str:
    return uuid.uuid4().hex[:8]


def novo_dispositivo(tipo: str, texto: str = "", titulo_estrutura: str = "",
                     paragrafo_unico_forcado: Optional[bool] = None,
                     pontuacao_automatica: bool = True) -> Disp:
    return {"uid": novo_uid(), "tipo": tipo, "texto": texto, "titulo_estrutura": titulo_estrutura,
            "paragrafo_unico_forcado": paragrafo_unico_forcado, "pontuacao_automatica": pontuacao_automatica}


# ---------------------------------------------------------------------------
# Blocos (um dispositivo + os que dependem dele)
# ---------------------------------------------------------------------------

def _nivel(d: Disp) -> int:
    """Títulos, capítulos e seções ficam no mesmo nível dos artigos para
    efeito de mover: sobem e descem entre blocos de artigos."""
    return 1 if d["tipo"] in TIPOS_ESTRUTURA else PROFUNDIDADE[d["tipo"]]


def fim_do_bloco(lista: List[Disp], i: int) -> int:
    """Índice seguinte ao último dependente de lista[i]."""
    if lista[i]["tipo"] in TIPOS_ESTRUTURA:
        return i + 1                      # o cabeçalho anda sozinho, sem arrastar o capítulo inteiro
    nivel = _nivel(lista[i])
    j = i + 1
    while j < len(lista) and lista[j]["tipo"] not in TIPOS_ESTRUTURA and _nivel(lista[j]) > nivel:
        j += 1
    return j


def dependentes(lista: List[Disp], i: int) -> int:
    """Quantos dispositivos acompanham lista[i] ao mover ou excluir."""
    return fim_do_bloco(lista, i) - i - 1


def _inicio_do_bloco_anterior(lista: List[Disp], i: int) -> Optional[int]:
    """Início do bloco irmão imediatamente acima, ou None se não houver
    (primeiro da lista, ou o de cima é o dispositivo ao qual este pertence)."""
    if i == 0:
        return None
    nivel = _nivel(lista[i])
    j = i - 1
    while j > 0 and lista[j]["tipo"] not in TIPOS_ESTRUTURA and _nivel(lista[j]) > nivel:
        j -= 1
    if _nivel(lista[j]) < nivel:
        return None                       # acima está o "pai": não se sobe para fora dele
    if _nivel(lista[j]) > nivel:
        return None                       # só há subordinados de outro dispositivo acima
    return j


def pode_subir(lista: List[Disp], i: int) -> bool:
    return _inicio_do_bloco_anterior(lista, i) is not None


def pode_descer(lista: List[Disp], i: int) -> bool:
    fim = fim_do_bloco(lista, i)
    return fim < len(lista) and _nivel(lista[fim]) == _nivel(lista[i])


def subir(lista: List[Disp], i: int) -> List[Disp]:
    """Troca o bloco de lista[i] com o bloco irmão de cima."""
    j = _inicio_do_bloco_anterior(lista, i)
    if j is None:
        return list(lista)
    fim = fim_do_bloco(lista, i)
    return lista[:j] + lista[i:fim] + lista[j:i] + lista[fim:]


def descer(lista: List[Disp], i: int) -> List[Disp]:
    """Troca o bloco de lista[i] com o bloco irmão de baixo."""
    if not pode_descer(lista, i):
        return list(lista)
    fim = fim_do_bloco(lista, i)
    fim_seguinte = fim_do_bloco(lista, fim)
    return lista[:i] + lista[fim:fim_seguinte] + lista[i:fim] + lista[fim_seguinte:]


def excluir(lista: List[Disp], i: int) -> List[Disp]:
    """Remove lista[i] e os dispositivos que dependem dele."""
    return lista[:i] + lista[fim_do_bloco(lista, i):]


def inserir(lista: List[Disp], novo: Disp, depois_de_uid: Optional[str] = None) -> List[Disp]:
    """Insere no fim ou logo depois do bloco do dispositivo indicado (depois
    dos parágrafos e incisos dele, para não separá-los)."""
    if depois_de_uid:
        for i, d in enumerate(lista):
            if d["uid"] == depois_de_uid:
                fim = fim_do_bloco(lista, i)
                if _nivel(novo) > _nivel(d):
                    # subordinado (ex.: Inciso depois de um Artigo): entra depois dos
                    # subordinados do mesmo tipo que já existem, mas antes dos de nível
                    # acima — um novo inciso do caput fica antes dos parágrafos
                    pos = i + 1
                    while pos < fim and _nivel(lista[pos]) >= _nivel(novo):
                        pos += 1
                else:
                    pos = fim     # mesmo nível ou acima: depois do bloco inteiro
                return lista[:pos] + [novo] + lista[pos:]
    return lista + [novo]


def proximo_tipo_sugerido(lista: List[Disp], depois_de_uid: Optional[str] = None) -> str:
    """Tipo mais provável para o próximo dispositivo, a partir do anterior."""
    anterior = None
    if depois_de_uid:
        anterior = next((d for d in lista if d["uid"] == depois_de_uid), None)
    elif lista:
        anterior = lista[-1]
    if anterior is None:
        return "artigo"
    return {
        "titulo": "capitulo", "capitulo": "artigo", "secao": "artigo", "subsecao": "artigo",
        "artigo": "artigo", "paragrafo": "paragrafo", "inciso": "inciso",
        "alinea": "alinea", "item": "item", TIPO_CITACAO: "artigo",
    }.get(anterior["tipo"], "artigo")


# ---------------------------------------------------------------------------
# Rascunho (salvar o trabalho num arquivo e continuar depois)
# ---------------------------------------------------------------------------

CAMPOS_TEXTO = ("ano", "genero", "autor_nome", "autor_partido_uf", "conector", "tipo_documento",
                "ementa", "justificativa", "fecho_pl",
                "artigo_ricd", "sigla_materia", "sigla_outra", "numero_materia", "ementa_materia", "fecho_req")


def exportar_rascunho(campos: Dict[str, Any], dispositivos: List[Disp]) -> str:
    """JSON com tudo o que foi digitado (campos de texto + dispositivos)."""
    return json.dumps({
        "formato": "gerador-pl-rascunho",
        "versao": VERSAO_RASCUNHO,
        "campos": {k: str(campos.get(k) or "") for k in CAMPOS_TEXTO},
        "dispositivos": dispositivos,
    }, ensure_ascii=False, indent=2)


def importar_rascunho(conteudo: str) -> Tuple[Dict[str, str], List[Disp]]:
    """
    Lê um rascunho salvo. Aceita só o que reconhece: campos de texto conhecidos
    e dispositivos de tipo válido; o resto é ignorado.

    Raises:
        ValueError: o arquivo não é um rascunho deste sistema
    """
    try:
        dados = json.loads(conteudo)
    except (TypeError, ValueError) as e:
        raise ValueError("o arquivo não é um rascunho válido") from e
    if not isinstance(dados, dict) or dados.get("formato") != "gerador-pl-rascunho":
        raise ValueError("o arquivo não é um rascunho deste sistema")

    brutos = dados.get("campos") if isinstance(dados.get("campos"), dict) else {}
    campos = {k: str(brutos[k]) for k in CAMPOS_TEXTO if isinstance(brutos.get(k), (str, int, float))}

    dispositivos: List[Disp] = []
    vistos = set()
    for d in dados.get("dispositivos") or []:
        if not isinstance(d, dict) or d.get("tipo") not in TIPOS_VALIDOS:
            continue
        uid = str(d.get("uid") or "")
        if not uid or uid in vistos:
            uid = novo_uid()
        vistos.add(uid)
        unico = d.get("paragrafo_unico_forcado")
        dispositivos.append({
            "uid": uid, "tipo": d["tipo"], "texto": str(d.get("texto") or ""),
            "titulo_estrutura": str(d.get("titulo_estrutura") or ""),
            "paragrafo_unico_forcado": unico if isinstance(unico, bool) else None,
            "pontuacao_automatica": d.get("pontuacao_automatica") is not False,
        })
    return campos, dispositivos
