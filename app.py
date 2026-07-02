# -*- coding: utf-8 -*-
import uuid
import streamlit as st

from gerador import (
    Dispositivo, TIPOS_ESTRUTURA, TIPOS_CONTEUDO, ROTULOS, PROFUNDIDADE,
    processar_dispositivos, gerar_docx,
)

st.set_page_config(page_title="Gerador de PL — Gabinete Zanatta", page_icon="📜", layout="wide")

# ---------------------------------------------------------------------------
# Conteúdo explicativo em linguagem simples (para quem não é da área jurídica)
# ---------------------------------------------------------------------------
TIPO_INFO = {
    "titulo": {
        "label": "📘 Título",
        "explicacao": "Divisão bem grande do texto, que agrupa vários Capítulos. "
                       "Só é necessária em leis muito extensas — a maioria dos PLs não usa.",
        "placeholder_bloco": "DAS DISPOSIÇÕES GERAIS",
    },
    "capitulo": {
        "label": "📗 Capítulo",
        "explicacao": "Organiza os Artigos por assunto, como os capítulos de um livro. "
                       "Aparece centralizado e em negrito. Ex.: \"DISPOSIÇÕES GERAIS\", \"DAS PENALIDADES\".",
        "placeholder_bloco": "DISPOSIÇÕES GERAIS",
    },
    "secao": {
        "label": "📙 Seção",
        "explicacao": "Divide um Capítulo em partes menores, quando ele trata de mais de um subtema.",
        "placeholder_bloco": "Da Fiscalização",
    },
    "subsecao": {
        "label": "📕 Subseção",
        "explicacao": "Uso raro — só quando uma Seção precisa ser dividida ainda mais.",
        "placeholder_bloco": "Das Penalidades Administrativas",
    },
    "artigo": {
        "label": "📄 Artigo  (uma regra)",
        "explicacao": "É a unidade básica do texto. Cada Artigo deve trazer **uma única** regra, "
                       "proibição ou determinação — como uma frase de comando.",
        "placeholder_texto": "Fica instituído o Programa XYZ, destinado a...",
    },
    "paragrafo": {
        "label": "↳ Parágrafo  (uma exceção ou detalhe)",
        "explicacao": "Complementa, detalha ou traz uma exceção à regra do **Artigo digitado logo acima**. "
                       "Nunca aparece sozinho — sempre depende de um Artigo.",
        "placeholder_texto": "O disposto no caput não se aplica nos casos de...",
    },
    "inciso": {
        "label": "•  Inciso  (item de uma lista)",
        "explicacao": "Use quando o Artigo ou Parágrafo anterior anuncia uma lista (geralmente termina "
                       "com dois-pontos). Cada Inciso é um item dessa lista, numerado em algarismo romano.",
        "placeholder_texto": "os servidores públicos federais ativos",
    },
    "alinea": {
        "label": "◦  Alínea  (sub-item)",
        "explicacao": "Quando um Inciso precisa se dividir em itens ainda mais específicos (letras a, b, c...).",
        "placeholder_texto": "os ocupantes exclusivamente de cargo em comissão",
    },
    "item": {
        "label": "▪  Item  (detalhe da alínea)",
        "explicacao": "Uso raro — um detalhe dentro de uma Alínea, numerado em algarismo (1, 2, 3...).",
        "placeholder_texto": "os nomeados a partir de 2020",
    },
}

ICONE = {
    "titulo": "📘", "capitulo": "📗", "secao": "📙", "subsecao": "📕",
    "artigo": "📄", "paragrafo": "↳", "inciso": "•", "alinea": "◦", "item": "▪",
}

TIPO_OPCOES = TIPOS_ESTRUTURA + TIPOS_CONTEUDO

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
.paper {
    background: white;
    border: 1px solid #e5e7eb;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    padding: 42px 48px;
    font-family: Arial, sans-serif;
    font-size: 14.5px;
    line-height: 1.5;
    color: #111827;
    max-height: 78vh;
    overflow-y: auto;
}
.paper .titulo-pl { text-align:center; font-weight:700; margin-bottom:2px;}
.paper .autor-pl { text-align:center; margin-bottom:22px;}
.paper .ementa { margin-left:44%; text-align:justify; margin-bottom:22px; }
.paper .decreta { text-indent:34px; text-align:justify; margin-bottom:14px;}
.paper .heading { text-align:center; font-weight:700; margin: 4px 0; }
.paper .secao-h { text-align:center; font-weight:700; margin: 10px 0 4px 0; font-size: 14px;}
.paper .conteudo { text-indent:34px; text-align:justify; margin: 6px 0; }
.paper .justif-h { text-align:center; font-weight:700; margin: 24px 0 10px 0; }
.paper .justif-p { text-indent:60px; text-align:justify; margin: 6px 0; }
.paper .fecho { text-align:center; margin-top:22px; }
.paper .assinatura { text-align:center; font-weight:700; margin-top:14px; }
.paper .partido { text-align:center; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Estado
# ---------------------------------------------------------------------------
if "dispositivos" not in st.session_state:
    st.session_state.dispositivos = []
if "editando_uid" not in st.session_state:
    st.session_state.editando_uid = None


def novo_uid():
    return str(uuid.uuid4())[:8]


def dispositivo_para_objeto(d: dict) -> Dispositivo:
    return Dispositivo(
        tipo=d["tipo"], texto=d.get("texto", ""),
        titulo_estrutura=d.get("titulo_estrutura", ""),
        paragrafo_unico_forcado=d.get("paragrafo_unico_forcado"),
        pontuacao_automatica=d.get("pontuacao_automatica", True),
        uid=d.get("uid", ""),
    )


def sugerir_proximo_tipo():
    if not st.session_state.dispositivos:
        return "artigo"
    ultimo = st.session_state.dispositivos[-1]["tipo"]
    return {
        "titulo": "capitulo", "capitulo": "artigo", "secao": "artigo", "subsecao": "artigo",
        "artigo": "artigo", "paragrafo": "paragrafo", "inciso": "inciso",
        "alinea": "alinea", "item": "item",
    }.get(ultimo, "artigo")


# ---------------------------------------------------------------------------
# Sidebar — dados gerais
# ---------------------------------------------------------------------------
st.sidebar.header("① Dados gerais")

ano = st.sidebar.text_input("Ano", value="2026")
genero = st.sidebar.radio("Autoria", ["Deputada", "Deputado"], horizontal=True)
autor_nome = st.sidebar.text_input("Nome do autor(a)", value="Júlia Zanatta")
autor_partido_uf = st.sidebar.text_input("Partido/UF", value="PL/SC")

autor_prefixo = "Da Sra." if genero == "Deputada" else "Do Sr."
autor_prefixo_assinatura = genero

st.sidebar.divider()
conector = st.sidebar.selectbox(
    "Conectivo no penúltimo item de uma lista",
    options=["e", "ou", "nenhum"],
    format_func=lambda x: {"e": '"; e"  (regra geral)', "ou": '"; ou"  (alternativas)', "nenhum": 'apenas ";"'}[x],
    help="Usado automaticamente no penúltimo Inciso/Alínea/Item de uma lista. Ex.: 'I – ...; II – ...; e III – ...'.",
)

st.sidebar.divider()
st.sidebar.caption(
    "O Word gerado usa a mesma fonte, margens, cabeçalho e rodapé do modelo "
    "oficial do gabinete. A numeração segue a Lei Complementar nº 95/1998."
)

# ---------------------------------------------------------------------------
# Corpo principal
# ---------------------------------------------------------------------------
st.title("📜 Gerador de Projeto de Lei")
st.caption("Preencha em linguagem simples — a numeração e a formatação técnica ficam por conta do sistema.")

with st.expander("💡 Guia rápido — o que é cada coisa?"):
    st.markdown("""
Pense assim, sem se preocupar com os nomes jurídicos:

- **Artigo** = uma regra. É a espinha dorsal do texto — cada uma traz um comando ou determinação.
- **Parágrafo** = uma exceção ou um detalhe daquela regra (do Artigo logo acima).
- **Inciso** = um item de uma lista dentro de um Artigo ou Parágrafo.
- **Alínea** = um item dentro de um Inciso (raro).
- **Capítulo / Seção** = apenas organizam os Artigos por tema, como os capítulos de um livro —
  não têm "conteúdo" próprio, só um título.

Você **não digita numeração nem pontuação final** — o sistema calcula "Art. 1º", "§ 2º", "I –", "a)"
e o "." ou ";" no fim de cada frase automaticamente.
""")

col_esq, col_dir = st.columns([1, 1], gap="large")

# ===========================================================================
# COLUNA ESQUERDA — construção do texto
# ===========================================================================
with col_esq:
    st.subheader("② Ementa")
    ementa = st.text_area(
        "O resumo do que a lei faz, em uma frase",
        height=80,
        placeholder="Institui o Programa XYZ e dá outras providências, e altera "
                    "**o art. 5º da Lei nº 12.345, de 1º de janeiro de 2020.**",
        label_visibility="collapsed",
    )
    st.caption("Dica: use **duplo-asterisco** ao redor de um trecho para deixá-lo em negrito "
               "(comum ao citar a lei que está sendo alterada).")

    st.divider()
    st.subheader("③ Dispositivos")
    st.caption("Adicione na ordem em que devem aparecer no texto final.")

    if not st.session_state.dispositivos:
        if st.button("🚀 Começar com uma estrutura básica (Capítulo + 3 artigos)"):
            base = [
                {"uid": novo_uid(), "tipo": "capitulo", "texto": "", "titulo_estrutura": "DISPOSIÇÕES GERAIS",
                 "paragrafo_unico_forcado": None, "pontuacao_automatica": True},
                {"uid": novo_uid(), "tipo": "artigo",
                 "texto": "[Descreva aqui o objetivo principal desta Lei]",
                 "titulo_estrutura": "", "paragrafo_unico_forcado": None, "pontuacao_automatica": True},
                {"uid": novo_uid(), "tipo": "artigo",
                 "texto": "[Se necessário, descreva obrigações, prazos ou penalidades]",
                 "titulo_estrutura": "", "paragrafo_unico_forcado": None, "pontuacao_automatica": True},
                {"uid": novo_uid(), "tipo": "artigo", "texto": "Esta Lei entra em vigor na data de sua publicação",
                 "titulo_estrutura": "", "paragrafo_unico_forcado": None, "pontuacao_automatica": True},
            ]
            st.session_state.dispositivos.extend(base)
            st.rerun()

    with st.container(border=True):
        tipo_default = sugerir_proximo_tipo()
        tipo_key = f"tipo_select_{len(st.session_state.dispositivos)}"
        tipo = st.selectbox(
            "O que você quer adicionar agora?",
            TIPO_OPCOES,
            index=TIPO_OPCOES.index(tipo_default),
            format_func=lambda t: TIPO_INFO[t]["label"],
            key=tipo_key,
        )
        st.info(TIPO_INFO[tipo]["explicacao"], icon="ℹ️")

        with st.form(f"add_form_{tipo_key}", clear_on_submit=True):
            if tipo in TIPOS_ESTRUTURA:
                titulo_estrutura = st.text_input(
                    "Título deste bloco", placeholder=TIPO_INFO[tipo]["placeholder_bloco"]
                )
                texto = ""
            else:
                texto = st.text_area(
                    "Texto (sem numeração e, se possível, sem pontuação final)",
                    placeholder=TIPO_INFO[tipo]["placeholder_texto"],
                    height=100 if tipo in ("artigo", "paragrafo") else 70,
                )
                titulo_estrutura = ""

            c1, c2 = st.columns(2)
            with c1:
                paragrafo_unico_opt = "Automático"
                if tipo == "paragrafo":
                    paragrafo_unico_opt = st.selectbox(
                        "É o único parágrafo deste artigo?", ["Automático", "Sim", "Não"],
                        help="No automático, o sistema detecta sozinho e escreve \"Parágrafo único.\" quando fizer sentido.",
                    )
            with c2:
                pontuacao_automatica = st.checkbox(
                    "Pontuação automática", value=True,
                    help="Desmarque se preferir digitar você mesmo(a) o ponto final ou o ';' no fim do texto.",
                )

            enviado = st.form_submit_button("➕ Adicionar", use_container_width=True, type="primary")
            if enviado:
                if tipo in TIPOS_ESTRUTURA and not titulo_estrutura.strip():
                    st.warning("Escreva o título deste bloco antes de adicionar.")
                elif tipo not in TIPOS_ESTRUTURA and not texto.strip():
                    st.warning("Escreva o texto antes de adicionar.")
                else:
                    par_unico_forcado = {"Sim": True, "Não": False}.get(paragrafo_unico_opt)
                    st.session_state.dispositivos.append({
                        "uid": novo_uid(), "tipo": tipo, "texto": texto.strip(),
                        "titulo_estrutura": titulo_estrutura.strip(),
                        "paragrafo_unico_forcado": par_unico_forcado,
                        "pontuacao_automatica": pontuacao_automatica,
                    })
                    st.rerun()

    if st.session_state.dispositivos:
        st.markdown("##### Já adicionados")
        for i, d in enumerate(st.session_state.dispositivos):
            depth = PROFUNDIDADE[d["tipo"]]
            indent = "\u00A0\u00A0\u00A0\u00A0" * depth
            preview_txt = d["titulo_estrutura"] if d["tipo"] in TIPOS_ESTRUTURA else d["texto"]
            preview_txt = (preview_txt[:70] + "…") if len(preview_txt) > 70 else preview_txt
            resumo = f"{indent}{ICONE[d['tipo']]} **{ROTULOS[d['tipo']]}** — {preview_txt or '_(vazio)_'}"

            with st.expander(resumo, expanded=(st.session_state.editando_uid == d["uid"])):
                bcol1, bcol2, bcol3, bcol4 = st.columns([1, 1, 1, 2])
                if bcol1.button("↑ Subir", key=f"up_{d['uid']}", disabled=(i == 0)):
                    st.session_state.dispositivos[i - 1], st.session_state.dispositivos[i] = (
                        st.session_state.dispositivos[i], st.session_state.dispositivos[i - 1])
                    st.rerun()
                if bcol2.button("↓ Descer", key=f"down_{d['uid']}", disabled=(i == len(st.session_state.dispositivos) - 1)):
                    st.session_state.dispositivos[i + 1], st.session_state.dispositivos[i] = (
                        st.session_state.dispositivos[i], st.session_state.dispositivos[i + 1])
                    st.rerun()
                if bcol3.button("🗑 Excluir", key=f"del_{d['uid']}"):
                    st.session_state.dispositivos.pop(i)
                    st.rerun()

                st.markdown("**Editar:**")
                if d["tipo"] in TIPOS_ESTRUTURA:
                    novo_titulo = st.text_input("Título do bloco", value=d["titulo_estrutura"], key=f"edit_tit_{d['uid']}")
                    novo_texto = ""
                else:
                    novo_texto = st.text_area("Texto", value=d["texto"], key=f"edit_txt_{d['uid']}", height=90)
                    novo_titulo = ""
                if st.button("💾 Salvar alterações", key=f"save_{d['uid']}"):
                    d["titulo_estrutura"] = novo_titulo.strip()
                    d["texto"] = novo_texto.strip()
                    st.rerun()

        if st.button("Limpar todos os dispositivos"):
            st.session_state.dispositivos = []
            st.rerun()

    st.divider()
    st.subheader("④ Justificação")
    justificativa = st.text_area(
        "Por que essa lei é necessária (separe parágrafos com uma linha em branco)",
        height=180,
        placeholder="A situação atual apresenta o seguinte problema...\n\nA presente proposição busca resolver isso por meio de...",
        label_visibility="collapsed",
    )
    local_data = st.text_input("Fecho", value="Sala das Sessões, na data de sua assinatura.")

# ===========================================================================
# COLUNA DIREITA — pré-visualização em formato de documento
# ===========================================================================
with col_dir:
    st.subheader("Pré-visualização")

    objetos = [dispositivo_para_objeto(d) for d in st.session_state.dispositivos]
    processados = processar_dispositivos(objetos, conector) if objetos else []

    import re as _re

    html = ['<div class="paper">']
    html.append(f'<div class="titulo-pl">PROJETO DE LEI Nº _____, DE {ano or "____"}</div>')
    html.append(f'<div class="autor-pl">({autor_prefixo} {autor_nome or "..."})</div>')

    if ementa.strip():
        ementa_html = _re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', ementa.strip())
        html.append(f'<div class="ementa">{ementa_html}</div>')
    else:
        html.append('<div class="ementa" style="color:#9ca3af;">(a ementa aparecerá aqui)</div>')

    html.append('<div class="decreta">O Congresso Nacional decreta:</div>')

    if not processados:
        html.append('<p style="color:#9ca3af;">(os dispositivos aparecerão aqui conforme você for adicionando)</p>')

    for item in processados:
        if item["tipo"] in TIPOS_ESTRUTURA:
            css_class = "heading" if item["tipo"] in ("titulo", "capitulo") else "secao-h"
            html.append(f'<div class="{css_class}">{item["prefixo"]}</div>')
            if item["titulo_estrutura"]:
                html.append(f'<div class="{css_class}">{item["titulo_estrutura"]}</div>')
        else:
            corpo = _re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', f'{item["prefixo"]} {item["texto"]}')
            html.append(f'<div class="conteudo">{corpo}</div>')

    if justificativa.strip():
        html.append('<div class="justif-h">JUSTIFICAÇÃO</div>')
        for par in [p.strip() for p in justificativa.split("\n\n") if p.strip()]:
            html.append(f'<div class="justif-p">{par.replace(chr(10), " ")}</div>')

    html.append(f'<div class="fecho">{local_data or "Sala das Sessões, na data de sua assinatura."}</div>')
    html.append(f'<div class="assinatura">{autor_prefixo_assinatura} {(autor_nome or "").upper()}</div>')
    html.append(f'<div class="partido">{autor_partido_uf}</div>')
    html.append('</div>')

    st.markdown("".join(html), unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Geração do .docx
# ---------------------------------------------------------------------------
st.divider()
gerar = st.button("📄 Gerar Word (.docx)", type="primary", use_container_width=True)

if gerar:
    if not ementa.strip():
        st.error("Preencha a ementa antes de gerar o documento.")
    elif not st.session_state.dispositivos:
        st.error("Adicione ao menos um dispositivo (ex.: um Artigo) antes de gerar o documento.")
    else:
        dados = {
            "ano": ano.strip() or "2026",
            "ementa": ementa.strip(),
            "autor_nome": autor_nome.strip(),
            "autor_prefixo": autor_prefixo,
            "autor_prefixo_assinatura": autor_prefixo_assinatura,
            "autor_partido_uf": autor_partido_uf.strip(),
            "dispositivos": [dispositivo_para_objeto(d) for d in st.session_state.dispositivos],
            "justificativa": justificativa,
            "local_data": local_data.strip(),
            "conector": conector,
        }
        try:
            docx_bytes = gerar_docx(dados)
            st.success("Documento gerado com sucesso!")
            st.download_button(
                "⬇️ Baixar PL.docx", data=docx_bytes,
                file_name=f"PL_{(autor_nome or 'PL').replace(' ', '_')}_{ano}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erro ao gerar o documento: {e}")
