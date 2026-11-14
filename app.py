# -*- coding: utf-8 -*-
import datetime
import html as html_lib
import re
import unicodedata
from zoneinfo import ZoneInfo

import streamlit as st

import editor
from gerador import (
    Dispositivo, TIPOS_ESTRUTURA, TIPOS_CONTEUDO, TIPO_CITACAO, ROTULOS, PROFUNDIDADE,
    processar_dispositivos, verificar_estrutura, paragrafos_justificacao, limpar_texto,
    gerar_docx, gerar_docx_requerimento,
    ementa_requerimento_urgencia, texto_requerimento_urgencia,
)

st.set_page_config(page_title="Gerador de Documentos — Gabinete Zanatta", page_icon="📜", layout="wide")

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
        "explicacao": "Complementa, detalha ou traz uma exceção à regra do **Artigo** a que pertence. "
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
    TIPO_CITACAO: {
        "label": "❝ Texto da lei alterada  (entre aspas)",
        "explicacao": "Use depois de um Artigo que **altera outra lei** (\"...passa a vigorar com a seguinte "
                       "redação:\"). Cole aqui o novo texto exatamente como deve ficar, com as aspas e o "
                       "**(NR)** no final. O sistema não numera nem pontua este bloco, e ele não conta "
                       "como artigo do projeto. Cada linha vira um parágrafo.",
        "placeholder_texto": "“Art. 5º É vedada a cobrança de...\n§ 1º O disposto no caput não se aplica a...” (NR)",
    },
}

ICONE = {
    "titulo": "📘", "capitulo": "📗", "secao": "📙", "subsecao": "📕",
    "artigo": "📄", "paragrafo": "↳", "inciso": "•", "alinea": "◦", "item": "▪", TIPO_CITACAO: "❝",
}

TIPO_OPCOES = TIPOS_ESTRUTURA + TIPOS_CONTEUDO + [TIPO_CITACAO]

SIGLAS_COMUNS = ["PL", "PLP", "PEC", "PDL", "PRC", "MPV", "REQ", "INC", "Outro"]

DOC_PL = "Projeto de Lei (PL)"
DOC_REQ = "Requerimento de Urgência"
FECHO_PL = "Sala das Sessões, na data de sua assinatura."
FECHO_REQ = "Sala das Sessões, na data de sua assinatura"
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

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
.paper .vazio { color:#9ca3af; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Estado
# ---------------------------------------------------------------------------
# Tudo o que o usuário digita fica no session_state sob estas chaves. Os campos
# do módulo que não está na tela deixam de ser desenhados, e o Streamlit apaga
# o valor de widget não desenhado — por isso cada chave é regravada a cada
# execução. Sem isso, trocar de "Projeto de Lei" para "Requerimento" e voltar
# apagava a ementa e a justificação.
PADROES = {
    "ano": str(datetime.datetime.now(ZoneInfo("America/Sao_Paulo")).year),
    "genero": "Deputada",
    "autor_nome": "Júlia Zanatta",
    "autor_partido_uf": "PL/SC",
    "conector": "e",
    "tipo_documento": DOC_PL,
    "ementa": "",
    "justificativa": "",
    "fecho_pl": FECHO_PL,
    "artigo_ricd": "155",
    "sigla_materia": "PL",
    "sigla_outra": "",
    "numero_materia": "",
    "ementa_materia": "",
    "fecho_req": FECHO_REQ,
}
OPCOES_VALIDAS = {
    "genero": ["Deputada", "Deputado"],
    "conector": ["e", "ou", "nenhum"],
    "tipo_documento": [DOC_PL, DOC_REQ],
    "artigo_ricd": ["154", "155"],
    "sigla_materia": SIGLAS_COMUNS,
}

for _k, _v in PADROES.items():
    st.session_state.setdefault(_k, _v)
    st.session_state[_k] = st.session_state[_k]
st.session_state.setdefault("dispositivos", [])
st.session_state.setdefault("posicao_add", None)

# Valores que precisam ser gravados antes de o widget correspondente ser desenhado
if "_posicao_pendente" in st.session_state:
    st.session_state["posicao_add"] = st.session_state.pop("_posicao_pendente")
if st.session_state["posicao_add"] not in [d["uid"] for d in st.session_state.dispositivos]:
    st.session_state["posicao_add"] = None


def dispositivo_para_objeto(d: dict) -> Dispositivo:
    return Dispositivo(
        tipo=d["tipo"], texto=d.get("texto", ""),
        titulo_estrutura=d.get("titulo_estrutura", ""),
        paragrafo_unico_forcado=d.get("paragrafo_unico_forcado"),
        pontuacao_automatica=d.get("pontuacao_automatica", True),
        uid=d.get("uid", ""),
    )


def negrito_html(texto: str) -> str:
    """Texto do usuário → HTML seguro, com **trecho** em negrito (como no Word)."""
    return re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', html_lib.escape(texto or ""))


def nome_de_arquivo(texto: str) -> str:
    """Parte de nome de arquivo: sem acentos, espaços nem caracteres problemáticos.
    (Com acento no nome — "Júlia" — o navegador salvava o arquivo como "download".)"""
    sem_acento = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode("ascii")
    limpo = re.sub(r'[^A-Za-z0-9.\-]+', '_', sem_acento.replace("/", "-"))
    return limpo.strip("_") or "documento"


def guardar_para_desfazer() -> None:
    st.session_state["_desfazer"] = list(st.session_state.dispositivos)


def aplicar_rascunho(campos: dict, dispositivos: list) -> None:
    for chave, valor in campos.items():
        if chave in OPCOES_VALIDAS and valor not in OPCOES_VALIDAS[chave]:
            continue        # opção desconhecida: mantém a atual
        st.session_state[chave] = valor
    st.session_state.dispositivos = dispositivos
    st.session_state["posicao_add"] = None
    st.session_state.pop("_desfazer", None)
    # campos de edição guardam o texto antigo de cada dispositivo
    for chave in [k for k in st.session_state.keys() if str(k).startswith("edit_")]:
        del st.session_state[chave]


# ---------------------------------------------------------------------------
# Sidebar — dados gerais (comuns a todos os módulos) e rascunho
# ---------------------------------------------------------------------------
# O rascunho aparece abaixo dos dados gerais, mas é processado antes: carregar
# um arquivo grava os campos antes de eles serem desenhados.
area_dados = st.sidebar.container()
area_rascunho = st.sidebar.container()

with area_rascunho:
    st.divider()
    st.header("💾 Rascunho")
    st.caption("O que você digita fica só nesta aba do navegador e se perde se a página for "
               "recarregada. Salve um rascunho para continuar depois.")
    arquivo = st.file_uploader("Abrir rascunho salvo", type=["json"], key="rascunho_upload",
                               help="Substitui o que está na tela pelo conteúdo do arquivo.")
    if arquivo is not None:
        marca = f"{arquivo.name}:{arquivo.size}:{getattr(arquivo, 'file_id', '')}"
        if st.session_state.get("_rascunho_aberto") != marca:
            try:
                campos_lidos, dispositivos_lidos = editor.importar_rascunho(arquivo.getvalue().decode("utf-8"))
            except (ValueError, UnicodeDecodeError) as erro:
                st.error(f"Não foi possível abrir: {erro}.")
            else:
                aplicar_rascunho(campos_lidos, dispositivos_lidos)
                st.session_state["_rascunho_aberto"] = marca
                st.success(f"Rascunho aberto: {len(dispositivos_lidos)} dispositivo(s).")
    st.download_button(
        "⬇️ Salvar rascunho",
        data=editor.exportar_rascunho({k: st.session_state[k] for k in PADROES}, st.session_state.dispositivos),
        file_name=f"rascunho_{nome_de_arquivo(st.session_state['autor_nome'])}_"
                  f"{datetime.datetime.now(ZoneInfo('America/Sao_Paulo')):%Y%m%d_%H%M}.json",
        mime="application/json", use_container_width=True, key="rascunho_download",
        help="Baixa um arquivo com tudo o que foi digitado nos dois módulos.",
    )

with area_dados:
    st.header("① Dados gerais")
    ano = st.text_input("Ano", key="ano")
    genero = st.radio("Autoria", OPCOES_VALIDAS["genero"], horizontal=True, key="genero")
    autor_nome = st.text_input("Nome do autor(a)", key="autor_nome")
    autor_partido_uf = st.text_input("Partido/UF", key="autor_partido_uf")

    st.divider()
    conector = st.selectbox(
        "Conectivo no penúltimo item de uma lista (só afeta o módulo PL)",
        options=OPCOES_VALIDAS["conector"],
        format_func=lambda x: {"e": '"; e"  (regra geral)', "ou": '"; ou"  (alternativas)', "nenhum": 'apenas ";"'}[x],
        help="Usado automaticamente no penúltimo Inciso/Alínea/Item de uma lista. Ex.: 'I – ...; II – ...; e III – ...'.",
        key="conector",
    )

    st.divider()
    st.caption(
        "O Word gerado usa a mesma fonte, margens, cabeçalho e rodapé do modelo "
        "oficial do gabinete."
    )

autor_prefixo = "Da Sra." if genero == "Deputada" else "Do Sr."
autor_prefixo_assinatura_pl = genero
autor_prefixo_assinatura_req = f"{genero} Federal"

# ---------------------------------------------------------------------------
# Cabeçalho + seletor de módulo
# ---------------------------------------------------------------------------
st.title("📜 Gerador de Documentos Legislativos")
st.caption("Preencha em linguagem simples — a formatação técnica fica por conta do sistema.")

tipo_documento = st.radio("Tipo de documento", OPCOES_VALIDAS["tipo_documento"], horizontal=True,
                          key="tipo_documento")
st.divider()

# ===========================================================================
# MÓDULO 1 — PROJETO DE LEI
# ===========================================================================
if tipo_documento == DOC_PL:

    with st.expander("💡 Guia rápido — o que é cada coisa?"):
        st.markdown("""
Pense assim, sem se preocupar com os nomes jurídicos:

- **Artigo** = uma regra. É a espinha dorsal do texto — cada uma traz um comando ou determinação.
- **Parágrafo** = uma exceção ou um detalhe daquela regra (do Artigo a que pertence).
- **Inciso** = um item de uma lista dentro de um Artigo ou Parágrafo.
- **Alínea** = um item dentro de um Inciso (raro).
- **Capítulo / Seção** = apenas organizam os Artigos por tema, como os capítulos de um livro —
  não têm "conteúdo" próprio, só um título.
- **Texto da lei alterada** = quando o projeto muda outra lei, é o novo texto entre aspas, que
  vem depois do Artigo "...passa a vigorar com a seguinte redação:".

Você **não digita numeração nem pontuação final** — o sistema calcula "Art. 1º", "§ 2º", "I –", "a)"
e o "." ou ";" no fim de cada frase automaticamente.
""")

    objetos = [dispositivo_para_objeto(d) for d in st.session_state.dispositivos]
    processados = processar_dispositivos(objetos, conector) if objetos else []
    avisos = verificar_estrutura(objetos, conector)
    uids_com_aviso = {a["uid"] for a in avisos if a["uid"]}

    def rotulo_curto(i: int, limite: int = 60) -> str:
        """ "Art. 2º — começo do texto…" para listas e seletores."""
        d, proc = st.session_state.dispositivos[i], processados[i]
        if d["tipo"] in TIPOS_ESTRUTURA:
            cabeca, corpo = proc["prefixo"], proc["titulo_estrutura"]
        elif d["tipo"] == TIPO_CITACAO:
            cabeca, corpo = ROTULOS[TIPO_CITACAO], proc["texto"].replace("\n", " ")
        else:
            cabeca, corpo = proc["prefixo"].rstrip(" –"), limpar_texto(d["texto"])
        corpo = (corpo[:limite] + "…") if len(corpo) > limite else corpo
        return f"{cabeca} — {corpo}" if corpo else f"{cabeca} — (vazio)"

    col_esq, col_dir = st.columns([1, 1], gap="large")

    # -----------------------------------------------------------------
    # Coluna esquerda — construção do texto
    # -----------------------------------------------------------------
    with col_esq:
        st.subheader("② Ementa")
        ementa = st.text_area(
            "O resumo do que a lei faz, em uma frase",
            height=80,
            placeholder="Institui o Programa XYZ e dá outras providências, e altera "
                        "**o art. 5º da Lei nº 12.345, de 1º de janeiro de 2020.**",
            label_visibility="collapsed",
            key="ementa",
        )
        st.caption("Dica: use **duplo-asterisco** ao redor de um trecho para deixá-lo em negrito "
                   "(comum ao citar a lei que está sendo alterada).")

        st.divider()
        st.subheader("③ Dispositivos")
        st.caption("Adicione na ordem em que devem aparecer no texto final — ou escolha depois de "
                   "qual dispositivo o novo deve entrar.")

        if not st.session_state.dispositivos:
            if st.button("🚀 Começar com uma estrutura básica (Capítulo + 3 artigos)"):
                st.session_state.dispositivos = [
                    editor.novo_dispositivo("capitulo", titulo_estrutura="DISPOSIÇÕES GERAIS"),
                    editor.novo_dispositivo("artigo", "[Descreva aqui o objetivo principal desta Lei]"),
                    editor.novo_dispositivo("artigo", "[Se necessário, descreva obrigações, prazos ou penalidades]"),
                    editor.novo_dispositivo("artigo", "Esta Lei entra em vigor na data de sua publicação"),
                ]
                st.rerun()

        with st.container(border=True):
            uids = [d["uid"] for d in st.session_state.dispositivos]
            indice_por_uid = {u: i for i, u in enumerate(uids)}
            if uids:
                posicao = st.selectbox(
                    "Onde o novo dispositivo deve entrar?",
                    [None] + uids,
                    format_func=lambda u: "No final do texto" if u is None else f"Depois de: {rotulo_curto(indice_por_uid[u])}",
                    key="posicao_add",
                    help="Escolhendo um Artigo, o novo dispositivo entra depois dos parágrafos e incisos dele.",
                )
            else:
                posicao = None

            tipo_default = editor.proximo_tipo_sugerido(st.session_state.dispositivos, posicao)
            tipo_key = f"tipo_select_{len(st.session_state.dispositivos)}_{posicao}"
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
                elif tipo == TIPO_CITACAO:
                    texto = st.text_area(
                        "Novo texto da lei alterada (com as aspas e o (NR) no fim)",
                        placeholder=TIPO_INFO[tipo]["placeholder_texto"], height=150,
                    )
                    titulo_estrutura = ""
                else:
                    texto = st.text_area(
                        "Texto (sem numeração e, se possível, sem pontuação final)",
                        placeholder=TIPO_INFO[tipo]["placeholder_texto"],
                        height=100 if tipo in ("artigo", "paragrafo") else 70,
                    )
                    titulo_estrutura = ""

                paragrafo_unico_opt = "Automático"
                pontuacao_automatica = True
                if tipo not in TIPOS_ESTRUTURA and tipo != TIPO_CITACAO:
                    c1, c2 = st.columns(2)
                    with c1:
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
                        novo = editor.novo_dispositivo(
                            tipo, texto.strip(), titulo_estrutura.strip(),
                            {"Sim": True, "Não": False}.get(paragrafo_unico_opt), pontuacao_automatica)
                        st.session_state.dispositivos = editor.inserir(st.session_state.dispositivos, novo, posicao)
                        if posicao is not None:
                            # quem está inserindo no meio costuma seguir dali: o próximo entra depois deste
                            st.session_state["_posicao_pendente"] = novo["uid"]
                        st.rerun()

        if st.session_state.dispositivos:
            st.markdown("##### Já adicionados")
            lista = st.session_state.dispositivos
            for i, d in enumerate(lista):
                indent = "\u00A0\u00A0\u00A0\u00A0" * (0 if d["tipo"] in TIPOS_ESTRUTURA else PROFUNDIDADE[d["tipo"]] - 1)
                alerta = " ⚠️" if d["uid"] in uids_com_aviso else ""
                cabeca, _, corpo = rotulo_curto(i, 70).partition(" — ")
                resumo = f"{indent}{ICONE[d['tipo']]} **{cabeca}** — {corpo}{alerta}"
                n_dep = editor.dependentes(lista, i)
                junto = f" (leva junto {n_dep} dispositivo(s) que dependem dele)" if n_dep else ""

                with st.expander(resumo):
                    for a in avisos:
                        if a["uid"] == d["uid"]:
                            st.warning(f"**{a['rotulo']}** {a['aviso']}", icon="⚠️")

                    bcol1, bcol2, bcol3, bcol4 = st.columns([1, 1, 1, 2])
                    if bcol1.button("↑ Subir", key=f"up_{d['uid']}", disabled=not editor.pode_subir(lista, i),
                                    help=f"Sobe uma posição{junto}."):
                        st.session_state.dispositivos = editor.subir(lista, i)
                        st.rerun()
                    if bcol2.button("↓ Descer", key=f"down_{d['uid']}", disabled=not editor.pode_descer(lista, i),
                                    help=f"Desce uma posição{junto}."):
                        st.session_state.dispositivos = editor.descer(lista, i)
                        st.rerun()
                    if bcol3.button("🗑 Excluir", key=f"del_{d['uid']}",
                                    help=f"Exclui este dispositivo{junto}. Dá para desfazer."):
                        guardar_para_desfazer()
                        st.session_state.dispositivos = editor.excluir(lista, i)
                        st.rerun()

                    st.markdown("**Editar:**")
                    novo_unico = d.get("paragrafo_unico_forcado")
                    nova_pontuacao = d.get("pontuacao_automatica", True)
                    if d["tipo"] in TIPOS_ESTRUTURA:
                        novo_titulo = st.text_input("Título do bloco", value=d["titulo_estrutura"], key=f"edit_tit_{d['uid']}")
                        novo_texto = ""
                    else:
                        novo_texto = st.text_area("Texto", value=d["texto"], key=f"edit_txt_{d['uid']}",
                                                  height=150 if d["tipo"] == TIPO_CITACAO else 90)
                        novo_titulo = ""
                        if d["tipo"] != TIPO_CITACAO:
                            e1, e2 = st.columns(2)
                            if d["tipo"] == "paragrafo":
                                opcoes_unico = ["Automático", "Sim", "Não"]
                                atual = {True: "Sim", False: "Não"}.get(novo_unico, "Automático")
                                escolha = e1.selectbox("É o único parágrafo deste artigo?", opcoes_unico,
                                                       index=opcoes_unico.index(atual), key=f"edit_unico_{d['uid']}")
                                novo_unico = {"Sim": True, "Não": False}.get(escolha)
                            nova_pontuacao = e2.checkbox("Pontuação automática", value=nova_pontuacao,
                                                         key=f"edit_pont_{d['uid']}")
                    if st.button("💾 Salvar alterações", key=f"save_{d['uid']}"):
                        d["titulo_estrutura"] = novo_titulo.strip()
                        d["texto"] = novo_texto.strip()
                        d["paragrafo_unico_forcado"] = novo_unico
                        d["pontuacao_automatica"] = nova_pontuacao
                        st.rerun()

            if st.session_state.get("_confirmar_limpar"):
                st.warning("Apagar **todos** os dispositivos?")
                s1, s2 = st.columns(2)
                if s1.button("Sim, apagar tudo", key="limpar_sim", type="primary", use_container_width=True):
                    guardar_para_desfazer()
                    st.session_state.dispositivos = []
                    st.session_state["_confirmar_limpar"] = False
                    st.rerun()
                if s2.button("Cancelar", key="limpar_nao", use_container_width=True):
                    st.session_state["_confirmar_limpar"] = False
                    st.rerun()
            elif st.button("Limpar todos os dispositivos", key="limpar"):
                st.session_state["_confirmar_limpar"] = True
                st.rerun()

        if st.session_state.get("_desfazer") is not None:
            if st.button("↩️ Desfazer a última exclusão", key="desfazer"):
                st.session_state.dispositivos = st.session_state.pop("_desfazer")
                st.rerun()

        st.divider()
        st.subheader("④ Justificação")
        justificativa = st.text_area(
            "Por que essa lei é necessária (separe parágrafos com uma linha em branco)",
            height=180,
            placeholder="A situação atual apresenta o seguinte problema...\n\nA presente proposição busca resolver isso por meio de...",
            label_visibility="collapsed",
            key="justificativa",
        )
        st.caption("Separe os parágrafos com uma linha em branco. O **duplo-asterisco** também funciona aqui.")
        local_data = st.text_input("Fecho", key="fecho_pl")

    # -----------------------------------------------------------------
    # Coluna direita — pré-visualização
    # -----------------------------------------------------------------
    with col_dir:
        st.subheader("Pré-visualização")

        html = ['<div class="paper">']
        html.append(f'<div class="titulo-pl">PROJETO DE LEI Nº _____, DE {html_lib.escape(ano.strip() or "____")}</div>')
        html.append(f'<div class="autor-pl">({autor_prefixo} {html_lib.escape(autor_nome.strip() or "...")})</div>')

        if ementa.strip():
            html.append(f'<div class="ementa">{negrito_html(limpar_texto(ementa))}</div>')
        else:
            html.append('<div class="ementa vazio">(a ementa aparecerá aqui)</div>')

        html.append('<div class="decreta">O Congresso Nacional decreta:</div>')

        if not processados:
            html.append('<p class="vazio">(os dispositivos aparecerão aqui conforme você for adicionando)</p>')

        for item in processados:
            if item["tipo"] in TIPOS_ESTRUTURA:
                css_class = "heading" if item["tipo"] in ("titulo", "capitulo") else "secao-h"
                html.append(f'<div class="{css_class}">{html_lib.escape(item["prefixo"])}</div>')
                if item["titulo_estrutura"]:
                    html.append(f'<div class="{css_class}">{html_lib.escape(item["titulo_estrutura"])}</div>')
            elif item["tipo"] == TIPO_CITACAO:
                for linha in item["texto"].split("\n"):
                    html.append(f'<div class="conteudo">{negrito_html(linha)}</div>')
            else:
                html.append(f'<div class="conteudo">{negrito_html(item["prefixo"] + " " + item["texto"])}</div>')

        html.append('<div class="justif-h">JUSTIFICAÇÃO</div>')
        pars_justificacao = paragrafos_justificacao(justificativa)
        if not pars_justificacao:
            html.append('<div class="justif-p vazio">(a justificação aparecerá aqui)</div>')
        for par in pars_justificacao:
            html.append(f'<div class="justif-p">{negrito_html(par)}</div>')

        html.append(f'<div class="fecho">{html_lib.escape(local_data.strip() or FECHO_PL)}</div>')
        html.append(f'<div class="assinatura">{autor_prefixo_assinatura_pl} {html_lib.escape(autor_nome.strip().upper())}</div>')
        html.append(f'<div class="partido">{html_lib.escape(autor_partido_uf.strip())}</div>')
        html.append('</div>')

        st.markdown("".join(html), unsafe_allow_html=True)

        if st.session_state.dispositivos:
            if avisos:
                with st.expander(f"⚠️ {len(avisos)} ponto(s) de atenção na estrutura", expanded=True):
                    for a in avisos:
                        st.markdown(f"- **{a['rotulo']}** {a['aviso']}")
                    st.caption("São avisos de técnica legislativa: não impedem gerar o documento.")
            else:
                st.caption("✅ Estrutura sem pendências de técnica legislativa.")

    st.divider()
    faltando = []
    if not ementa.strip():
        faltando.append("a ementa")
    if not st.session_state.dispositivos:
        faltando.append("ao menos um dispositivo (ex.: um Artigo)")

    if faltando:
        st.info("Para gerar o Word, falta preencher " + " e ".join(faltando) + ".")
        st.button("📄 Baixar Word (.docx)", type="primary", use_container_width=True, key="gerar_pl", disabled=True)
    else:
        dados = {
            "ano": ano.strip() or PADROES["ano"],
            "ementa": ementa.strip(),
            "autor_nome": autor_nome.strip(),
            "autor_prefixo": autor_prefixo,
            "autor_prefixo_assinatura": autor_prefixo_assinatura_pl,
            "autor_partido_uf": autor_partido_uf.strip(),
            "dispositivos": objetos,
            "justificativa": justificativa,
            "local_data": local_data.strip() or FECHO_PL,
            "conector": conector,
        }
        try:
            docx_bytes = gerar_docx(dados)
        except Exception as e:
            st.error(f"Erro ao gerar o documento: {e}")
        else:
            st.download_button(
                "📄 Baixar Word (.docx)", data=docx_bytes,
                file_name=f"PL_{nome_de_arquivo(autor_nome)}_{nome_de_arquivo(ano)}.docx",
                mime=MIME_DOCX, type="primary", use_container_width=True, key="gerar_pl",
            )
            if not justificativa.strip():
                st.caption("⚠️ A justificação está vazia: o Word sairá só com o título \"JUSTIFICAÇÃO\".")

# ===========================================================================
# MÓDULO 2 — REQUERIMENTO DE URGÊNCIA
# ===========================================================================
else:
    with st.expander("💡 Guia rápido"):
        st.markdown("""
Esse módulo gera o requerimento padrão de urgência para apreciação de uma matéria
(PL, PLP, PEC, PDL...), nos termos do art. 154 ou 155 do Regimento Interno da
Câmara dos Deputados (RICD). Você só preenche os espaços em branco abaixo — a
**ementa é montada sozinha**, no mesmo formato usado nos Projetos de Lei:

> Requeremos, nos termos do art. **_____** [154 ou 155] do Regimento Interno da
> Câmara dos Deputados, urgência para apreciação do **_____** [sigla] **_____**
> [número/ano], que "**_____**" [ementa da matéria].

**Confirme com a assessoria legislativa qual artigo (154 ou 155) se aplica ao caso
concreto** antes de protocolar — a escolha depende do rito de urgência pretendido.
""")

    col_esq, col_dir = st.columns([1, 1], gap="large")

    with col_esq:
        st.subheader("② Dados do requerimento")

        artigo_ricd = st.radio(
            "Fundamento regimental", OPCOES_VALIDAS["artigo_ricd"], horizontal=True,
            format_func=lambda x: f"Art. {x} do RICD", key="artigo_ricd",
        )

        c1, c2 = st.columns([1, 2])
        with c1:
            sigla_materia = st.selectbox("Sigla da matéria", SIGLAS_COMUNS, key="sigla_materia")
            if sigla_materia == "Outro":
                sigla_materia = st.text_input("Digite a sigla", key="sigla_outra").strip().upper()
        with c2:
            numero_materia = st.text_input("Número/ano da matéria", placeholder="2.548/2025", key="numero_materia")

        ementa_materia = st.text_area(
            "Ementa oficial da matéria (cole exatamente como está registrada — "
            "vai aparecer entre aspas, sem precisar digitar as aspas)",
            height=110,
            placeholder='Altera a Lei nº 13.465, de 11 de julho de 2017, para garantir o direito de '
                        'propriedade dos moradores ocupantes anteriormente à decretação de Área de '
                        'Proteção Ambiental (APA), e dá outras providências.',
            key="ementa_materia",
        )

        fecho_req = st.text_input("Fecho", key="fecho_req")

    with col_dir:
        st.subheader("Pré-visualização")

        ementa_req = ementa_requerimento_urgencia(artigo_ricd, sigla_materia, numero_materia)
        texto_requerimento = texto_requerimento_urgencia(artigo_ricd, sigla_materia, numero_materia, ementa_materia)

        html = ['<div class="paper">']
        html.append(f'<div class="titulo-pl">REQUERIMENTO Nº _____ DE {html_lib.escape(ano.strip() or "____")}</div>')
        html.append('<div style="height:14px;"></div>')
        html.append(f'<div class="ementa">{html_lib.escape(ementa_req)}</div>')
        html.append('<div style="height:14px;"></div>')
        html.append('<div class="conteudo" style="text-indent:0;">Senhor Presidente,</div>')
        html.append(f'<div class="conteudo">{html_lib.escape(texto_requerimento)}</div>')
        html.append('<div style="height:40px;"></div>')
        html.append(f'<div class="fecho">{html_lib.escape(fecho_req.strip() or FECHO_REQ)}</div>')
        html.append(f'<div class="assinatura">{autor_prefixo_assinatura_req} {html_lib.escape(autor_nome.strip() or "...")}</div>')
        html.append(f'<div class="partido">({html_lib.escape(autor_partido_uf.strip())})</div>')
        html.append('</div>')

        st.markdown("".join(html), unsafe_allow_html=True)

        st.caption(f"📝 Ementa gerada automaticamente: *{ementa_req}*")

    st.divider()
    faltando = []
    if not sigla_materia:
        faltando.append("a sigla da matéria")
    if not numero_materia.strip():
        faltando.append("o número/ano da matéria")
    if not ementa_materia.strip():
        faltando.append("a ementa da matéria")

    if faltando:
        st.info("Para gerar o Word, falta preencher " + ", ".join(faltando) + ".")
        st.button("📄 Baixar Word (.docx)", type="primary", use_container_width=True, key="gerar_req", disabled=True)
    else:
        dados = {
            "ano": ano.strip() or PADROES["ano"],
            "artigo_ricd": artigo_ricd,
            "sigla_materia": sigla_materia,
            "numero_materia": numero_materia.strip(),
            "ementa_materia": ementa_materia.strip(),
            "autor_nome": autor_nome.strip(),
            "autor_prefixo_assinatura": autor_prefixo_assinatura_req,
            "autor_partido_uf": autor_partido_uf.strip(),
            "fecho": fecho_req.strip() or FECHO_REQ,
        }
        try:
            docx_bytes = gerar_docx_requerimento(dados)
        except Exception as e:
            st.error(f"Erro ao gerar o documento: {e}")
        else:
            st.download_button(
                "📄 Baixar Word (.docx)", data=docx_bytes,
                file_name=f"REQ_Urgencia_{nome_de_arquivo(sigla_materia)}_{nome_de_arquivo(numero_materia)}.docx",
                mime=MIME_DOCX, type="primary", use_container_width=True, key="gerar_req",
            )
