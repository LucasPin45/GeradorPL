# -*- coding: utf-8 -*-
import io
import uuid
import streamlit as st

from gerador import (
    Dispositivo, TIPOS_ESTRUTURA, TIPOS_CONTEUDO, ROTULOS,
    processar_dispositivos, gerar_docx,
)

st.set_page_config(page_title="Gerador de PL — Gabinete Zanatta", page_icon="📜", layout="wide")

# ---------------------------------------------------------------------------
# Estado
# ---------------------------------------------------------------------------
if "dispositivos" not in st.session_state:
    st.session_state.dispositivos = []  # lista de dicts (serializável)

TIPO_OPCOES = TIPOS_ESTRUTURA + TIPOS_CONTEUDO


def novo_uid():
    return str(uuid.uuid4())[:8]


def dispositivo_para_objeto(d: dict) -> Dispositivo:
    return Dispositivo(
        tipo=d["tipo"],
        texto=d.get("texto", ""),
        titulo_estrutura=d.get("titulo_estrutura", ""),
        paragrafo_unico_forcado=d.get("paragrafo_unico_forcado"),
        pontuacao_automatica=d.get("pontuacao_automatica", True),
        uid=d.get("uid", ""),
    )


# ---------------------------------------------------------------------------
# Sidebar — dados gerais
# ---------------------------------------------------------------------------
st.sidebar.header("Dados gerais do PL")

ano = st.sidebar.text_input("Ano", value="2026")

genero = st.sidebar.radio("Autoria", ["Deputada", "Deputado"], horizontal=True)
autor_nome = st.sidebar.text_input("Nome do autor(a)", value="Júlia Zanatta")
autor_partido_uf = st.sidebar.text_input("Partido/UF", value="PL/SC")

if genero == "Deputada":
    autor_prefixo = "Da Sra."
    autor_prefixo_assinatura = "Deputada"
else:
    autor_prefixo = "Do Sr."
    autor_prefixo_assinatura = "Deputado"

st.sidebar.divider()
conector = st.sidebar.selectbox(
    "Conectivo no penúltimo item de enumerações",
    options=["e", "ou", "nenhum"],
    format_func=lambda x: {"e": '"; e" (regra geral)', "ou": '"; ou" (alternativas)', "nenhum": "apenas \";\""}[x],
)

st.sidebar.divider()
st.sidebar.caption(
    "Modelo base: documento do gabinete (margens, fonte Arial 12, "
    "cabeçalho e rodapé da Câmara dos Deputados). Numeração segue a "
    "LC 95/98 (art. 10 a 12)."
)

# ---------------------------------------------------------------------------
# Corpo principal
# ---------------------------------------------------------------------------
st.title("📜 Gerador de Projeto de Lei")
st.caption("Monte a ementa, os dispositivos e a justificação — o Word sai pronto no padrão do gabinete.")

st.subheader("Ementa")
ementa = st.text_area(
    "Ementa (use **texto** para negrito, ex. na cláusula de alteração de lei)",
    height=90,
    placeholder="Limita o crescimento dos encargos setoriais..., e altera **o art. 13 da Lei nº 10.438, de 26 de abril de 2002.**",
)

st.divider()
st.subheader("Dispositivos")
st.caption(
    "Adicione na ordem em que devem aparecer no texto. A numeração (Art., §, "
    "incisos, alíneas, itens, Capítulos, Seções...) é calculada automaticamente."
)

with st.form("add_dispositivo", clear_on_submit=True):
    col1, col2 = st.columns([1, 3])
    with col1:
        tipo = st.selectbox("Tipo", TIPO_OPCOES, format_func=lambda t: ROTULOS[t])
    with col2:
        if tipo in TIPOS_ESTRUTURA:
            titulo_estrutura = st.text_input(
                "Título do bloco (ex.: DISPOSIÇÕES GERAIS)", key="titulo_estrutura_input"
            )
            texto = ""
        else:
            texto = st.text_area(
                "Texto (sem numeração — ela é gerada automaticamente; "
                "sem pontuação final, se quiser pontuação automática)",
                key="texto_input", height=80,
            )
            titulo_estrutura = ""

    col3, col4 = st.columns(2)
    with col3:
        paragrafo_unico_opt = None
        if tipo == "paragrafo":
            paragrafo_unico_opt = st.selectbox(
                "Parágrafo único?", ["Automático", "Sim", "Não"], key="par_unico_input"
            )
    with col4:
        pontuacao_automatica = st.checkbox("Pontuação automática", value=True, key="pont_auto_input")

    submitted = st.form_submit_button("➕ Adicionar dispositivo")
    if submitted:
        if tipo in TIPOS_ESTRUTURA and not titulo_estrutura.strip():
            st.warning("Informe o título do bloco (ex.: DISPOSIÇÕES GERAIS).")
        elif tipo not in TIPOS_ESTRUTURA and not texto.strip():
            st.warning("Digite o texto do dispositivo.")
        else:
            par_unico_forcado = None
            if paragrafo_unico_opt == "Sim":
                par_unico_forcado = True
            elif paragrafo_unico_opt == "Não":
                par_unico_forcado = False
            st.session_state.dispositivos.append({
                "uid": novo_uid(),
                "tipo": tipo,
                "texto": texto.strip(),
                "titulo_estrutura": titulo_estrutura.strip(),
                "paragrafo_unico_forcado": par_unico_forcado,
                "pontuacao_automatica": pontuacao_automatica,
            })
            st.rerun()

# ---------------------------------------------------------------------------
# Lista de dispositivos já adicionados (com reordenar / remover)
# ---------------------------------------------------------------------------
if st.session_state.dispositivos:
    st.markdown("##### Dispositivos adicionados")
    for i, d in enumerate(st.session_state.dispositivos):
        c1, c2, c3, c4, c5 = st.columns([0.6, 5, 0.5, 0.5, 0.5])
        rotulo = ROTULOS[d["tipo"]]
        preview = d["titulo_estrutura"] if d["tipo"] in TIPOS_ESTRUTURA else d["texto"]
        c1.markdown(f"**{rotulo}**")
        c2.markdown(preview[:140] + ("…" if len(preview) > 140 else ""))
        if c3.button("↑", key=f"up_{d['uid']}") and i > 0:
            st.session_state.dispositivos[i - 1], st.session_state.dispositivos[i] = (
                st.session_state.dispositivos[i], st.session_state.dispositivos[i - 1]
            )
            st.rerun()
        if c4.button("↓", key=f"down_{d['uid']}") and i < len(st.session_state.dispositivos) - 1:
            st.session_state.dispositivos[i + 1], st.session_state.dispositivos[i] = (
                st.session_state.dispositivos[i], st.session_state.dispositivos[i + 1]
            )
            st.rerun()
        if c5.button("🗑", key=f"del_{d['uid']}"):
            st.session_state.dispositivos.pop(i)
            st.rerun()

    if st.button("Limpar todos os dispositivos"):
        st.session_state.dispositivos = []
        st.rerun()
else:
    st.info("Nenhum dispositivo adicionado ainda.")

st.divider()
st.subheader("Justificação")
justificativa = st.text_area(
    "Texto da justificação (separe parágrafos com linha em branco)",
    height=200,
    placeholder="A tarifa de energia elétrica no Brasil deixou de refletir apenas os custos...\n\nA presente proposição busca...",
)

local_data = st.text_input("Fecho", value="Sala das Sessões, na data de sua assinatura.")

st.divider()

# ---------------------------------------------------------------------------
# Pré-visualização da numeração
# ---------------------------------------------------------------------------
st.subheader("Pré-visualização")

if st.session_state.dispositivos:
    objetos = [dispositivo_para_objeto(d) for d in st.session_state.dispositivos]
    processados = processar_dispositivos(objetos, conector)

    linhas = []
    if ementa.strip():
        linhas.append(f"*{ementa.strip()}*")
        linhas.append("")
    for item in processados:
        if item["tipo"] in TIPOS_ESTRUTURA:
            linhas.append(f"**{item['prefixo']}**" + (f" — **{item['titulo_estrutura']}**" if item["titulo_estrutura"] else ""))
        else:
            linhas.append(f"{item['prefixo']} {item['texto']}")
    st.markdown("  \n".join(linhas))
else:
    st.caption("A pré-visualização aparece aqui conforme você adiciona dispositivos.")

st.divider()

# ---------------------------------------------------------------------------
# Geração do .docx
# ---------------------------------------------------------------------------
col_gerar, _ = st.columns([1, 3])
with col_gerar:
    gerar = st.button("📄 Gerar Word (.docx)", type="primary", use_container_width=True)

if gerar:
    if not ementa.strip():
        st.error("Preencha a ementa antes de gerar o documento.")
    elif not st.session_state.dispositivos:
        st.error("Adicione ao menos um dispositivo (ex.: Art. 1º) antes de gerar o documento.")
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
                "⬇️ Baixar PL.docx",
                data=docx_bytes,
                file_name=f"PL_{autor_nome.replace(' ', '_')}_{ano}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erro ao gerar o documento: {e}")
