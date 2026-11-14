# -*- coding: utf-8 -*-
"""Testes do gerador: numeração, pontuação (LC 95/98), verificação de
estrutura e geração do .docx. Rodar com:  python -m pytest  (ou python tests/test_gerador.py)."""
import io
import os
import sys
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gerador import (  # noqa: E402
    Dispositivo as D, processar_dispositivos, verificar_estrutura, gerar_docx, gerar_docx_requerimento,
    ementa_requerimento_urgencia, texto_requerimento_urgencia, limpar_texto, paragrafos_justificacao,
    to_roman, to_letra,
)


def linhas(dispositivos, conector="e"):
    return [f'{i["prefixo"]} {i["texto"]}'.strip() for i in processar_dispositivos(dispositivos, conector)]


def corpo_docx(conteudo: bytes) -> str:
    return zipfile.ZipFile(io.BytesIO(conteudo)).read("word/document.xml").decode("utf-8")


DADOS_PL = {
    "ano": "2026", "ementa": "Institui o Programa XYZ.", "autor_nome": "Júlia Zanatta",
    "autor_prefixo": "Da Sra.", "autor_prefixo_assinatura": "Deputada", "autor_partido_uf": "PL/SC",
    "justificativa": "Primeiro parágrafo.\n\nSegundo parágrafo.", "local_data": "Sala das Sessões.", "conector": "e",
}


# ── numeração ──────────────────────────────────────────────────────────────

def test_artigos_ordinal_ate_o_nono_cardinal_depois():
    r = linhas([D("artigo", f"Regra {i}") for i in range(11)])
    assert r[8] == "Art. 9º Regra 8."
    assert r[9] == "Art. 10. Regra 9."
    assert r[10] == "Art. 11. Regra 10."


def test_paragrafo_unico_e_numerado():
    assert linhas([D("artigo", "A"), D("paragrafo", "B")])[1] == "Parágrafo único. B."
    r = linhas([D("artigo", "A"), D("paragrafo", "B"), D("paragrafo", "C")])
    assert r[1:] == ["§ 1º B.", "§ 2º C."]
    # forçado pelo usuário
    assert linhas([D("artigo", "A"), D("paragrafo", "B", paragrafo_unico_forcado=False)])[1] == "§ 1º B."


def test_numeracao_reinicia_a_cada_artigo_e_paragrafo():
    r = linhas([D("artigo", "A"), D("inciso", "x"), D("inciso", "y"), D("paragrafo", "B"), D("inciso", "z"), D("inciso", "w"),
                D("artigo", "C"), D("paragrafo", "D")])
    assert r == ["Art. 1º A:", "I – x; e", "II – y.", "Parágrafo único. B:", "I – z; e", "II – w.",
                 "Art. 2º C.", "Parágrafo único. D."]


def test_estrutura_em_romanos_e_caixa_alta():
    p = processar_dispositivos([D("titulo", "", "das disposições"), D("capitulo", "", "do objeto"),
                                D("secao", "", "Da Fiscalização"), D("capitulo", "", "final"), D("secao", "", "Outra")])
    assert [i["prefixo"] for i in p] == ["TÍTULO I", "CAPÍTULO I", "Seção I", "CAPÍTULO II", "Seção I"]
    assert p[0]["titulo_estrutura"] == "DAS DISPOSIÇÕES" and p[2]["titulo_estrutura"] == "Da Fiscalização"


def test_romanos_e_letras():
    assert [to_roman(n) for n in (1, 4, 9, 14, 40, 90)] == ["I", "IV", "IX", "XIV", "XL", "XC"]
    assert [to_letra(n) for n in (1, 26, 27)] == ["a", "z", "aa"]


# ── pontuação ──────────────────────────────────────────────────────────────

def test_lista_simples():
    r = linhas([D("artigo", "São beneficiários"), D("inciso", "a"), D("inciso", "b"), D("inciso", "c")])
    assert r == ["Art. 1º São beneficiários:", "I – a;", "II – b; e", "III – c."]


def test_conector_ou_e_nenhum():
    ds = [D("artigo", "X"), D("inciso", "a"), D("inciso", "b"), D("inciso", "c")]
    assert linhas(ds, "ou")[2] == "II – b; ou"
    assert linhas(ds, "nenhum")[2] == "II – b;"


def test_ultima_alinea_de_inciso_que_nao_e_o_ultimo_termina_em_ponto_e_virgula():
    r = linhas([D("artigo", "São beneficiários"), D("inciso", "os servidores"), D("alinea", "ativos"), D("alinea", "inativos"),
                D("inciso", "os militares"), D("inciso", "os pensionistas")])
    assert r == ["Art. 1º São beneficiários:", "I – os servidores:", "a) ativos; e", "b) inativos;",
                 "II – os militares; e", "III – os pensionistas."]


def test_penultimo_inciso_e_o_que_antecede_o_ultimo_mesmo_com_alineas_no_meio():
    r = linhas([D("artigo", "O programa compreende"), D("inciso", "saúde"), D("inciso", "educação"),
                D("alinea", "básica"), D("alinea", "superior"), D("inciso", "segurança")])
    assert r == ["Art. 1º O programa compreende:", "I – saúde;", "II – educação:", "a) básica; e",
                 "b) superior; e", "III – segurança."]


def test_ultimo_item_de_tudo_termina_em_ponto():
    r = linhas([D("artigo", "X"), D("inciso", "a"), D("inciso", "b"), D("alinea", "c"), D("item", "d"), D("item", "e"),
                D("paragrafo", "Y")])
    assert r[-3:] == ["1. d; e", "2. e.", "Parágrafo único. Y."]


def test_respeita_pontuacao_digitada():
    r = linhas([D("artigo", "Já tem ponto."), D("artigo", "Termina com dois-pontos:"), D("inciso", "item digitado; e"),
                D("inciso", "outro"), D("artigo", "Pergunta?")])
    assert r == ["Art. 1º Já tem ponto.", "Art. 2º Termina com dois-pontos:", "I – item digitado; e", "II – outro.",
                 "Art. 3º Pergunta?"]


def test_nr_nao_recebe_ponto_e_ordinal_recebe():
    r = linhas([D("artigo", "O art. 5º passa a vigorar assim: “Art. 5º É vedado.” (NR)"), D("artigo", "Fica revogado o art. 7º"),
                D("paragrafo", "Aplica-se o § 2º")])
    assert r[0].endswith("(NR)")
    assert r[1] == "Art. 2º Fica revogado o art. 7º."
    assert r[2] == "Parágrafo único. Aplica-se o § 2º."


def test_pontuacao_automatica_desligada():
    assert linhas([D("artigo", "Sem ponto", pontuacao_automatica=False)]) == ["Art. 1º Sem ponto"]


def test_texto_de_lei_alterada_nao_e_numerado_nem_pontuado():
    r = processar_dispositivos([
        D("artigo", "O art. 5º da Lei nº 1 passa a vigorar com a seguinte redação"),
        D("citacao", "“Art. 5º É vedado o uso.\n§ 1º Excetua-se o caso tal.” (NR)"),
        D("artigo", "Esta Lei entra em vigor na data de sua publicação")])
    assert r[0]["texto"].endswith("redação:")
    assert r[1]["prefixo"] == "" and r[1]["texto"] == "“Art. 5º É vedado o uso.\n§ 1º Excetua-se o caso tal.” (NR)"
    assert r[2]["prefixo"] == "Art. 2º"       # a citação não conta como artigo


# ── limpeza de texto ───────────────────────────────────────────────────────

def test_limpar_texto():
    assert limpar_texto("  linha um\n linha\tdois\x0b\x0c  ") == "linha um linha dois"
    assert limpar_texto("a\r\n\r\nb  c", manter_quebras=True) == "a\nb c"


def test_paragrafos_da_justificacao():
    assert paragrafos_justificacao("Um\ncontinua.\n\n \nDois.") == ["Um continua.", "Dois."]
    assert paragrafos_justificacao("   ") == []


# ── verificação de estrutura ───────────────────────────────────────────────

def avisos(dispositivos):
    return [f'{a["rotulo"]}: {a["aviso"]}' for a in verificar_estrutura(dispositivos)]


def test_estrutura_correta_nao_gera_aviso():
    assert avisos([D("capitulo", "", "Geral"), D("artigo", "São"), D("inciso", "a"), D("inciso", "b"),
                   D("paragrafo", "Exceção"), D("artigo", "Esta Lei entra em vigor na data de sua publicação")]) == []


def test_dispositivo_sem_pai():
    a = avisos([D("paragrafo", "solto"), D("artigo", "A"), D("alinea", "x"), D("alinea", "y"),
                D("artigo", "Esta Lei entra em vigor hoje")])
    assert any(x.startswith("Parágrafo único: precisa vir depois de um Artigo") for x in a)
    assert sum("precisa vir depois de um Inciso" in x for x in a) == 2


def test_lista_de_um_item_dois_pontos_sem_lista_vazio_e_vigencia():
    a = avisos([D("artigo", "São"), D("inciso", "único"), D("artigo", "Anuncia:"), D("artigo", "")])
    assert any("único item da lista" in x for x in a)
    assert any("termina em dois-pontos" in x for x in a)
    assert any("está sem texto" in x for x in a)
    assert any("cláusula de vigência" in x for x in a)
    assert avisos([D("capitulo", "", "Só título")]) == ["Projeto: ainda não tem nenhum Artigo."]
    assert avisos([]) == []


# ── documento Word ─────────────────────────────────────────────────────────

def test_docx_do_pl_e_valido_e_tem_o_conteudo():
    ds = [D("capitulo", "", "Disposições Gerais"), D("artigo", "Esta Lei institui o **Programa**"),
          D("inciso", "a"), D("inciso", "b"), D("artigo", "Esta Lei entra em vigor na data de sua publicação")]
    b = gerar_docx({**DADOS_PL, "dispositivos": ds})
    z = zipfile.ZipFile(io.BytesIO(b))
    assert z.testzip() is None
    for nome in ("word/styles.xml", "word/header1.xml", "word/footer1.xml", "[Content_Types].xml"):
        assert nome in z.namelist()         # estilos, cabeçalho e rodapé do modelo do gabinete
    xml = corpo_docx(b)
    ET.fromstring(xml)
    for trecho in ("PROJETO DE LEI Nº _____, DE 2026", "(Da Sra. Júlia Zanatta)", "CAPÍTULO I", "DISPOSIÇÕES GERAIS",
                   "Art. 1º Esta Lei institui o ", "II – b.", "JUSTIFICAÇÃO", "Segundo parágrafo.",
                   "Deputada JÚLIA ZANATTA", "PL/SC", "<w:sectPr"):
        assert trecho in xml, trecho
    assert "<w:b/></w:rPr><w:t xml:space=\"preserve\">Programa</w:t>" in xml     # **negrito**
    assert "Limita o crescimento dos encargos" not in xml                      # o texto de exemplo do modelo sai


def test_docx_aceita_caracteres_colados_de_outros_programas():
    ds = [D("artigo", "Linha um\nlinha dois\x0b\x0c & <tag> “aspas” aqui\x01")]
    b = gerar_docx({**DADOS_PL, "ementa": "Ementa\x0b com\x1f controle", "dispositivos": ds,
                    "justificativa": "Texto\x0c com \x08 lixo"})
    xml = corpo_docx(b)
    ET.fromstring(xml)                                # antes: XML inválido → Word não abria o arquivo
    assert "Art. 1º Linha um linha dois &amp; &lt;tag&gt; “aspas” aqui." in xml


def test_docx_com_texto_de_lei_alterada():
    ds = [D("artigo", "O art. 5º da Lei nº 1 passa a vigorar com a seguinte redação"),
          D("citacao", "“Art. 5º É vedado.\n§ 1º Exceção.” (NR)"), D("artigo", "Esta Lei entra em vigor na data de sua publicação")]
    xml = corpo_docx(gerar_docx({**DADOS_PL, "dispositivos": ds}))
    ET.fromstring(xml)
    assert ">“Art. 5º É vedado.</w:t>" in xml and ">§ 1º Exceção.” (NR)</w:t>" in xml
    assert "Art. 2º Esta Lei entra em vigor" in xml


# ── requerimento de urgência ───────────────────────────────────────────────

def test_requerimento_concorda_o_artigo_com_a_sigla():
    assert ementa_requerimento_urgencia("155", "PL", "2.548/2025").endswith("apreciação do PL 2.548/2025.")
    assert ementa_requerimento_urgencia("154", "PEC", "3/2025").endswith("apreciação da PEC 3/2025.")
    assert "da MPV 1.300/2025" in texto_requerimento_urgencia("155", "mpv", "1.300/2025", "Dispõe sobre X")
    assert ementa_requerimento_urgencia("155", "PL", "  ").endswith("do PL ____/____.")


def test_requerimento_nao_duplica_aspas_nem_ponto_da_ementa():
    t = texto_requerimento_urgencia("155", "PL", "1/2025", '  "Altera a Lei nº 1, e dá outras providências."  ')
    assert t.endswith('que "Altera a Lei nº 1, e dá outras providências".')


def test_docx_do_requerimento():
    b = gerar_docx_requerimento({
        "ano": "2026", "artigo_ricd": "155", "sigla_materia": "PEC", "numero_materia": "3/2025",
        "ementa_materia": "Altera o art. 5º da Constituição.", "autor_nome": "Júlia Zanatta",
        "autor_prefixo_assinatura": "Deputada Federal", "autor_partido_uf": "PL/SC", "fecho": "Sala das Sessões"})
    xml = corpo_docx(b)
    ET.fromstring(xml)
    for trecho in ("REQUERIMENTO Nº _____ DE 2026", "Senhor Presidente,", "urgência para apreciação da PEC 3/2025, que",
                   "“".replace("“", '"') + "Altera o art. 5º da Constituição" + '".', "Deputada Federal Júlia Zanatta", "(PL/SC)"):
        assert trecho in xml, trecho


if __name__ == "__main__":
    falhas = 0
    for nome, func in sorted(globals().items()):
        if nome.startswith("test_") and callable(func):
            try:
                func()
                print("OK   ", nome)
            except AssertionError as e:
                falhas += 1
                print("FALHA", nome, e)
    print("\n" + ("TODOS OS TESTES PASSARAM" if not falhas else f"{falhas} FALHA(S)"))
    sys.exit(1 if falhas else 0)
