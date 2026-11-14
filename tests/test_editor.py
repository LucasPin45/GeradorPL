# -*- coding: utf-8 -*-
"""Testes das operações de edição (mover, excluir, inserir em bloco) e do rascunho."""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import editor  # noqa: E402


def lista(*tipos):
    """Lista de dispositivos em que o texto identifica cada um: "artigo#0", "paragrafo#1"…"""
    return [editor.novo_dispositivo(t, f"{t}#{i}", f"{t}#{i}") for i, t in enumerate(tipos)]


def nomes(dispositivos):
    return [d["texto"] for d in dispositivos]


BASE = ("capitulo", "artigo", "inciso", "inciso", "paragrafo", "artigo", "paragrafo", "inciso", "alinea", "artigo")
#        0           1         2         3         4            5         6            7         8         9


def test_bloco_e_dependentes():
    l = lista(*BASE)
    assert editor.dependentes(l, 1) == 3      # Art. leva incisos e parágrafo
    assert editor.dependentes(l, 5) == 3
    assert editor.dependentes(l, 6) == 2      # parágrafo leva inciso e alínea
    assert editor.dependentes(l, 0) == 0      # o capítulo anda sozinho
    assert editor.dependentes(l, 9) == 0


def test_subir_leva_os_dependentes():
    l = lista(*BASE)
    r = editor.subir(l, 5)                    # 2º artigo sobe acima do 1º, com tudo que é dele
    assert nomes(r) == ["capitulo#0", "artigo#5", "paragrafo#6", "inciso#7", "alinea#8",
                        "artigo#1", "inciso#2", "inciso#3", "paragrafo#4", "artigo#9"]
    assert len(r) == len(l) and {d["uid"] for d in r} == {d["uid"] for d in l}


def test_descer_leva_os_dependentes():
    l = lista(*BASE)
    r = editor.descer(l, 1)
    assert nomes(r) == ["capitulo#0", "artigo#5", "paragrafo#6", "inciso#7", "alinea#8",
                        "artigo#1", "inciso#2", "inciso#3", "paragrafo#4", "artigo#9"]


def test_subordinado_nao_sai_do_dispositivo_a_que_pertence():
    l = lista(*BASE)
    assert not editor.pode_subir(l, 2)        # 1º inciso do artigo
    assert editor.pode_subir(l, 3) and nomes(editor.subir(l, 3))[2:4] == ["inciso#3", "inciso#2"]
    assert not editor.pode_descer(l, 3)       # depois dele vem um parágrafo, não outro inciso
    assert not editor.pode_subir(l, 4)        # parágrafo não passa para cima dos incisos do caput
    assert not editor.pode_descer(l, 4)       # nem para dentro do artigo seguinte
    assert editor.subir(l, 2) == l and editor.descer(l, 4) == l


def test_capitulo_e_artigo_trocam_de_lugar():
    l = lista(*BASE)
    assert not editor.pode_subir(l, 0)
    r = editor.descer(l, 0)                   # o capítulo desce para depois do 1º artigo inteiro
    assert nomes(r)[:5] == ["artigo#1", "inciso#2", "inciso#3", "paragrafo#4", "capitulo#0"]
    assert nomes(editor.subir(r, 4)) == nomes(l)
    assert not editor.pode_descer(l, 9)


def test_excluir_remove_o_bloco():
    l = lista(*BASE)
    assert nomes(editor.excluir(l, 5)) == ["capitulo#0", "artigo#1", "inciso#2", "inciso#3", "paragrafo#4", "artigo#9"]
    assert nomes(editor.excluir(l, 0)) == nomes(l)[1:]
    assert nomes(editor.excluir(l, 8)) == nomes(l)[:8] + nomes(l)[9:]


def test_inserir_no_fim_e_depois_de_um_bloco():
    l = lista(*BASE)
    novo = editor.novo_dispositivo("artigo", "NOVO")
    assert nomes(editor.inserir(l, novo))[-1] == "NOVO"
    # depois do 1º artigo = depois dos incisos e do parágrafo dele
    assert nomes(editor.inserir(l, novo, l[1]["uid"]))[4:7] == ["paragrafo#4", "NOVO", "artigo#5"]
    assert nomes(editor.inserir(l, novo, "uid-que-nao-existe"))[-1] == "NOVO"


def test_inserir_subordinado_respeita_a_ordem_caput_incisos_paragrafos():
    l = lista(*BASE)
    inciso = editor.novo_dispositivo("inciso", "INC")
    paragrafo = editor.novo_dispositivo("paragrafo", "PAR")
    # novo inciso do caput: depois dos incisos existentes, antes do parágrafo
    assert nomes(editor.inserir(l, inciso, l[1]["uid"]))[1:6] == ["artigo#1", "inciso#2", "inciso#3", "INC", "paragrafo#4"]
    # novo parágrafo do artigo: no fim do bloco
    assert nomes(editor.inserir(l, paragrafo, l[1]["uid"]))[4:7] == ["paragrafo#4", "PAR", "artigo#5"]
    # inciso depois de um inciso: logo abaixo dele
    assert nomes(editor.inserir(l, inciso, l[2]["uid"]))[2:5] == ["inciso#2", "INC", "inciso#3"]


def test_proximo_tipo_sugerido():
    l = lista(*BASE)
    assert editor.proximo_tipo_sugerido([]) == "artigo"
    assert editor.proximo_tipo_sugerido(l) == "artigo"
    assert editor.proximo_tipo_sugerido(l, l[0]["uid"]) == "artigo"       # depois de um capítulo
    assert editor.proximo_tipo_sugerido(l, l[7]["uid"]) == "inciso"
    assert editor.proximo_tipo_sugerido(lista("artigo", "citacao")) == "artigo"


def test_rascunho_ida_e_volta():
    l = lista(*BASE)
    l[4]["paragrafo_unico_forcado"] = True
    l[2]["pontuacao_automatica"] = False
    campos = {"ano": "2026", "autor_nome": "Júlia Zanatta", "ementa": "Institui “X” & <y>", "justificativa": "A\n\nB",
              "tipo_documento": "Projeto de Lei (PL)", "numero_materia": "2.548/2025"}
    texto = editor.exportar_rascunho(campos, l)
    campos2, l2 = editor.importar_rascunho(texto)
    assert l2 == l
    assert all(campos2[k] == v for k, v in campos.items())
    assert campos2["fecho_req"] == ""          # campo não informado volta vazio, não some


def test_rascunho_ignora_o_que_nao_reconhece():
    bruto = json.dumps({"formato": "gerador-pl-rascunho", "versao": 1,
                        "campos": {"ano": 2027, "campo_estranho": "x", "ementa": ["lista"]},
                        "dispositivos": [{"tipo": "artigo", "texto": "ok", "uid": "a"}, {"tipo": "artigo", "texto": "dup", "uid": "a"},
                                         {"tipo": "invalido", "texto": "x"}, "lixo", {"tipo": "paragrafo", "paragrafo_unico_forcado": "sim"}]})
    campos, l = editor.importar_rascunho(bruto)
    assert campos == {"ano": "2027"}
    assert [d["tipo"] for d in l] == ["artigo", "artigo", "paragrafo"]
    assert len({d["uid"] for d in l}) == 3                   # uid repetido ou ausente é refeito
    assert l[2]["paragrafo_unico_forcado"] is None and l[2]["pontuacao_automatica"] is True


def test_rascunho_recusa_arquivo_que_nao_e_rascunho():
    for ruim in ("não é json", "[]", '{"formato": "outro"}', ""):
        try:
            editor.importar_rascunho(ruim)
        except ValueError:
            continue
        raise AssertionError(f"deveria recusar: {ruim!r}")


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
