# Gerador de PL — Gabinete Deputada Júlia Zanatta

Sistema em Streamlit para montar Projetos de Lei dispositivo por dispositivo
e gerar automaticamente o `.docx` no padrão do gabinete, com numeração e
pontuação de acordo com a Lei Complementar nº 95/1998 (LC 95/98).

## Como funciona

1. Preencha os **dados gerais** na barra lateral (ano, autor, partido/UF).
2. Escreva a **ementa**.
3. Adicione os **dispositivos** um a um, na ordem do texto: Título, Capítulo,
   Seção, Subseção, Artigo, Parágrafo, Inciso, Alínea, Item. Você digita só o
   texto — a numeração ("Art. 1º", "§ 2º", "I –", "a)", "1.") e a pontuação
   final (";", "; e", ":", ".") são calculadas automaticamente, seguindo os
   arts. 10 a 12 da LC 95/98:
   - Artigos numerados de forma ordinal até o 9º e cardinal a partir do 10º
     (ex.: "Art. 9º", depois "Art. 10.").
   - A mesma regra vale para parágrafos; se houver só um no artigo, vira
     "Parágrafo único." automaticamente.
   - Incisos em algarismos romanos, alíneas em letras minúsculas, itens em
     algarismos arábicos.
   - Use `**texto**` dentro de qualquer campo para deixar um trecho em
     negrito (útil na ementa, ao citar a lei alterada).
4. Use as setas ↑ ↓ para reordenar e a lixeira para remover.
5. Escreva a **justificação** (parágrafos separados por linha em branco).
6. Clique em **Gerar Word (.docx)** e baixe o arquivo.

## Rodando localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy no Streamlit Community Cloud

1. Suba esta pasta (`app.py`, `gerador.py`, `requirements.txt`,
   `assets/template_gabinete.docx`) para um repositório no GitHub.
2. Em https://share.streamlit.io, clique em "New app", aponte para o
   repositório e selecione `app.py` como arquivo principal.
3. Não é necessário nenhum "secret" ou variável de ambiente — o modelo
   `.docx` do gabinete já vai embutido no repositório
   (`assets/template_gabinete.docx`).

## Estrutura de arquivos

```
app.py                          # interface Streamlit
gerador.py                      # numeração LC 95/98 + geração do OOXML
assets/template_gabinete.docx   # modelo com estilos, cabeçalho e rodapé do gabinete
requirements.txt
```

## Trocar o modelo (template)

Se o padrão de formatação do gabinete mudar, basta substituir
`assets/template_gabinete.docx` por um novo modelo com a mesma estrutura
(estilos `Normal`/`SemEspaamento`, cabeçalho `rId8`, rodapé `rId9`). O
gerador reaproveita o `sectPr`, o cabeçalho e o rodapé do arquivo que
estiver nesse caminho — não é necessário alterar `gerador.py`, a menos que
os nomes dos estilos ou os recuos (em `w:ind`) mudem.
