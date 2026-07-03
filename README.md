# Gerador de Documentos Legislativos — Gabinete Deputada Júlia Zanatta

Sistema em Streamlit com dois módulos, selecionáveis no topo da tela:

1. **Projeto de Lei (PL)** — monta o PL dispositivo por dispositivo e gera o
   `.docx` no padrão do gabinete, com numeração e pontuação de acordo com a
   Lei Complementar nº 95/1998 (LC 95/98).
2. **Requerimento de Urgência** — gera o requerimento de urgência (art. 154
   ou 155 do RICD) para apreciação de uma matéria, preenchendo só os campos
   variáveis (artigo, sigla, número/ano da matéria e ementa).

## Módulo Projeto de Lei — como funciona

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

## Módulo Requerimento de Urgência — como funciona

1. Escolha **Requerimento de Urgência** no seletor do topo.
2. Selecione o fundamento regimental (**art. 154** ou **155** do RICD) —
   confirme com a assessoria qual se aplica ao caso concreto.
3. Escolha a sigla da matéria (PL, PLP, PEC, PDL...) e informe o número/ano
   (ex.: `2.548/2025`).
4. Cole a ementa oficial da matéria — ela entra automaticamente entre aspas
   no texto final, sem precisar digitar as aspas.
5. Clique em **Gerar Word (.docx)**.

O texto final segue exatamente o modelo do gabinete:

> Requeremos, nos termos do art. \_\_\_ do Regimento Interno da Câmara dos
> Deputados, urgência para apreciação do \_\_\_ \_\_\_, que "\_\_\_".

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
