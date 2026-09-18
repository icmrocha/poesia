# Poesias — revisão

Página: https://icmrocha.github.io/poesia/

Sistema para reler os 509 poemas, marcar o que já foi visto, favoritar, taguear e agrupar em compilações.

Há **três peças**. Elas não fazem a mesma coisa.

| Peça | Onde mora | Para que serve |
|---|---|---|
| Arquivos Word (`.docx`) | Google Drive, pasta **Poesias** | Texto original do poema. É o mestre da **escrita**. |
| `poesias.json` | Este repositório GitHub | Tags, Checado, Favorito, Compilação + uma cópia do texto. É o mestre da **revisão**. |
| `index.html` | Este repositório | A página que você abre no navegador. |

O GitHub Pages **só publica** o que está neste repo. Ele não lê o Drive.

---

## 1. Revisar no site (o dia a dia)

Não precisa de Python.

1. Abra https://icmrocha.github.io/poesia/
2. Espere aparecer **“JSON do site carregado.”** ao lado do botão Salvar.  
   Isso quer dizer que a página leu o `poesias.json` do GitHub, não uma cópia velha grudada no HTML.
3. Use busca, chips de Tags (azul) e Compilações (verde), e os botões:
   - **Esconder revisados** — some quem já tem Checado
   - **Ver favoritos** — só os com estrela
   - **Ver tags vazias** — só quem ainda não tem tag  
   Clique de novo no mesmo botão para desligar.
4. Em cada poema: Checado, Favorito, Tags, Compilação.
5. Clique **Salvar JSON**. O Firefox baixa `poesias.json` na pasta Downloads.
6. No GitHub, abra o repo `poesia` → o arquivo `poesias.json` → faça upload e **substitua** pelo arquivo baixado (mesmo nome).
7. Espere 1 a 2 minutos o Pages atualizar. Depois **Cmd+Shift+R** (recarregar sem cache).

Enquanto você não sobe o JSON, as marcas existem só naquele Firefox (`localStorage`). Outro computador, outro browser ou “limpar dados do site” apaga esse rascunho. Por isso o passo 6 importa.

---

## 2. Quando o texto do Word muda

Exemplo: você corrige um verso no Drive, ou acrescenta um poema novo `.docx`.

O site **não** olha o Google Drive. O `index.html` e o `poesias.json` só conhecem o texto da última vez que o script rodou. Por isso é preciso gerar de novo.

No Mac:

```bash
cd /Users/icmrocha/poesia
python3 gerar_poesias.py "/Users/icmrocha/Library/CloudStorage/GoogleDrive-ivancmrocha@gmail.com/Meu Drive/Poesias"
```

O clone neste Mac é `/Users/icmrocha/poesia` (aí estão o `.py` e o `poesias.json`).  
O segundo é a pasta **Poesias** do Drive, com os `.docx`.

O script:

1. Lê cada Word da raiz do Drive (não entra em subpasta).
2. Tira o título do negrito e o corpo (itálico vira itálico no HTML).
3. Abre o `poesias.json` que já estava no repo e **copia de volta** Checado / Favorito / Tags / Compilação, batendo pelo nome do arquivo (`Ivan Rocha - 4am.docx`).
4. Grava `poesias.json` e `index.html` novos.

Depois suba **os dois** arquivos no GitHub.

Se o JSON antigo não estiver na pasta de onde você rodou o comando, as tags nascem vazias. Sempre rode o `.py` de dentro do clone do repo.

Dependência: `pip install python-docx`

---

## 3. O que NÃO fazer

- Subir só o HTML e achar que o JSON novo entra. Sem o `poesias.json` ao lado, a página no ar não tem de onde puxar tags.
- Subir só o JSON e deixar um `index.html` **antigo**, de antes desta versão. O HTML velho não busca o JSON; ele traz os poemas colados por dentro.
- Colocar os 509 `.docx` no Git. Não precisa; o Drive já guarda o original.
- Editar o poema no HTML. Essa página não grava verso. Muda o Word e roda o script.

---

## 4. Tag vs compilação

São listas diferentes de propósito.

- **Tag** (chip azul) = assunto: `Família`, `Concreto`, `Série`…
- **Compilação** (chip verde) = livro ou capítulo: `Céu Aberto`, `Concreto` (o capítulo), `A Lua por trás dos Jasmins`…

Um poema pode ter a tag Concreto e **não** estar no capítulo Concreto (caso do 4am). O filtro do chip verde só olha o campo Compilação.

---

## 5. Se algo “voltar” ou sumir

O Firefox guarda uma sessão em `localStorage` com a chave `poesias-revisao-v5`.

Se a página brigar com o JSON do site:

1. Cmd+Opt+I (DevTools)
2. Storage → Local Storage → o endereço do site
3. Apague `poesias-revisao-v5`
4. Recarregue

Aí vale só o `poesias.json` que está no GitHub.

---

## 6. Arquivos neste repo

- `index.html` — a página (o Pages só usa esse nome, em minúsculas)
- `poesias.json` — dados da revisão
- `gerar_poesias.py` — gerador; comentado para manutenção
- `README.md` — este texto
