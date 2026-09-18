#!/usr/bin/env python3
"""
gerar_poesias.py
================
Lê os arquivos Word (.docx) da pasta de poemas e gera dois arquivos:

  poesias.json   → dados (texto, tags, checado, favorito, compilação)
  index.html     → página de revisão (GitHub Pages / navegador)

O HTML no GitHub Pages CARREGA o poesias.json pela internet. Por isso,
depois de só mudar tags, basta subir o JSON. Este script só precisa
rodar de novo quando o TEXTO do poema mudou no Word.

Uso (na pasta onde estão o .py e o poesias.json):
  pip install python-docx
  python3 gerar_poesias.py "/caminho/da/pasta/com/os/docx"

O script junta o JSON antigo (tags etc.) com o texto novo dos .docx,
pelo nome do arquivo (ex.: "Ivan Rocha - 4am.docx").
"""

from __future__ import annotations

import json
import re
import sys
import unicodedata
from pathlib import Path

try:
    from docx import Document
except ImportError:
    sys.exit("Instale com:  pip install python-docx")


def title_from_name(name: str) -> str:
    """Se o Word não tiver título em negrito, usa o nome do arquivo."""
    n = re.sub(r"\.docx$", "", name, flags=re.I)
    n = re.sub(r"^Ivan Rocha\s*-\s*", "", n, flags=re.I)
    return n.strip()


def paragraph_is_bold(p) -> bool:
    """True se o parágrafo inteiro está em negrito ou é estilo Título."""
    runs = [r for r in p.runs if (r.text or "").strip()]
    if runs and all(r.bold for r in runs):
        return True
    style = (p.style.name if p.style is not None else "") or ""
    return style.lower().startswith("heading") or "título" in style.lower()


def extract_text(path: Path) -> str:
    """Texto puro do Word, sem formatação. Vai para o campo 'text' do JSON."""
    doc = Document(str(path))
    lines = [p.text.replace("\xa0", " ").rstrip() for p in doc.paragraphs]
    while lines and not lines[-1].strip():
        lines.pop()
    while lines and not lines[0].strip():
        lines.pop(0)
    return "\n".join(lines)


def runs_html(p) -> str:
    """Converte um parágrafo do Word em HTML simples (itálico vira <em>)."""
    parts = []
    for r in p.runs:
        t = html_escape((r.text or "").replace("\xa0", " "))
        if not t:
            continue
        if r.italic:
            t = f"<em>{t}</em>"
        parts.append(t)
    return "".join(parts).rstrip()


def extract_poem(path: Path) -> tuple[str, str, str]:
    """
    Lê um .docx e devolve (titulo, corpo_texto, corpo_html).

    Título = primeiro trecho em negrito (não a linha "Ivan Rocha").
    Se na mesma linha vier "*para fulano" em itálico, isso vai para o corpo.
    """
    doc = Document(str(path))
    title = ""
    html_lines = []
    text_lines = []
    for p in doc.paragraphs:
        text = p.text.replace("\xa0", " ").rstrip()
        s = text.strip()
        if AUTHOR_LINE.match(s):
            continue
        if not title and s:
            bold_bits = []
            rest_started = False
            rest_html = []
            rest_txt = []
            for r in p.runs:
                rt = (r.text or "").replace("\xa0", " ")
                if not rest_started and r.bold and rt.strip():
                    bold_bits.append(rt)
                    continue
                if bold_bits and not rest_started:
                    rest_started = True
                if rest_started or (not r.bold and rt.strip()):
                    rest_started = True
                    chunk = html_escape(rt)
                    if r.italic and chunk:
                        chunk = f"<em>{chunk}</em>"
                    rest_html.append(chunk)
                    rest_txt.append(rt)
            if bold_bits:
                title = "".join(bold_bits).strip()
                # se o negrito pegou a dedicatória também, separa
                if title.casefold().startswith(title_from_name(path.name).casefold()):
                    pass
                rh = "".join(rest_html).rstrip()
                rt = "".join(rest_txt).rstrip()
                if rh.strip():
                    html_lines.append(rh)
                    text_lines.append(rt)
                continue
            if paragraph_is_bold(p) and not AUTHOR_LINE.match(s):
                # parágrafo inteiro em negrito: título pode vir com dedicatória
                m = re.match(r"^(.*?)(\s+\*.*)$", s)
                if m:
                    title = m.group(1).strip()
                    ded = m.group(2).strip()
                    html_lines.append(f"<em>{html_escape(ded)}</em>")
                    text_lines.append(ded)
                else:
                    title = s
                continue
        html_lines.append(runs_html(p) if p.runs else html_escape(text))
        text_lines.append(text)
    if not title:
        title = title_from_name(path.name)
    raw = "\n".join(text_lines)
    body = body_only(raw, title)
    body_html = body_only_html("\n".join(html_lines), title)
    return title, body, body_html


# Linha que é só o nome do autor — não entra no corpo do poema.
AUTHOR_LINE = re.compile(
    r"^\s*[\(\[\{]?\s*Ivan\s+Rocha\s*[\)\]\}]?\s*$",
    re.I,
)


def body_only(text: str, title: str) -> str:
    """Tira título repetido no topo e qualquer linha só com (Ivan Rocha)."""
    lines = text.splitlines()
    cleaned = []
    title_k = sort_key(title)
    skipped_title = False
    for line in lines:
        s = line.strip()
        sk = sort_key(s)
        if not skipped_title and (sk == title_k or sk.startswith(title_k)):
            skipped_title = True
            rest = s
            first = s.split(None, 1)
            if first and sort_key(first[0]) == title_k:
                rest = first[1] if len(first) > 1 else ""
            elif sk.startswith(title_k):
                rest = s[len(title):].lstrip() if s.casefold().startswith(title.casefold()) else s
                if rest == s and first:
                    rest = first[1] if len(first) > 1 else ""
            if rest:
                cleaned.append(rest)
            continue
        if AUTHOR_LINE.match(s):
            continue
        cleaned.append(line.rstrip())
    while cleaned and not cleaned[0].strip():
        cleaned.pop(0)
    while cleaned and not cleaned[-1].strip():
        cleaned.pop()
    return "\n".join(cleaned)


def sort_key(s: str) -> str:
    n = unicodedata.normalize("NFD", s or "")
    n = "".join(ch for ch in n if unicodedata.category(ch) != "Mn")
    return n.casefold()


def body_only_html(html: str, title: str) -> str:
    lines = html.splitlines()
    cleaned = []
    title_cf = title.casefold()
    skipped = False
    for line in lines:
        plain = re.sub(r"<[^>]+>", "", line).strip()
        if not skipped and plain.casefold().startswith(title_cf):
            skipped = True
            rest = line
            # tira o título do começo da linha HTML
            idx = plain.casefold().find(title_cf)
            if idx == 0:
                # remove first title-length visible chars — fallback: se a linha ficou só título, drop
                if plain.casefold() == title_cf:
                    continue
            m = re.match(r"^(?:<[^>]+>)*" + re.escape(title) + r"\s*", line, re.I)
            if m:
                rest = line[m.end() :]
            if rest.strip():
                if rest.strip().startswith("*") and "<em>" not in rest:
                    rest = f"<em>{rest.strip()}</em>"
                cleaned.append(rest.rstrip())
            continue
        if AUTHOR_LINE.match(plain):
            continue
        cleaned.append(line.rstrip())
    while cleaned and not cleaned[0].strip():
        cleaned.pop(0)
    while cleaned and not cleaned[-1].strip():
        cleaned.pop()
    return "\n".join(cleaned)


def html_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def build_html(poems: list[dict]) -> str:
    """Monta o index.html inteiro (CSS + lista + poemas + JavaScript)."""
    poems = sorted(poems, key=lambda p: sort_key(p["title"]))
    index_items = list(enumerate(poems))
    index_html = []
    for i, p in index_items:
        pid = html_escape(p["file_name"])
        title = html_escape(p["title"])
        rev = "checked" if p.get("reviewed") else ""
        fav = "★" if p.get("favorite") else ""
        index_html.append(
            f'<li data-id="{pid}">'
            f'<input type="checkbox" class="ck-index" data-id="{pid}" {rev}> '
            f'<a href="#p{i}">{title}</a> <span class="star">{fav}</span></li>'
        )

    poems_html = []
    for i, p in enumerate(poems):
        pid = html_escape(p["file_name"])
        title = p["title"]
        raw = p.get("text") or p.get("body") or ""
        body_plain = body_only(raw, title)
        body = p.get("body_html") or html_escape(body_plain)
        if body == html_escape(body_plain):
            # dedicatória *para… em itálico mesmo no texto puro
            lined = []
            for ln in body.split("\n"):
                raw_ln = body_plain.split("\n")[len(lined)] if len(lined) < len(body_plain.split("\n")) else ""
                if re.match(r"\s*\*.+", raw_ln):
                    lined.append(f"<em>{ln}</em>")
                else:
                    lined.append(ln)
            body = "\n".join(lined)
        title = html_escape(title)
        def split_comp(val):
            items = val if isinstance(val, list) else ([val] if val else [])
            out = []
            for c in items:
                if not c:
                    continue
                if " / " in str(c):
                    out.extend(x.strip() for x in str(c).split(" / ") if x.strip())
                else:
                    out.append(str(c).strip())
            seen = set()
            uniq = []
            for x in out:
                k = sort_key(x)
                if k not in seen:
                    seen.add(k)
                    uniq.append(x)
            return uniq

        def join_field(val):
            if isinstance(val, list):
                return ", ".join(val)
            return val or ""

        tags_s = html_escape(join_field(p.get("tags")))
        comps = split_comp(p.get("compilacoes") or p.get("compilacao"))
        p["compilacoes"] = comps
        comp_s = html_escape(join_field(comps))
        note_s = html_escape(p.get("comentario") or "")
        rev = "checked" if p.get("reviewed") else ""
        fav = "checked" if p.get("favorite") else ""
        poems_html.append(
            f'<article class="poem" id="p{i}" data-id="{pid}">\n'
            f'<h1>{title}</h1>\n'
            f'<div class="verse">{body}</div>\n'
            f'<div class="meta">\n'
            f'<label><input type="checkbox" class="ck-poem" data-id="{pid}" {rev}> Checado</label>\n'
            f'<label><input type="checkbox" class="ck-fav" data-id="{pid}" {fav}> Favorito</label>\n'
            f'<label>Tags <input type="text" class="tags" data-id="{pid}" value="{tags_s}"></label>\n'
            f'<label>Compilação <input type="text" class="comp" data-id="{pid}" value="{comp_s}"></label>\n'
            f'<label>Comentário <input type="text" class="note" data-id="{pid}" value="{note_s}"></label>\n'
            f'</div>\n'
            f'<a class="back" href="#indice">← voltar ao índice</a>\n'
            f'</article>\n'
        )
        if i < len(poems) - 1:
            poems_html.append('<div class="sep" aria-hidden="true"></div>\n')

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Poesias — revisão</title>
<style>
  @page {{ size: A4; margin: 12mm; }}
  * {{ box-sizing: border-box; }}
  html, body {{
    margin: 0;
    padding: 0;
    font: 11pt/1.35 Helvetica, Arial, sans-serif;
    color: #000;
    background: #fff;
  }}
  header.bar {{
    position: sticky;
    top: 0;
    background: #fff;
    border-bottom: 1px solid #ccc;
    padding: 10px 0 12px 0;
    margin: 0 0 16px 0;
    z-index: 2;
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 28px;
  }}
  #progress {{ font-weight: 700; }}
  .brand-name {{ font-weight: 700; line-height: 1.2; }}
  .brand-sub {{ font-weight: 400; font-size: 10pt; color: #444; }}
  .bar-status {{
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 2px;
    margin-left: 8px;
  }}
  footer.site {{
    margin-top: 8px;
    font-size: 9pt;
    color: #444;
  }}
  nav#indice {{
    scroll-margin-top: 88px;
  }}
  nav#indice h2 {{
    margin: 22px 0 12px 0;
  }}
  .busca-wrap {{
    display: block;
    margin: 0 0 8px 0;
    font-size: 10pt;
  }}
  #busca {{
    display: block;
    width: 100%;
    max-width: 28em;
    margin-top: 4px;
    font: 11pt Helvetica, Arial, sans-serif;
    padding: 6px 8px;
  }}
  #busca-status {{
    margin: 0 0 8px 0;
    font-size: 10pt;
    color: #444;
  }}
  .chip-row {{
    margin: 0 0 8px 0;
  }}
  .chip-row .lbl {{
    font-size: 9pt;
    font-weight: 700;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    color: #555;
    margin: 0 0 4px 0;
  }}
  .chips {{
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
  }}
  .chips button {{
    font: 10pt Helvetica, Arial, sans-serif;
    padding: 3px 8px;
    border: 1px solid #bbb;
    background: #fff;
    border-radius: 999px;
    cursor: pointer;
  }}
  .chips button.on {{
    color: #fff;
    border-color: transparent;
  }}
  #chips-tags button.on {{ background: #2c5aa0; }}
  #chips-comps button.on {{ background: #2a7a4b; }}
  #chips-tags button {{ border-color: #8aa4cc; }}
  #chips-comps button {{ border-color: #8bbb9a; }}
  .is-hidden {{ display: none !important; }}
  ul.indice {{
    columns: 2;
    column-gap: 24px;
    list-style: none;
    padding: 0;
    margin: 0 0 24px 0;
  }}
  ul.indice li {{
    break-inside: avoid;
    margin: 0 0 4px 0;
  }}
  ul.indice a {{ color: #000; text-decoration: none; }}
  ul.indice a:hover {{ text-decoration: underline; }}
  ul.indice li.done a {{ color: #666; }}
  .poem {{
    break-inside: avoid;
    page-break-inside: avoid;
    margin: 0;
    padding: 0 0 0.4em 0;
    scroll-margin-top: 88px;
  }}
  .poem.done h1 {{ color: #444; }}
  .sep {{
    margin: 1.2em 0 1.4em 0;
    border: 0;
    border-top: 1px dotted #444;
    height: 0;
  }}
  .back {{
    display: inline-block;
    margin-top: 14px;
    color: #000;
    font-size: 10pt;
  }}
  .poem h1 {{
    font-size: 12pt;
    font-weight: 700;
    margin: 0 0 0.85em 0;
  }}
  .poem h1 label {{ cursor: pointer; }}
  .poem .verse {{
    margin: 0;
    font: inherit;
    white-space: pre-wrap;
    word-wrap: break-word;
  }}
  .poem .verse em {{
    font-style: italic;
  }}
  .poem pre {{
    margin: 0;
    font: inherit;
    white-space: pre-wrap;
    word-wrap: break-word;
  }}
  .meta {{
    margin: 2.4em 0 0 0;
    padding: 12px 14px;
    font-size: 10pt;
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    gap: 6px;
    background: #f3f3f3;
    border-radius: 10px;
    max-width: 28em;
  }}
  .meta label {{
    display: block;
  }}
  .meta .tags, .meta .comp {{
    width: 16em;
    font: 10pt Helvetica, Arial, sans-serif;
    margin-left: 6px;
  }}
  html {{ scroll-behavior: smooth; }}
  header.bar > div:last-child {{
    display: flex;
    flex-wrap: wrap;
    justify-content: flex-end;
    align-items: center;
    gap: 8px;
  }}
  header.bar button {{
    font: 10pt Helvetica, Arial, sans-serif;
  }}
  header.bar button.on {{
    background: #333;
    color: #fff;
    border-color: #333;
  }}
  #save-status {{ color: #444; margin-left: 8px; }}
  @media screen {{
    body {{ max-width: 210mm; margin: 16px auto; padding: 12mm; }}
  }}
  @media print {{
    header.bar {{ position: static; border: 0; }}
    nav#indice, .meta, .back, header.bar button, #save-status, .busca-wrap {{ display: none; }}
  }}
</style>
</head>
<body>
<header class="bar">
  <div class="brand">
    <div class="brand-name">Ivan Rocha</div>
    <div class="brand-sub">Poesias</div>
  </div>
  <div>
    <button type="button" id="btn-hide-rev">Esconder revisados</button>
    <button type="button" id="btn-only-fav">Ver favoritos</button>
    <button type="button" id="btn-empty-tags">Ver tags vazias</button>
    <button type="button" id="btn-save">Salvar JSON</button>
    <div class="bar-status">
      <span id="save-status"></span>
      <div id="progress">Revisados: 0 / {len(poems)}</div>
    </div>
  </div>
</header>
<nav id="indice">
  <label class="busca-wrap">
    <input type="search" id="busca" placeholder="Busca por título, tag, compilação ou comentário" autocomplete="off">
  </label>
  <p id="busca-status"></p>
  <div class="chip-row">
    <div class="lbl">Tags</div>
    <div class="chips" id="chips-tags"></div>
  </div>
  <div class="chip-row">
    <div class="lbl">Compilações</div>
    <div class="chips" id="chips-comps"></div>
  </div>
  <h2>Índice</h2>
  <ul class="indice">
    {''.join(index_html)}
  </ul>
</nav>
<div class="sep" aria-hidden="true"></div>
{''.join(poems_html)}
<div class="sep" aria-hidden="true"></div>
<footer class="site">Todos os textos © Ivan César Martins Rocha</footer>
<script>
// Cópia local no Firefox (rascunho). Se mudar o nome, a sessão antiga é ignorada.
const KEY = "poesias-revisao-v6";
const total = {len(poems)};
let BASE = {json.dumps([{"file_name": p["file_name"], "title": p["title"], "text": p.get("text") or "", "body": p.get("body") or "", "reviewed": bool(p.get("reviewed")), "favorite": bool(p.get("favorite")), "tags": p.get("tags") or [], "compilacoes": p.get("compilacoes") or p.get("compilacao") or [], "comentario": p.get("comentario") or ""} for p in poems], ensure_ascii=False)};
let fileHandle = null;

function joinList(val) {{
  if (Array.isArray(val)) return val.join(", ");
  return val || "";
}}
function seedFromBase() {{
  const s = {{}};
  BASE.forEach(p => {{
    s[p.file_name] = {{
      reviewed: !!p.reviewed,
      favorite: !!p.favorite,
      tags: joinList(p.tags),
      compilacoes: joinList(p.compilacoes),
      comentario: p.comentario || ""
    }};
  }});
  return s;
}}
function loadMeta() {{
  let parsed = {{}};
  try {{
    const raw = localStorage.getItem(KEY);
    if (raw) parsed = JSON.parse(raw) || {{}};
  }} catch (e) {{
    parsed = {{}};
  }}
  const base = seedFromBase();
  Object.keys(base).forEach(id => {{
    const b = base[id];
    const p = parsed[id] || {{}};
    parsed[id] = {{
      reviewed: !!(p.reviewed || b.reviewed),
      favorite: !!(p.favorite || b.favorite),
      tags: ("tags" in p) ? p.tags : (b.tags || ""),
      compilacoes: ("compilacoes" in p) ? p.compilacoes : (b.compilacoes || ""),
      comentario: ("comentario" in p) ? p.comentario : (b.comentario || "")
    }};
  }});
  return parsed;
}}
function saveMeta(state) {{
  localStorage.setItem(KEY, JSON.stringify(state));
}}
function metaOf(id) {{
  const s = loadMeta();
  return Object.assign({{reviewed:false, favorite:false, tags:"", compilacoes:"", comentario:""}}, s[id] || {{}});
}}
function setMeta(id, patch) {{
  const s = loadMeta();
  s[id] = Object.assign(metaOf(id), patch);
  saveMeta(s);
  apply();
}}
function apply() {{
  const s = loadMeta();
  let n = 0, favs = 0;
  document.querySelectorAll("article.poem").forEach(art => {{
    const id = art.getAttribute("data-id");
    const m = Object.assign({{reviewed:false, favorite:false, tags:"", compilacoes:"", comentario:""}}, s[id] || {{}});
    const ck = art.querySelector(".ck-poem");
    const fv = art.querySelector(".ck-fav");
    const tg = art.querySelector(".tags");
    const cp = art.querySelector(".comp");
    const nt = art.querySelector(".note");
    if (ck) ck.checked = !!m.reviewed;
    if (fv) fv.checked = !!m.favorite;
    if (tg && document.activeElement !== tg) tg.value = m.tags || "";
    if (cp && document.activeElement !== cp) cp.value = m.compilacoes || "";
    if (nt && document.activeElement !== nt) nt.value = m.comentario || "";
    art.classList.toggle("done", !!m.reviewed);
    if (m.reviewed) n++;
    if (m.favorite) favs++;
  }});
  document.querySelectorAll("#indice li").forEach(li => {{
    const id = li.getAttribute("data-id");
    const m = Object.assign({{reviewed:false, favorite:false}}, s[id] || {{}});
    const ck = li.querySelector(".ck-index");
    if (ck) ck.checked = !!m.reviewed;
    li.classList.toggle("done", !!m.reviewed);
    const star = li.querySelector(".star");
    if (star) star.textContent = m.favorite ? "★" : "";
  }});
  document.getElementById("progress").textContent =
    "Revisados: " + n + " / " + total + " · Favoritos: " + favs;
}}
function buildExport() {{
  const s = loadMeta();
  const poems = BASE.map(p => {{
    const m = Object.assign({{reviewed:false, favorite:false, tags:"", compilacoes:"", comentario:""}}, s[p.file_name] || {{}});
    const split = (v) => (v || "").split(",").map(t => t.trim()).filter(Boolean);
    return Object.assign({{}}, p, {{
      reviewed: !!m.reviewed,
      favorite: !!m.favorite,
      tags: split(m.tags),
      compilacoes: split(m.compilacoes),
      comentario: m.comentario || ""
    }});
  }});
  return {{
    count: poems.length,
    poems: poems
  }};
}}
function downloadJson() {{
  const blob = new Blob([JSON.stringify(buildExport(), null, 2)], {{type:"application/json"}});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "poesias.json";
  a.click();
  URL.revokeObjectURL(a.href);
  document.getElementById("save-status").textContent = "JSON baixado.";
}}
async function saveToServer() {{
  const status = document.getElementById("save-status");
  const payload = JSON.stringify(buildExport(), null, 2);
  if (location.protocol === "file:") {{
    downloadJson();
    return;
  }}
  try {{
    const res = await fetch("/salvar-json", {{
      method: "POST",
      headers: {{"Content-Type": "application/json"}},
      body: payload
    }});
    if (!res.ok) throw new Error(await res.text());
    const info = await res.json();
    status.textContent = "Salvo em " + (info.path || "poesias.json");
  }} catch (err) {{
    status.textContent = "Falha ao gravar; baixando. " + err;
    downloadJson();
  }}
}}
async function linkJson() {{
  if (!window.showOpenFilePicker) {{
    document.getElementById("save-status").textContent =
      "Este navegador não deixa gravar file://. Use Chrome/Edge ou Baixar JSON.";
    return;
  }}
  const handles = await showOpenFilePicker({{
    types: [{{description:"JSON", accept: {{"application/json":[".json"]}}}}]
  }});
  fileHandle = handles[0];
  document.getElementById("save-status").textContent = "Vinculado: " + fileHandle.name;
}}
async function saveJson() {{
  const data = JSON.stringify(buildExport(), null, 2);
  if (fileHandle && fileHandle.createWritable) {{
    const w = await fileHandle.createWritable();
    await w.write(data);
    await w.close();
    document.getElementById("save-status").textContent = "Salvo no JSON.";
    return;
  }}
  downloadJson();
}}
document.addEventListener("change", (ev) => {{
  const t = ev.target;
  const id = t.getAttribute && t.getAttribute("data-id");
  if (!id) return;
  if (t.matches(".ck-poem, .ck-index")) setMeta(id, {{reviewed: t.checked}});
  if (t.matches(".ck-fav")) setMeta(id, {{favorite: t.checked}});
  if (t.matches(".tags")) setMeta(id, {{tags: t.value}});
  if (t.matches(".comp")) setMeta(id, {{compilacoes: t.value}});
  if (t.matches(".note")) setMeta(id, {{comentario: t.value}});
}});
document.addEventListener("input", (ev) => {{
  const t = ev.target;
  if (t.matches(".tags")) setMeta(t.getAttribute("data-id"), {{tags: t.value}});
  if (t.matches(".comp")) setMeta(t.getAttribute("data-id"), {{compilacoes: t.value}});
  if (t.matches(".note")) setMeta(t.getAttribute("data-id"), {{comentario: t.value}});
}});
document.getElementById("btn-save").onclick = () => saveToServer();
let hideReviewed = false;
let onlyFav = false;
let onlyEmptyTags = false;
function toggleBtn(btn, on) {{
  if (btn) btn.classList.toggle("on", on);
}}
document.getElementById("btn-hide-rev").onclick = () => {{
  hideReviewed = !hideReviewed;
  toggleBtn(document.getElementById("btn-hide-rev"), hideReviewed);
  runSearch();
}};
document.getElementById("btn-only-fav").onclick = () => {{
  onlyFav = !onlyFav;
  toggleBtn(document.getElementById("btn-only-fav"), onlyFav);
  runSearch();
}};
document.getElementById("btn-empty-tags").onclick = () => {{
  onlyEmptyTags = !onlyEmptyTags;
  toggleBtn(document.getElementById("btn-empty-tags"), onlyEmptyTags);
  runSearch();
}};
const activeTags = new Set();
const activeComps = new Set();
function fold(s) {{
  return (s || "").normalize("NFD").replace(/\\p{{M}}/gu, "").toLowerCase();
}}
function splitVals(v) {{
  return (v || "").split(",").map(t => t.trim()).filter(Boolean);
}}
function haystack(art) {{
  const id = art.getAttribute("data-id");
  const title = (art.querySelector("h1") || {{}}).textContent || "";
  const tags = (art.querySelector(".tags") || {{}}).value || "";
  const comp = (art.querySelector(".comp") || {{}}).value || "";
  const note = (art.querySelector(".note") || {{}}).value || "";
  return fold(title + " " + tags + " " + comp + " " + note + " " + id);
}}
function poemTags(art) {{
  return splitVals((art.querySelector(".tags") || {{}}).value || "");
}}
function poemComps(art) {{
  return splitVals((art.querySelector(".comp") || {{}}).value || "");
}}
function fillChips(boxId, seen, active) {{
  const box = document.getElementById(boxId);
  if (!box) return;
  [...active].forEach(k => {{ if (!seen.has(k)) active.delete(k); }});
  box.innerHTML = "";
  [...seen.entries()].sort((a,b) => a[1].localeCompare(b[1], "pt")).forEach(([k, label]) => {{
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = label;
    b.className = active.has(k) ? "on" : "";
    b.onclick = () => {{
      if (active.has(k)) active.delete(k);
      else active.add(k);
      rebuildChips();
      runSearch();
    }};
    box.appendChild(b);
  }});
}}
function rebuildChips() {{
  const tagsSeen = new Map();
  const compsSeen = new Map();
  document.querySelectorAll("article.poem").forEach(art => {{
    poemTags(art).forEach(v => {{ const k = fold(v); if (!tagsSeen.has(k)) tagsSeen.set(k, v); }});
    poemComps(art).forEach(v => {{ const k = fold(v); if (!compsSeen.has(k)) compsSeen.set(k, v); }});
  }});
  fillChips("chips-tags", tagsSeen, activeTags);
  fillChips("chips-comps", compsSeen, activeComps);
}}
function runSearch() {{
  const q = fold((document.getElementById("busca") || {{}}).value || "");
  let n = 0;
  document.querySelectorAll("article.poem").forEach(art => {{
    const tags = poemTags(art).map(fold);
    const comps = poemComps(art).map(fold);
    const tagsOk = !activeTags.size || [...activeTags].every(k => tags.includes(k));
    const compsOk = !activeComps.size || [...activeComps].every(k => comps.includes(k));
    const m = metaOf(art.getAttribute("data-id"));
    const hideRevOk = !hideReviewed || !m.reviewed;
    const favOk = !onlyFav || !!m.favorite;
    const emptyOk = !onlyEmptyTags || !(m.tags || "").trim();
    const ok = tagsOk && compsOk && hideRevOk && favOk && emptyOk && (!q || haystack(art).includes(q));
    art.classList.toggle("is-hidden", !ok);
    const sep = art.nextElementSibling;
    if (sep && sep.classList.contains("sep")) sep.classList.toggle("is-hidden", !ok);
    const id = art.getAttribute("data-id");
    const li = document.querySelector('ul.indice li[data-id="' + CSS.escape(id) + '"]');
    if (li) li.classList.toggle("is-hidden", !ok);
    if (ok) n++;
  }});
  const st = document.getElementById("busca-status");
  if (st) st.textContent = (q || activeTags.size || activeComps.size || hideReviewed || onlyFav || onlyEmptyTags) ? (n + " poema(s)") : "";
}}
document.getElementById("busca").addEventListener("input", runSearch);
const _apply = apply;
apply = function() {{
  _apply();
  rebuildChips();
  runSearch();
}};
function normPoem(p) {{
  return {{
    file_name: p.file_name,
    title: p.title,
    text: p.text || "",
    body: p.body || "",
    reviewed: !!p.reviewed,
    favorite: !!p.favorite,
    tags: p.tags || [],
    compilacoes: p.compilacoes || p.compilacao || [],
    comentario: p.comentario || ""
  }};
}}
async function loadSiteJson() {{
  if (location.protocol === "file:") return false;
  try {{
    const res = await fetch("poesias.json?t=" + Date.now());
    if (!res.ok) return false;
    const data = await res.json();
    const poems = data.poems || data;
    if (!Array.isArray(poems) || !poems.length) return false;
    const prev = {{}};
    BASE.forEach(p => {{ prev[p.file_name] = p; }});
    BASE = poems.map(p => Object.assign({{}}, prev[p.file_name] || {{}}, normPoem(p)));
    const el = document.getElementById("save-status");
    if (el) el.textContent = "JSON do site carregado.";
    return true;
  }} catch (e) {{
    return false;
  }}
}}
loadSiteJson().finally(() => apply());
</script>
</body>
</html>
"""


def main() -> None:
    """Ponto de entrada: lista .docx, mistura com o JSON antigo, grava JSON + index.html."""
    if len(sys.argv) < 2:
        print("Uso: python3 gerar_poesias.py \"/caminho/da/pasta/Poesias\"")
        sys.exit(1)

    folder = Path(sys.argv[1]).expanduser().resolve()
    if not folder.is_dir():
        sys.exit(f"Pasta não encontrada: {folder}")

    files = sorted(
        [f for f in folder.iterdir() if f.is_file() and f.suffix.lower() == ".docx"],
        key=lambda p: p.name.casefold(),
    )
    if not files:
        sys.exit(f"Nenhum .docx na raiz de {folder}")

    # Tags/checado do JSON que já existe NESTA pasta (cwd), não da pasta dos .docx.
    old = {}
    old_json = Path.cwd() / "poesias.json"
    if old_json.is_file():
        try:
            prev = json.loads(old_json.read_text(encoding="utf-8"))
            for item in prev.get("poems", []):
                old[item.get("file_name")] = item
        except Exception:
            old = {}

    poems = []
    errors = []
    for f in files:
        try:
            title, body, body_html = extract_poem(f)
            raw = extract_text(f)
        except Exception as e:
            errors.append(f"{f.name}: {e}")
            title = title_from_name(f.name)
            raw = ""
            body = ""
            body_html = ""
        prev = old.get(f.name, {})
        poems.append(
            {
                "file_name": f.name,
                "title": title,
                "text": raw,
                "body": body,
                "body_html": body_html,
                "reviewed": bool(prev.get("reviewed")),
                "favorite": bool(prev.get("favorite")),
                "tags": prev.get("tags") or [],
                "compilacoes": prev.get("compilacoes") or prev.get("compilacao") or [],
                "comentario": prev.get("comentario") or "",
            }
        )

    # Saída = pasta de onde o comando foi rodado (cwd), para o Git acompanhar.
    out_dir = Path.cwd()
    json_path = out_dir / "poesias.json"
    html_path = out_dir / "index.html"

    payload = {
        "folder": str(folder),
        "count": len(poems),
        "errors": errors,
        "poems": poems,
    }
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    html_path.write_text(build_html(poems), encoding="utf-8")

    chars = sum(len(p.get("body") or "") for p in poems)
    # ~1800 caracteres úteis por página A4 compacta (11pt, margem 12mm)
    pages = max(1, round(chars / 1800 + len(poems) * 0.08))
    sheets = (pages + 1) // 2

    print(f"Arquivos .docx: {len(poems)}")
    print(f"Erros: {len(errors)}")
    print(f"JSON: {json_path}")
    print(f"HTML: {html_path}")
    print(f"Estimativa impressão compacta: ~{pages} páginas A4  →  ~{sheets} folhas frente e verso")
    print("Abra o HTML no navegador e imprima (frente e verso).")


if __name__ == "__main__":
    main()
