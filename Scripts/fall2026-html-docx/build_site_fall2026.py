#!/usr/bin/env python3
"""
build_site_fall2026.py — Nouvelle arborescence du site (travail fall 2026).

Scinde le document compilé v6 (REFERENCE_FORMANTIQUE_COMPLETE.html) en deux études
distinctes, chacune dans son propre dossier html-docx :

    travail fall 2026/
      Etude-formantique/ETUDE_FORMANTIQUE.html   (sections I–VII)  + media/
      Etude-volumique/ETUDE_VOLUMIQUE.html       (ancienne section VIII)

Aucune analyse n'est relancée et aucun DOCX n'est généré : le script ne fait que
réorganiser le HTML déjà produit par Scripts/v6-html-docx/build_document_complet.py.
La page d'accueil (index.html à la racine) pointe vers ces deux études.

Usage : lancer depuis la racine du repo
    python3 Scripts/fall2026-html-docx/build_site_fall2026.py
"""
import os
import re
import shutil

BASE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC_DIR = os.path.join(BASE, 'Etude-Formants', 'Versions-html-and-docx')
SRC_HTML = os.path.join(SRC_DIR, 'REFERENCE_FORMANTIQUE_COMPLETE.html')

OUT_ROOT = os.path.join(BASE, 'travail fall 2026')
FORM_DIR = os.path.join(OUT_ROOT, 'Etude-formantique')
VOL_DIR = os.path.join(OUT_ROOT, 'Etude-volumique')
FORM_HTML = os.path.join(FORM_DIR, 'ETUDE_FORMANTIQUE.html')
VOL_HTML = os.path.join(VOL_DIR, 'ETUDE_VOLUMIQUE.html')

# Liens relatifs depuis chaque étude
HOME_LINK = '../../index.html'
FORM_LINK = '../Etude-formantique/ETUDE_FORMANTIQUE.html'
VOL_LINK = '../Etude-volumique/ETUDE_VOLUMIQUE.html'

VOL_ANCHOR = 'viii-volume'

# Sections de l'étude volumique : les h2 deviennent des h1 numérotés
VOL_SECTIONS = [
    ('vol-intro',       'I'),
    ('vol-methodo',     'II'),
    ('vol-table',       'III'),
    ('vol-mutes',       'IV'),
    ('vol-homo',        'V'),
    ('vol-plans',       'VI'),
    ('vol-zones',       'VII'),
    ('vol-discoveries', 'VIII'),
]

# Libellés courts du sommaire latéral (étude volumique)
VOL_TOC = [
    (0, "I. Cadre théorique",                    "vol-intro"),
    (0, "II. Méthodologie : Indice de Volume",   "vol-methodo"),
    (0, "III. Indice de Volume par instrument",  "vol-table"),
    (0, "IV. Effet des sourdines",               "vol-mutes"),
    (0, "V. Matrice d'homogénéité (51×51)",      "vol-homo"),
    (0, "VI. Plans orchestraux",                 "vol-plans"),
    (0, "VII. Convergences par zone vocalique",  "vol-zones"),
    (1, "Tableau complet des convergences",      "vol-conv-table"),
    (0, "VIII. Découvertes au-delà du classicisme", "vol-discoveries"),
]

# Titres de groupe de l'étude formantique à doter d'un id (tag, titre, ancre)
GROUP_ANCHORS = [
    ('h2', "Flûtes",                       "bois-flutes"),
    ('h2', "Anches doubles",               "bois-anches"),
    ('h2', "Clarinettes",                  "bois-clarinettes"),
    ('h2', "Cuivres principaux",           "cuivres-principaux"),
    ('h3', "Trompette avec sourdines",     "cuivres-tpt-sord"),
    ('h3', "Trombone avec sourdines",      "cuivres-trb-sord"),
    ('h2', "Saxophones absents du corpus", "sax-absents"),
    ('h2', "Cordes solistes",              "cordes-solistes"),
    ('h2', "Cordes d'ensemble",            "cordes-ensembles"),
]

# ─── CSS ajouté : navigation entre études ────────────────────
CSS_NAV = """
/* ─── Navigation entre études (fall 2026) ─── */
.site-nav { padding: 8px 14px 10px 14px; border-bottom: 1px solid #333; }
.site-nav a {
  display: block; padding: 4px 8px; margin: 2px 0; border-radius: 4px;
  color: #e0e0e0; text-decoration: none; font-size: 12px;
}
.site-nav a:hover { background: #2a2a4e; }
.site-nav a.current { background: #2a2a4e; color: #f6c90e; font-weight: bold; }
.breadcrumb { font-size: 0.85em; color: #888; margin: 0 0 14px 0; }
.breadcrumb a { color: #1565c0; text-decoration: none; }
.breadcrumb a:hover { text-decoration: underline; }
.doc-header.volume {
  background: linear-gradient(135deg, #4a148c 0%, #6a1b9a 60%, #8e24aa 100%);
}
.doc-header.volume .subtitle { color: #e1bee7; }
.doc-header.volume .meta { color: #ce93d8; }
"""

LEVEL_STYLES = {
    0: dict(color='#f6c90e', bold=True,  border='3px solid #f6c90e', pad_left='14px'),
    1: dict(color='#a8d8ea', bold=False, border='0',                 pad_left='28px'),
    2: dict(color='#b8e0d2', bold=False, border='0',                 pad_left='42px'),
}
PAD_TO_LEVEL = {'14px': 0, '28px': 1, '42px': 2}


def read_source():
    with open(SRC_HTML, encoding='utf-8') as f:
        return f.read()


def parse_sidebar_entries(html):
    """Relit le sommaire latéral v6 : liste de (niveau, libellé, ancre)."""
    sidebar = re.search(r'<div class="toc-links">(.*?)</div>', html, re.DOTALL).group(1)
    entries = []
    for m in re.finditer(r'<a href="#([^"]+)" style="[^"]*?padding:3px 14px 3px (\d+px);[^"]*"[^>]*>(.*?)</a>',
                         sidebar):
        anchor, pad, label = m.groups()
        entries.append((PAD_TO_LEVEL[pad], label, anchor))
    return entries


def build_sidebar(entries, current):
    """Sommaire latéral : liens d'étude en tête, puis la table de l'étude courante."""
    nav = [
        ('🏠 Accueil', HOME_LINK, None),
        ('〰 Étude formantique', FORM_LINK, 'formants'),
        ('◉ Étude volumique', VOL_LINK, 'volume'),
    ]
    lines = ['''
<div id="toc-sidebar">
  <div class="toc-header">
    <strong>📋 Sommaire</strong>
    <button onclick="toggleToc()">✕</button>
  </div>
  <div class="site-nav">''']
    for label, href, key in nav:
        cls = ' class="current"' if key == current else ''
        lines.append(f'    <a href="{href}"{cls}>{label}</a>')
    lines.append('  </div>\n  <div class="toc-links">')
    for level, label, anchor in entries:
        s = LEVEL_STYLES[level]
        style = (
            f'display:block; padding:3px 14px 3px {s["pad_left"]}; '
            f'color:{s["color"]}; text-decoration:none; '
            f'border-left:{s["border"]}; line-height:1.4; '
            f'font-weight:{"bold" if s["bold"] else "normal"};'
        )
        lines.append(
            f'    <a href="#{anchor}" style="{style}" '
            f'onmouseover="this.style.background=\'#2a2a4e\'" '
            f'onmouseout="this.style.background=\'\'">'
            f'{label}</a>'
        )
    lines.append('  </div>\n</div>')
    return '\n'.join(lines)


def breadcrumb(label):
    return (f'<p class="breadcrumb"><a href="{HOME_LINK}">Accueil</a> › {label}</p>\n')


def page(title, css, sidebar, content, scripts):
    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<style>
{css}
{CSS_NAV}
</style>
</head>
<body>
{sidebar}
<button id="toc-toggle-btn" onclick="toggleToc()">☰ Sommaire</button>
{content}
{scripts}
</body>
</html>
"""


def split_document(html):
    css = re.search(r'<style>(.*?)</style>', html, re.DOTALL).group(1)

    # Contenu = après le bouton toggle, jusqu'aux scripts finaux
    start = html.index('<button id="toc-toggle-btn"')
    start = html.index('\n', start) + 1
    vol_start = html.index(f'<h1 id="{VOL_ANCHOR}">')
    first_script = html.index('<script>', vol_start)
    end_body = html.rindex('</body>')

    formant_content = html[start:vol_start]
    volume_content = html[html.index('\n', vol_start) + 1:first_script]
    scripts = html[first_script:end_body]

    # Les scripts de tri/filtre ne concernent que l'étude volumique ;
    # le dernier bloc (toggle sommaire) est commun aux deux pages.
    sidebar_js_start = scripts.rindex('<script>')
    vol_scripts = scripts
    form_scripts = scripts[sidebar_js_start:]

    # Retirer le séparateur final laissé avant l'ancienne section VIII
    formant_content = re.sub(r'\s*<hr class="section-sep"/>\s*$', '\n', formant_content)
    return css, formant_content, volume_content, form_scripts, vol_scripts


def build_formants(html, css, content, scripts):
    entries = [e for e in parse_sidebar_entries(html)]
    cut = next(i for i, e in enumerate(entries) if e[2] == VOL_ANCHOR)
    entries = entries[:cut]

    # Table des matières en tête : ajouter la section VII qui y manquait
    content = content.replace(
        '  <li><a href="#vi-synthese">VI. Synthèse — Convergences Formantiques</a></li>\n',
        '  <li><a href="#vi-synthese">VI. Synthèse — Convergences Formantiques</a></li>\n'
        '  <li><a href="#vii-enveloppes">VII. Enveloppes spectrales par instrument</a></li>\n',
    )
    # Titres de groupe sans id dans v6 : ancres du sommaire latéral mortes
    for tag, title, anchor in GROUP_ANCHORS:
        content = content.replace(f'<{tag}>{title}</{tag}>',
                                  f'<{tag} id="{anchor}">{title}</{tag}>', 1)
    content = content.replace(
        '<div class="doc-header">',
        breadcrumb('Étude formantique') + '<div class="doc-header">', 1,
    )
    content = content.replace(
        '<p class="subtitle">Étude quantitative des formants spectraux instrumentaux</p>',
        '<p class="subtitle">Étude formantique — formants spectraux instrumentaux</p>', 1,
    )
    # Renvoi vers l'étude volumique en fin de document
    content += (
        '\n<hr class="section-sep"/>\n'
        '<div class="section-intro general"><p>L\'étude des attributs du timbre selon Koechlin '
        '(volume, densité, homogénéité, plans orchestraux) fait désormais l\'objet d\'un document '
        f'distinct : <a href="{VOL_LINK}"><strong>Étude volumique</strong></a>.</p></div>\n'
    )

    out = page(
        "Étude formantique — Référence Formantique de l'Orchestre",
        css, build_sidebar(entries, 'formants'), content, scripts,
    )
    os.makedirs(FORM_DIR, exist_ok=True)
    with open(FORM_HTML, 'w', encoding='utf-8') as f:
        f.write(out)
    print(f"✓ {os.path.relpath(FORM_HTML, BASE)} — {len(out):,} caractères")
    return content


def build_volume(css, content, scripts):
    # Hiérarchie : h2 → h1 numérotés, h3 → h2, h4 → h3
    content = re.sub(r'<(/?)h4\b', r'<\1h5_tmp', content)
    content = re.sub(r'<(/?)h3\b', r'<\1h4_tmp', content)
    for anchor, num in VOL_SECTIONS:
        content = re.sub(rf'<h2 id="{anchor}">(.*?)</h2>',
                         rf'<h1 id="{anchor}">{num}. \1</h1>', content, count=1)
    content = re.sub(r'<(/?)h4_tmp\b', r'<\1h2', content)
    content = re.sub(r'<(/?)h5_tmp\b', r'<\1h3', content)
    leftover = re.findall(r'<h2 id="vol-[^"]*"', content)
    assert not any(a.split('"')[1] in dict(VOL_SECTIONS) for a in leftover), leftover

    header = breadcrumb('Étude volumique') + '''<div class="doc-header volume">
  <h1>Étude volumique de l'Orchestre</h1>
  <p class="subtitle">Attributs du timbre selon Koechlin — volume, densité, homogénéité</p>
  <p class="subtitle">51 instruments (sourdines comprises) · 5&nbsp;758 échantillons · 185 segments de registre</p>
  <p class="meta">
    Cadre : Koechlin (1954) · Chiasson &amp; Duchesneau (2008) · Chiasson (2010) · Stevens · van Dinther &amp; Patterson (2006)<br>
    Croisement avec l'<a href="''' + FORM_LINK + '''" style="color:#f3e5f5;">étude formantique</a> (cluster de convergence 450–502 Hz)
  </p>
</div>

<h2>Table des matières</h2>
<nav style="background:white;border:1px solid #ddd;border-radius:6px;padding:18px 24px;margin:16px 0;">
<ol style="line-height:2.2;" type="I">
'''
    for level, label, anchor in VOL_TOC:
        if level == 0:
            header += f'  <li><a href="#{anchor}">{label.split(". ", 1)[1]}</a></li>\n'
    header += '</ol>\n</nav>\n'

    out = page(
        "Étude volumique — Attributs du timbre selon Koechlin",
        css, build_sidebar(VOL_TOC, 'volume'), header + content, scripts,
    )
    os.makedirs(VOL_DIR, exist_ok=True)
    with open(VOL_HTML, 'w', encoding='utf-8') as f:
        f.write(out)
    print(f"✓ {os.path.relpath(VOL_HTML, BASE)} — {len(out):,} caractères")


def copy_media(content, dest_dir):
    """Copie uniquement les images référencées par le document."""
    refs = sorted(set(re.findall(r'src="(media/[^"]+)"', content)))
    os.makedirs(os.path.join(dest_dir, 'media'), exist_ok=True)
    missing = []
    for rel in refs:
        src = os.path.join(SRC_DIR, rel)
        if not os.path.exists(src):
            missing.append(rel)
            continue
        shutil.copy2(src, os.path.join(dest_dir, rel))
    print(f"✓ {len(refs) - len(missing)} images copiées dans {os.path.relpath(dest_dir, BASE)}/media")
    for rel in missing:
        print(f"  ⚠ image absente : {rel}")


if __name__ == '__main__':
    html = read_source()
    css, form_content, vol_content, form_scripts, vol_scripts = split_document(html)
    form_content = build_formants(html, css, form_content, form_scripts)
    copy_media(form_content, FORM_DIR)
    if 'src="media/' in vol_content:
        copy_media(vol_content, VOL_DIR)
    build_volume(css, vol_content, vol_scripts)
