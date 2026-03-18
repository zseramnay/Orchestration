#!/usr/bin/env python3
"""
build_document_complet.py — Assemble toutes les sections en un seul fichier HTML + DOCX.

HTML : sidebar de navigation fixe (même style que v33-REFERENCE_FORMANTIQUE_v3.html)
       avec 3 niveaux : sections principales (jaune), instruments (bleu clair), sous-sections (vert clair)
       Contenu dans une zone scrollable avec marge gauche de 320 px.

DOCX : table des matières Word automatique (champs TC + styles Heading),
       chaque section commence sur une nouvelle page.

Usage : lancer depuis la racine du repo Formants/
    python3 Scripts/v4-html-docx-enriched/build_document_complet.py
"""
import os
import sys
import unicodedata
import re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *

# ─── python-docx spécifique ──────────────────────────────────
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

load_all_csvs()

OUT_HTML = os.path.join(OUT_DIR, 'REFERENCE_FORMANTIQUE_COMPLETE.html')
OUT_DOCX = os.path.join(OUT_DIR, 'REFERENCE_FORMANTIQUE_COMPLETE.docx')

# ═══════════════════════════════════════════════════════════
# STRUCTURE DE NAVIGATION
# Hiérarchie : (niveau, label, anchor, is_section)
# niveau 0 = section principale (jaune, bold, barre gauche)
# niveau 1 = instrument / sous-section (bleu clair)
# niveau 2 = sous-sous-section (vert clair)
# ═══════════════════════════════════════════════════════════

TOC_STRUCTURE = [
    (0, "I. Introduction et Méthodologie",     "i-introduction"),
    (1, "Sources des données",                  "sources"),
    (1, "Méthodologie d'extraction",            "methodo"),
    (1, "Le Formant Principal (Fp)",            "fp-centroide"),
    (1, "Correspondance voyelles–fréquences",   "voyelles"),

    (0, "II. Les Bois",                         "ii-bois"),
    (1, "Flûtes",                               "bois-flutes"),
    (2, "Petite flûte",                         "bois_piccolo"),
    (2, "Flûte",                                "bois_flute"),
    (2, "Flûte basse",                          "bois_bass_flute"),
    (2, "Flûte contrebasse",                    "bois_contrabass_flute"),
    (1, "Anches doubles",                       "bois-anches"),
    (2, "Hautbois",                             "bois_hautbois"),
    (2, "Cor anglais",                          "bois_cor_anglais"),
    (2, "Basson",                               "bois_basson"),
    (2, "Contrebasson",                         "bois_contrebasson"),
    (1, "Clarinettes",                          "bois-clarinettes"),
    (2, "Clarinette en Mib",                    "bois_clarinet_eb"),
    (2, "Clarinette en Sib",                    "bois_clarinet_bb"),
    (2, "Clarinette basse",                     "bois_clarinet_basse"),
    (2, "Clarinette contrebasse",               "bois_clarinet_cb"),

    (0, "III. Les Cuivres",                     "iii-cuivres"),
    (1, "Instruments principaux",               "cuivres-principaux"),
    (2, "Cor en Fa",                            "cuivres_cor"),
    (2, "Trompette en Ut",                      "cuivres_trompette"),
    (2, "Trombone ténor",                       "cuivres_trombone"),
    (2, "Trombone basse",                       "cuivres_trb_basse"),
    (2, "Tuba basse",                           "cuivres_tuba_basse"),
    (2, "Tuba contrebasse",                     "cuivres_tuba_cb"),
    (1, "Cor avec sourdine",                    "cuivres_cor_sord"),
    (1, "Trombone avec sourdines",              "cuivres-trb-sord"),
    (2, "Trombone + sourd. cup",                "cuivres_trb_cup"),
    (2, "Trombone + sourd. sèche",              "cuivres_trb_straight"),
    (2, "Trombone + sourd. harmon",             "cuivres_trb_harmon"),
    (2, "Trombone + sourd. wah (ouvert)",       "cuivres_trb_wah_open"),
    (2, "Trombone + sourd. wah (fermé)",        "cuivres_trb_wah_closed"),
    (1, "Trompette avec sourdines",             "cuivres-tpt-sord"),
    (2, "Trompette + sourd. cup",               "cuivres_tpt_cup"),
    (2, "Trompette + sourd. sèche",             "cuivres_tpt_straight"),
    (2, "Trompette + sourd. harmon",            "cuivres_tpt_harmon"),
    (2, "Trompette + sourd. wah (ouvert)",      "cuivres_tpt_wah_open"),
    (2, "Trompette + sourd. wah (fermé)",       "cuivres_tpt_wah_closed"),
    (1, "Tuba avec sourdine",                   "cuivres_tuba_sord"),

    (0, "IV. Les Saxophones",                   "iv-saxophones"),
    (1, "Saxophone alto",                       "sax_alto"),
    (1, "Saxophones absents du corpus",         "sax-absents"),

    (0, "V. Les Cordes",                        "v-cordes"),
    (1, "Cordes solistes",                      "cordes-solistes"),
    (2, "Violon",                               "cordes_violon"),
    (2, "Violon + sourdine",                    "cordes_vln_sord"),
    (2, "Violon + sourdine piombo",             "cordes_vln_piombo"),
    (2, "Alto",                                 "cordes_alto"),
    (2, "Alto + sourdine",                      "cordes_alt_sord"),
    (2, "Alto + sourdine piombo",               "cordes_alt_piombo"),
    (2, "Violoncelle",                          "cordes_violoncelle"),
    (2, "Violoncelle + sourdine",               "cordes_vcl_sord"),
    (2, "Violoncelle + sourdine piombo",        "cordes_vcl_piombo"),
    (2, "Contrebasse",                          "cordes_contrebasse"),
    (2, "Contrebasse + sourdine",               "cordes_cb_sord"),
    (1, "Cordes d'ensemble",                    "cordes-ensembles"),
    (2, "Ensemble de violons",                  "cordes_vln_ens"),
    (2, "Ensemble de violons + sourdine",       "cordes_vln_ens_sord"),
    (2, "Ensemble d'altos",                     "cordes_alt_ens"),
    (2, "Ensemble d'altos + sourdine",          "cordes_alt_ens_sord"),
    (2, "Ensemble de violoncelles",             "cordes_vcl_ens"),
    (2, "Ensemble de violoncelles + sourdine",  "cordes_vcl_ens_sord"),
    (2, "Ensemble de contrebasses",             "cordes_cb_ens"),

    (0, "VI. Synthèse",                         "vi-synthese"),
    (1, "Cluster de convergence 450–502 Hz",    "cluster"),
    (1, "Positions F1/Fp",                      "positions-f1"),
    (1, "Matrice de convergence — base",        "matrices"),
    (1, "Matrice de convergence — complète",    "matrices"),
    (1, "Doublures vérifiées",                  "doublures-verifiees"),
    (1, "6 Principes d'orchestration",          "principes"),
    (1, "Concordance multi-sources",            "concordance"),
]

# ═══════════════════════════════════════════════════════════
# HELPER : slugification pour les ancres id=""
# ═══════════════════════════════════════════════════════════
def slugify(text):
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    text = text.lower()
    text = re.sub(r'[^a-z0-9]+', '-', text)
    return text.strip('-')

# ═══════════════════════════════════════════════════════════
# GÉNÉRATION DU SIDEBAR HTML
# ═══════════════════════════════════════════════════════════
def build_sidebar():
    # Couleurs et styles par niveau, calqués sur le v33
    level_styles = {
        0: dict(color='#f6c90e', bold=True,   border='3px solid #f6c90e', pad_left='14px'),
        1: dict(color='#a8d8ea', bold=False,  border='0',                  pad_left='28px'),
        2: dict(color='#b8e0d2', bold=False,  border='0',                  pad_left='42px'),
    }

    lines = []
    lines.append('''
<div id="toc-sidebar">
  <div class="toc-header">
    <strong>📋 Sommaire</strong>
    <button onclick="toggleToc()">✕</button>
  </div>
  <div class="toc-links">''')

    for level, label, anchor in TOC_STRUCTURE:
        s = level_styles[level]
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

# ═══════════════════════════════════════════════════════════
# IMPORTATION DES SCRIPTS DE SECTION
# Chaque script expose build_html(output_path) mais on veut
# juste récupérer le <body> pour l'injecter ici.
# On importe les données et fonctions directement.
# ═══════════════════════════════════════════════════════════

def extract_body(script_path):
    """
    Exécute le build_html d'un script dans un fichier temporaire
    et retourne uniquement le contenu entre <body> et </body>.
    """
    import subprocess, tempfile
    tmp = tempfile.mktemp(suffix='.html')
    # Exécuter le script en redirigeant OUT_DIR vers un tmp
    env_code = f"""
import sys, os
sys.path.insert(0, os.path.dirname('{script_path}'))
# Override OUT_DIR to tmp location
import common
common.OUT_DIR = '{os.path.dirname(tmp)}'
common.OUT_IMG = os.path.join(common.OUT_DIR, 'media')
"""
    result = subprocess.run(
        ['python3', script_path],
        capture_output=True, text=True,
        cwd=BASE
    )
    return None  # On n'utilise pas cette approche


# Alternative propre : chaque script a une fonction get_body_html()
# On va plutôt construire le HTML complet en appelant les fonctions
# directement après import. Mais les scripts ont des globals (all_info)
# qui se construisent à l'import. On va donc construire le HTML
# section par section en répliquant la logique ici,
# en s'appuyant sur les modules déjà importés.

# ─── Import des données de chaque section ────────────────────
# On importe les modules de section en les chargeant comme des modules
# Le plus propre : chaque build_xxx exporte une fonction build_html_body()
# Puisqu'ils n'en ont pas, on les exécute dans un sous-process et on lit
# le fichier temporaire généré.

import subprocess, tempfile, shutil

def get_section_body(script_name, tmp_dir):
    """
    Lance un script, lit son HTML de sortie, extrait le contenu <body>.
    Réécrit les chemins d'images relatifs pour pointer vers media/.
    """
    script_path = os.path.join(BASE, 'Scripts', 'v4-html-docx-enriched', script_name)
    out_file_map = {
        'build_intro_html_docx.py':    'section_intro_v4.html',
        'build_bois_html_docx.py':     'section_bois_v4.html',
        'build_cuivres_html_docx.py':  'section_cuivres_v4.html',
        'build_sax_html_docx.py':      'section_sax_v4.html',
        'build_cordes_html_docx.py':   'section_cordes_v4.html',
        'build_synthese_html_docx.py': 'section_synthese_v4.html',
    }
    out_filename = out_file_map[script_name]
    out_path = os.path.join(OUT_DIR, out_filename)

    print(f"  → Génération {script_name} ...", flush=True)
    result = subprocess.run(
        ['python3', script_path],
        cwd=BASE,
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"  ⚠ Erreur dans {script_name}:\n{result.stderr[:500]}")
        return f'<div class="section-error">Erreur : {script_name}</div>'

    if not os.path.exists(out_path):
        print(f"  ⚠ Fichier non trouvé : {out_path}")
        return ''

    with open(out_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # Extraire contenu entre <body> et </body>
    m = re.search(r'<body[^>]*>(.*)</body>', html, re.DOTALL | re.IGNORECASE)
    if not m:
        return html

    body = m.group(1).strip()

    # Supprimer les <style> résiduels qui pourraient dupliquer le CSS global
    body = re.sub(r'<style[^>]*>.*?</style>', '', body, flags=re.DOTALL)

    print(f"  ✓ {script_name} — {len(body):,} caractères")
    return body


# ═══════════════════════════════════════════════════════════
# CSS DOCUMENT COMPLET (adapté pour sidebar fixe)
# ═══════════════════════════════════════════════════════════
CSS_COMPLET = """
/* ─── Reset & layout ─── */
*, *::before, *::after { box-sizing: border-box; }
html { margin: 0; padding: 0; background: #e6e9ed; }
body {
  font-family: 'Segoe UI', Helvetica, Arial, sans-serif;
  font-size: 13px; line-height: 1.5; color: #333;
  background: #ffffff;
  max-width: 900px;
  margin: 0 auto;
  margin-left: 340px;   /* laisser place au sidebar (320px + 20px) */
  padding: 40px 50px;
  min-height: 100vh;
  box-shadow: 0 4px 20px rgba(0,0,0,0.12);
  transition: margin-left 0.3s;
}

/* ─── Sidebar ─── */
#toc-sidebar {
  position: fixed; top: 0; left: 0;
  width: 320px; height: 100vh;
  overflow-y: auto; overflow-x: hidden;
  background: #1a1a2e; color: #e0e0e0;
  font-family: sans-serif; font-size: 12px;
  z-index: 9999;
  box-shadow: 3px 0 10px rgba(0,0,0,0.4);
  padding: 0;
  transform: translateX(0);
  transition: transform 0.3s;
}
.toc-header {
  padding: 10px 14px;
  border-bottom: 1px solid #333;
  display: flex;
  justify-content: space-between;
  align-items: center;
  position: sticky; top: 0;
  background: #1a1a2e;
  z-index: 2;
}
.toc-header strong { font-size: 13px; color: #a8d8ea; }
.toc-header button {
  background: #333; color: #ccc; border: none;
  cursor: pointer; padding: 2px 8px;
  border-radius: 3px; font-size: 11px;
}
.toc-links { padding: 6px 0; }
.toc-links a:hover { background: #2a2a4e !important; }

/* ─── Bouton toggle ─── */
#toc-toggle-btn {
  position: fixed; top: 10px; left: 10px; z-index: 10000;
  background: #1a1a2e; color: #a8d8ea;
  border: 1px solid #a8d8ea;
  padding: 5px 10px; cursor: pointer;
  border-radius: 4px; font-size: 12px;
  display: none;
}

/* ─── Typographie ─── */
h1 { color: #1a237e; border-bottom: 3px solid #2e7d32; padding-bottom: 10px;
     font-size: 1.8em; margin-top: 50px; margin-bottom: 16px; }
h2 { color: #2e7d32; margin-top: 44px; border-left: 5px solid #2e7d32;
     padding-left: 14px; font-size: 1.35em; }
h3 { color: #1b5e20; margin-top: 32px; font-size: 1.15em; }
h4 { color: #555; margin-top: 14px; font-size: 1.02em; }
p  { margin: 6px 0; }

/* ─── Cartes instrument ─── */
.instrument-card {
  background: white; border: 1px solid #ddd; border-radius: 8px;
  padding: 20px; margin: 20px 0; box-shadow: 0 2px 6px rgba(0,0,0,0.07);
  scroll-margin-top: 20px;
}
.formant-graph { max-width: 100%; border: 1px solid #eee; border-radius: 4px; }
.description {
  font-style: italic; color: #555; background: #e8f5e9;
  padding: 10px 14px; border-left: 4px solid #4caf50;
  margin: 10px 0; border-radius: 0 4px 4px 0;
}
.fp-note { color: #1b5e20; font-weight: bold; margin: 8px 0; }
.note-v4 {
  background: #fff3e0; border-left: 4px solid #ff9800;
  padding: 8px 12px; margin: 10px 0; font-size: 0.9em; color: #555;
}
.cluster-badge {
  display: inline-block; background: #e53935; color: white;
  padding: 2px 8px; border-radius: 10px; font-size: 0.78em;
  margin-left: 8px; font-weight: bold;
}
.yan-badge {
  display: inline-block; background: #ff9800; color: white;
  padding: 2px 7px; border-radius: 10px; font-size: 0.78em; margin-left: 8px;
}

/* ─── Introductions de section ─── */
.section-intro {
  padding: 14px 18px; border-radius: 6px; margin: 14px 0;
}
.section-intro.intro   { background: #e8eaf6; border-left: 5px solid #283593; }
.section-intro.bois    { background: #e8f5e9; border-left: 5px solid #2e7d32; }
.section-intro.cuivres { background: #fff3e0; border-left: 5px solid #e64a19; }
.section-intro.cordes  { background: #e3f2fd; border-left: 5px solid #1565c0; }
.section-intro.sax     { background: #fce4ec; border-left: 5px solid #ad1457; }
.section-intro.general { background: #f3e5f5; border-left: 5px solid #6a1b9a; }
.fp-explanation {
  background: #e8f5e9; border: 1px solid #a5d6a7;
  border-radius: 6px; padding: 12px 16px; margin: 14px 0; font-size: 0.93em;
}

/* ─── Tableaux ─── */
table { border-collapse: collapse; width: 100%; margin: 12px 0;
        font-size: 0.87em; word-break: break-word; }
th, td { border: 1px solid #ccc; padding: 5px 8px; text-align: center; }
.tech-table .header th,
.ref-table .header th  { background: #1a3a5c; color: white; font-size: 0.88em; }
.ref-table .header th  { background: #37474f; }
.vowel-table .header th { background: #4a148c; color: white; }
tr:nth-child(even) { background: #fafafa; }

/* ─── Doublures ─── */
.doublures-box {
  background: #fffde7; border: 1px solid #f9a825; border-radius: 6px;
  padding: 12px 16px; margin: 16px 0;
}
.doublures-box h4 { color: #f57f17; margin: 0 0 8px 0; }
.doublures-table th { background: #f57f17; color: white; }
.doublures-table td, .doublures-table th { border-color: #f9a825; padding: 4px 7px; }

/* ─── Matrices ─── */
.matrix-container { overflow-x: auto; margin: 14px 0; }

/* ─── Séparateur de section ─── */
.section-sep {
  border: none; border-top: 2px solid #e0e0e0;
  margin: 50px 0 30px 0;
}

/* ─── Source note ─── */
.source-note {
  font-size: 0.82em; color: #888; margin-top: 36px;
  border-top: 1px solid #ddd; padding-top: 10px;
}

/* ─── Doc header ─── */
.doc-header {
  background: linear-gradient(135deg, #1a237e 0%, #283593 60%, #1565c0 100%);
  color: white; padding: 30px 36px; border-radius: 10px; margin-bottom: 30px;
}
.doc-header h1 { color: white; border: none; margin: 0 0 10px 0; font-size: 1.9em; }
.doc-header .subtitle { color: #b3c5ff; font-size: 1.02em; margin: 3px 0; }
.doc-header .meta { color: #90a4c8; font-size: 0.86em; margin-top: 14px; }
"""

# ─── JavaScript sidebar ──────────────────────────────────────
JS_SIDEBAR = """
<script>
var tocVisible = true;
function toggleToc() {
  var toc = document.getElementById('toc-sidebar');
  var btn = document.getElementById('toc-toggle-btn');
  tocVisible = !tocVisible;
  if (tocVisible) {
    toc.style.transform = 'translateX(0)';
    btn.style.display = 'none';
    document.body.style.marginLeft = '340px';
  } else {
    toc.style.transform = 'translateX(-320px)';
    btn.style.display = 'block';
    document.body.style.marginLeft = '20px';
  }
}
// Highlight active section in sidebar
window.addEventListener('scroll', function() {
  var sections = document.querySelectorAll('[id]');
  var links = document.querySelectorAll('#toc-sidebar a');
  var scrollTop = window.scrollY + 80;
  var current = '';
  sections.forEach(function(s) {
    if (s.offsetTop <= scrollTop) current = s.id;
  });
  links.forEach(function(a) {
    var href = a.getAttribute('href');
    if (href && href === '#' + current) {
      a.style.background = '#2a2a4e';
    } else {
      a.style.background = '';
    }
  });
});
</script>
"""

# ═══════════════════════════════════════════════════════════
# BUILD HTML COMPLET
# ═══════════════════════════════════════════════════════════
def build_html_complet():
    print("\n=== BUILD HTML COMPLET ===")
    tmp_dir = tempfile.mkdtemp()

    sections_order = [
        'build_intro_html_docx.py',
        'build_bois_html_docx.py',
        'build_cuivres_html_docx.py',
        'build_sax_html_docx.py',
        'build_cordes_html_docx.py',
        'build_synthese_html_docx.py',
    ]

    # Générer chaque section
    bodies = {}
    for script in sections_order:
        bodies[script] = get_section_body(script, tmp_dir)

    # Assembler
    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Référence Formantique de l'Orchestre — Document Complet</title>
<style>
{CSS_COMPLET}
</style>
</head>
<body>
"""
    # Sidebar
    html += build_sidebar()
    html += '\n<button id="toc-toggle-btn" onclick="toggleToc()">☰ Sommaire</button>\n'

    # Contenu
    for i, script in enumerate(sections_order):
        if i > 0:
            html += '\n<hr class="section-sep"/>\n'
        html += bodies[script]

    html += '\n' + JS_SIDEBAR
    html += '\n</body>\n</html>\n'

    with open(OUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"\n✓ HTML complet : {OUT_HTML}")
    print(f"  Taille : {os.path.getsize(OUT_HTML):,} octets")

    shutil.rmtree(tmp_dir, ignore_errors=True)


# ═══════════════════════════════════════════════════════════
# BUILD DOCX COMPLET avec Table des matières
# ═══════════════════════════════════════════════════════════

def add_toc_field(doc):
    """
    Insère un champ TOC Word automatique (mis à jour à l'ouverture du doc).
    """
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fld_char_begin = OxmlElement('w:fldChar')
    fld_char_begin.set(qn('w:fldCharType'), 'begin')
    run._r.append(fld_char_begin)

    instr_text = OxmlElement('w:instrText')
    instr_text.set(qn('xml:space'), 'preserve')
    instr_text.text = 'TOC \\o "1-3" \\h \\z \\u'
    run._r.append(instr_text)

    fld_char_sep = OxmlElement('w:fldChar')
    fld_char_sep.set(qn('w:fldCharType'), 'separate')
    run._r.append(fld_char_sep)

    fld_char_end = OxmlElement('w:fldChar')
    fld_char_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_char_end)

    # Note utilisateur
    note = doc.add_paragraph()
    r = note.add_run("⟳ Faire un clic droit sur la table des matières → Mettre à jour les champs → Mettre à jour toute la table")
    r.font.size = Pt(9)
    r.font.italic = True
    r.font.color.rgb = RGBColor(150, 150, 150)


def add_page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)


def build_docx_complet():
    """
    Assemble les sections en un seul DOCX.
    Chaque section commence sur une nouvelle page.
    Table des matières Word automatique en début de document.
    """
    print("\n=== BUILD DOCX COMPLET ===")

    # On importe chaque module de section et on appelle build_docx
    # sur un fichier temporaire, puis on copie le contenu dans le doc final.
    # Approche plus simple et robuste : lancer chaque build_docx,
    # puis fusionner les fichiers avec python-docx (copie de sections).

    import importlib.util

    # Générer les DOCX de section
    section_scripts = [
        ('build_intro_html_docx.py',    'section_intro_v4.docx'),
        ('build_bois_html_docx.py',     'section_bois_v4.docx'),
        ('build_cuivres_html_docx.py',  'section_cuivres_v4.docx'),
        ('build_sax_html_docx.py',      'section_sax_v4.docx'),
        ('build_cordes_html_docx.py',   'section_cordes_v4.docx'),
        ('build_synthese_html_docx.py', 'section_synthese_v4.docx'),
    ]

    section_docx_paths = []
    for script_name, out_filename in section_scripts:
        script_path = os.path.join(BASE, 'Scripts', 'v4-html-docx-enriched', script_name)
        out_path = os.path.join(OUT_DIR, out_filename)
        if not os.path.exists(out_path):
            print(f"  → Génération {script_name} pour DOCX ...", flush=True)
            result = subprocess.run(['python3', script_path], cwd=BASE,
                                     capture_output=True, text=True)
            if result.returncode != 0:
                print(f"  ⚠ Erreur : {result.stderr[:300]}")
                continue
        section_docx_paths.append(out_path)
        print(f"  ✓ {out_filename}")

    # Créer le document final
    doc = new_docx()
    # Marges légèrement plus grandes pour le doc complet
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin   = Cm(3.0)
        section.right_margin  = Cm(2.5)

    # ── Page de garde ──
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Référence Formantique de l'Orchestre")
    run.bold = True
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(26, 35, 126)

    doc.add_paragraph()
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("Étude quantitative des formants spectraux instrumentaux")
    r2.font.size = Pt(14)
    r2.font.color.rgb = RGBColor(70, 70, 70)

    doc.add_paragraph()
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run(
        "5 914 échantillons · 30 instruments · Mars 2026\n"
        "SOL2020 IRCAM · Yan_Adds · Backus (1969) · Giesler (1985) · Meyer (2009)\n"
        "Pipeline v22 validé · 93 % de concordance multi-sources"
    )
    r3.font.size = Pt(10)
    r3.font.color.rgb = RGBColor(100, 100, 100)

    add_page_break(doc)

    # ── Table des matières ──
    h_toc = doc.add_paragraph("Table des matières", style='Heading 1')
    h_toc.runs[0].font.color.rgb = RGBColor(26, 35, 126)

    add_toc_field(doc)
    add_page_break(doc)

    # ── Fusion des sections ──
    # On utilise la technique de copy de body XML entre documents
    from lxml import etree

    def copy_body_content(src_path, dest_doc):
        """
        Copie le contenu body d'un document source dans dest_doc.
        Réinsère les images avec des rId uniques pour éviter les doublons.
        """
        src_doc = Document(src_path)
        add_page_break(dest_doc)

        src_body  = src_doc.element.body
        src_rels  = src_doc.part.rels
        dest_part = dest_doc.part

        # Map rId source → rId destination (pour éviter les doublons)
        rid_map = {}

        # Pré-enregistrer toutes les images du document source
        for rid, rel in src_rels.items():
            if "image" in rel.reltype:
                try:
                    img_part = rel.target_part
                    # Créer une relation dans dest_doc et mémoriser le nouveau rId
                    new_rid = dest_part.relate_to(
                        img_part,
                        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
                    )
                    rid_map[rid] = new_rid
                except Exception:
                    rid_map[rid] = rid  # fallback

        import copy
        for child in src_body:
            if child.tag.endswith('}sectPr'):
                continue
            new_child = copy.deepcopy(child)

            # Remapper les références d'images (blipFill)
            for blip in new_child.iter(qn('a:blip')):
                embed = blip.get(qn('r:embed'))
                if embed and embed in rid_map:
                    blip.set(qn('r:embed'), rid_map[embed])

            dest_doc.element.body.append(new_child)

    for path in section_docx_paths:
        print(f"  → Fusion : {os.path.basename(path)}", flush=True)
        try:
            copy_body_content(path, doc)
            print(f"  ✓ {os.path.basename(path)}")
        except Exception as e:
            print(f"  ⚠ Erreur fusion {os.path.basename(path)}: {e}")

    doc.save(OUT_DOCX)
    print(f"\n✓ DOCX complet : {OUT_DOCX}")
    print(f"  Taille : {os.path.getsize(OUT_DOCX):,} octets")
    print("  ⚠ Ouvrir le DOCX et appuyer sur Ctrl+A puis F9 pour mettre à jour la table des matières.")


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════
if __name__ == '__main__':
    import sys
    mode = sys.argv[1] if len(sys.argv) > 1 else 'both'

    if mode in ('html', 'both'):
        build_html_complet()

    if mode in ('docx', 'both'):
        build_docx_complet()

    print(f"\n{'='*60}")
    if mode in ('html', 'both'):
        print(f"HTML : {OUT_HTML}")
    if mode in ('docx', 'both'):
        print(f"DOCX : {OUT_DOCX}")
    print("Usage : python3 build_document_complet.py [html|docx|both]")
