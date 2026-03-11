#!/usr/bin/env python3
"""
Build Section Saxophones + Section Cuivres — HTML + graphs
Source unique: CSV v22
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
import matplotlib.ticker
import csv, os

OUT_IMG = "/home/claude/Formants/Version html/media"
OUT_HTML = "/home/claude/Formants/Version html"

def sf(v):
    try: return float(v)
    except: return 0.0

DATA = {}
with open('/mnt/user-data/uploads/formants_all_techniques.csv', 'r') as f:
    for row in csv.DictReader(f): DATA[(row['instrument'], row['technique'])] = row
with open('/home/claude/formants_yan_adds.csv', 'r') as f:
    for row in csv.DictReader(f): DATA[(row['instrument'], row['technique'])] = row

def get_f(inst, tech):
    r = DATA.get((inst, tech))
    if not r: return None
    return {'n': int(r['n_samples']), 'F': [round(sf(r[f'F{i}_hz'])) for i in range(1,7)]}

FP = {
    'Sax_Alto': 1440,
    'Horn': 738, 'Trumpet_C': 1046, 'Trombone': 1218, 'Bass_Tuba': 1239,
    'Bass_Trombone': 1335, 'Contrabass_Tuba': 1182,
}

VOWEL_ZONES = [
    (100, 400, '#DCEEFB', 'u (oo)\nProfondeur'), (400, 600, '#D5ECD5', 'o (oh)\nPlénitude'),
    (600, 800, '#FDE8CE', 'å (aw)\nTransition'), (800, 1250, '#F8D5D5', 'a (ah)\nPuissance'),
    (1250, 2600, '#E8D5F0', 'e (eh)\nClarté'), (2600, 6000, '#FFF8D0', 'i (ee)\nBrillance'),
]
FC = ['#D32F2F', '#E64A19', '#F57C00', '#FFA000', '#FBC02D', '#CDDC39']
FA = [1.0, 0.85, 0.7, 0.55, 0.4, 0.3]

def make_graph(display, filename, n, formants, fp=None, family_color='#E65100'):
    valid = [(i,f) for i,f in enumerate(formants) if f > 0]
    if not valid: return None
    mf = max(f for _,f in valid) + 500; mf = min(max(mf, 3000), 6500)
    fig, ax = plt.subplots(figsize=(9.6, 4.8), dpi=150)
    for lo,hi,c,l in VOWEL_ZONES:
        if lo < mf:
            ax.axvspan(lo, min(hi,mf), alpha=0.35, color=c, zorder=0)
            mid = (lo + min(hi,mf)) / 2
            if mid < mf*0.95:
                ax.text(mid, 0.97, l, ha='center', va='top', fontsize=7, color='#666', fontweight='bold', transform=ax.get_xaxis_transform())
    bw = mf * 0.012
    for i,freq in valid:
        bh = 1.0*(1.0-i*0.12)
        ax.bar(freq, bh, width=bw*(1.2-i*0.05), color=FC[i], alpha=FA[i], edgecolor='#333', linewidth=0.8, zorder=3)
        ax.text(freq, bh+0.03, f"F{i+1}\n{freq} Hz", ha='center', va='bottom', fontsize=8, fontweight='bold', color='#333', zorder=5)
    f1 = formants[0]
    if fp and abs(fp-f1) > 30:
        ax.plot(fp, 0.5, marker='D', markersize=14, color='#1B5E20', markeredgecolor='black', markeredgewidth=1.5, zorder=6)
        ax.annotate(f"Fp = {fp} Hz\n(centroïde)", xy=(fp,0.5), xytext=(fp,0.65), ha='center', fontsize=8, fontweight='bold', color='#1B5E20', arrowprops=dict(arrowstyle='->', color='#1B5E20', lw=1.5), zorder=7)
    ax.axvspan(420, 550, alpha=0.12, color='red', zorder=1, linestyle='--')
    ax.text(485, 0.02, 'cluster /o/', ha='center', va='bottom', fontsize=7, color='#C62828', fontstyle='italic', transform=ax.get_xaxis_transform())
    ax.set_xlim(100,mf); ax.set_ylim(0,1.25)
    ax.set_xlabel("Fréquence (Hz)", fontsize=10, fontweight='bold')
    ax.set_ylabel("Importance relative du formant", fontsize=10, fontweight='bold')
    ax.set_title(f"{display}  —  Formants spectraux F1–F6 (ordinario, N={n})", fontsize=12, fontweight='bold', color=family_color, pad=12)
    ax.set_xscale('log')
    ticks = [t for t in [100,150,200,300,400,500,600,800,1000,1500,2000,3000,4000,5000,6000] if t<=mf]
    ax.set_xticks(ticks); ax.set_xticklabels([str(t) for t in ticks], fontsize=8)
    ax.get_xaxis().set_major_formatter(matplotlib.ticker.ScalarFormatter())
    ax.set_yticks([])
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    le = [mpatches.Patch(facecolor=FC[i], alpha=FA[i], edgecolor='#333', label=f'F{i+1} = {formants[i]} Hz') for i,f in valid]
    if fp and abs(fp-f1)>30:
        le.append(Line2D([0],[0], marker='D', color='w', markerfacecolor='#1B5E20', markeredgecolor='black', markersize=10, label=f'Fp centroïde = {fp} Hz'))
    ax.legend(handles=le, loc='upper right', fontsize=7, framealpha=0.9, edgecolor='#CCC')
    ax.text(0.01, -0.08, f"Source : CSV v22 (SOL2020 + Yan_Adds)", transform=ax.transAxes, fontsize=7, color='#888')
    plt.tight_layout()
    out = os.path.join(OUT_IMG, f"{filename}.png")
    fig.savefig(out, dpi=150, bbox_inches='tight', facecolor='white'); plt.close(fig)
    return f"media/{filename}.png"

def fmt_hz(v): return '—' if v == 0 else f"{v:,}".replace(',', ' ')

def tech_table(inst):
    techs = sorted([(t, DATA[(i,t)]) for (i,t) in DATA if i == inst], key=lambda x: x[0])
    if not techs: return ""
    rows = []
    for tech, r in techs:
        fs = [round(sf(r[f'F{i}_hz'])) for i in range(1,7)]
        is_ord = 'ordinario' in tech.lower() and 'to_' not in tech
        bg = ' style="background-color:#dff0d8;"' if is_ord else ''
        rows.append(f'<tr><td{bg}><b>{tech}</b></td><td{bg}>{r["n_samples"]}</td>{"".join(f"<td{bg}>{fmt_hz(f)}</td>" for f in fs)}</tr>')
    return f'<table class="tech-table"><tr class="header"><th>Technique</th><th>N</th><th>F1</th><th>F2</th><th>F3</th><th>F4</th><th>F5</th><th>F6</th></tr>{"".join(rows)}</table>'

DESC = {
    'Sax_Alto': "Son chaud et expressif, grande flexibilité dynamique. F1=398 Hz dans la zone /o/ (plénitude). Le tuyau conique du saxophone produit un spectre riche en harmoniques pairs et impairs. Fp centroïde à 1440 Hz. Giesler note la « o-ähnlicher Vokalfärbung » (coloration vocalique en /o/) partagée avec le basson, le cor et le trombone.",
    # Cuivres
    'Horn': "Son rond et chaleureux, emblématique de la noblesse orchestrale. F1=388 Hz dans la zone /o/ (plénitude). Giesler : « u-ähnlich, hauptsächlich o ». Le cor possède le spectre le plus homogène des cuivres.",
    'Horn+sordina': "Son voilé et lointain. F1 descend de 388 à 344 Hz, la sourdine déplace les formants vers le grave avec une légère nasalité. Utilisé pour les effets d'écho et de distance.",
    'Trumpet_C': "Son brillant et incisif, grande projection. F1=786 Hz dans la zone /a/ (puissance). Fp centroïde à 1046 Hz est remarquablement stable (σ=98) malgré la très forte variabilité de F2 (σ=1018).",
    'Trumpet_C+sordina_straight': "Son piquant et nasal, le plus utilisé en orchestre. Renforce la zone 1000–2000 Hz, donnant un caractère « pointu » et projeté.",
    'Trumpet_C+sordina_cup': "Son doux et arrondi, perd la brillance caractéristique de la trompette. Timbre proche du bugle, utile pour les passages lyriques.",
    'Trumpet_C+sordina_wah_open': "Tige insérée, position ouverte : son nasal et brillant, voyelle /a/. Plus de projection que la position fermée.",
    'Trumpet_C+sordina_wah_closed': "Tige insérée, position fermée : son très étouffé, voyelle /u/. Spectre fortement filtré, caractère introverti.",
    'Trumpet_C+sordina_harmon': "Son « miles-davisien » : fondamentale quasi pure avec un buzz harmonique discret. Timbre très intime, sans tige (stem out).",
    'Trombone': "Son plein et puissant, grande projection. F1=237 Hz (zone /u/) avec Fp centroïde à 1218 Hz. Le trombone couvre un large spectre.",
    'Trombone+sordina_straight': "Son nasal et métallique. La sourdine straight comprime le spectre et renforce la zone 800–1500 Hz, donnant un caractère incisif.",
    'Trombone+sordina_cup': "Son voilé et sombre, projection réduite. La sourdine cup absorbe les harmoniques aigus, timbre mat et feutré.",
    'Trombone+sordina_wah_open': "Position ouverte : son brillant et nasal, spectre riche. Caractère expressif.",
    'Trombone+sordina_wah_closed': "Position fermée : son très étouffé et nasal. F1 remonte à 398 Hz, formants comprimés.",
    'Trombone+sordina_harmon': "Son très concentré, quasi sinusoïdal dans le grave (F1=162 Hz). Timbre « de jazz » intime, projection très directionnelle.",
    'Bass_Tuba': "Son profond et rond, fondamental de l'orchestre. F1=226 Hz (voyelle /u/). Le tuba basse ancre le registre grave avec une couleur sombre et enveloppante.",
    'Bass_Tuba+sordina': "Son assourdi et compact. F1 reste à 226 Hz mais la projection est réduite. Timbre plus mat et concentré.",
    'Bass_Trombone': "Son profond et puissant, plus sombre que le ténor. F1=258 Hz (zone /u/). Élargit l'assise grave de la section cuivres.",
    'Contrabass_Tuba': "Son extrêmement grave et massif. F1=226 Hz identique au tuba basse, F2=463 Hz confirme sa place dans le cluster 450–502 Hz.",
}

# ═══════════════════════════════════════════
# INSTRUMENT DEFINITIONS — in requested order
# ═══════════════════════════════════════════

SAX = [
    ('Sax_Alto', 'Saxophone alto', 'sax_alto', 'ordinario', '#E65100'),
]

CUIVRES = [
    ('Horn',     'Cor en Fa',   'cuivres_horn',    'ordinario', '#C62828'),
    ('Horn+sordina', 'Cor — sourdine', 'cuivres_horn_sord', 'ordinario', '#C62828'),
    ('Trumpet_C', 'Trompette en Do', 'cuivres_trumpet', 'ordinario', '#C62828'),
    ('Trumpet_C+sordina_straight', 'Trompette — sourdine straight', 'cuivres_tpt_straight', 'ordinario', '#C62828'),
    ('Trumpet_C+sordina_cup', 'Trompette — sourdine cup', 'cuivres_tpt_cup', 'ordinario', '#C62828'),
    ('Trumpet_C+sordina_wah', 'Trompette — sourdine wah open', 'cuivres_tpt_wah_open', 'ordinario_open', '#C62828'),
    ('Trumpet_C+sordina_wah', 'Trompette — sourdine wah closed', 'cuivres_tpt_wah_closed', 'ordinario_closed', '#C62828'),
    ('Trumpet_C+sordina_harmon', 'Trompette — sourdine harmon', 'cuivres_tpt_harmon', 'ordinario', '#C62828'),
    ('Trombone', 'Trombone ténor', 'cuivres_trombone', 'ordinario', '#C62828'),
    ('Trombone+sordina_straight', 'Trombone — sourdine straight', 'cuivres_trb_straight', 'ordinario', '#C62828'),
    ('Trombone+sordina_cup', 'Trombone — sourdine cup', 'cuivres_trb_cup', 'ordinario', '#C62828'),
    ('Trombone+sordina_wah', 'Trombone — sourdine wah open', 'cuivres_trb_wah_open', 'ordinario_open', '#C62828'),
    ('Trombone+sordina_wah', 'Trombone — sourdine wah closed', 'cuivres_trb_wah_closed', 'ordinario_closed', '#C62828'),
    ('Trombone+sordina_harmon', 'Trombone — sourdine harmon', 'cuivres_trb_harmon', 'ordinario', '#C62828'),
    ('Bass_Tuba', 'Tuba basse', 'cuivres_tuba', 'ordinario', '#C62828'),
    ('Bass_Tuba+sordina', 'Tuba basse — sourdine', 'cuivres_tuba_sord', 'ordinario', '#C62828'),
    ('Bass_Trombone', 'Trombone basse', 'cuivres_bass_trombone', 'ordinario', '#C62828'),
    ('Contrabass_Tuba', 'Tuba contrebasse', 'cuivres_contrabass_tuba', 'ordinario', '#C62828'),
]

# Generate all graphs
all_info = {}
for section_list in [SAX, CUIVRES]:
    for csv_name, display, gfx, tech, fcolor in section_list:
        d = get_f(csv_name, tech)
        if not d:
            print(f"  ⚠ MANQUANT: {csv_name}/{tech}")
            continue
        fp = FP.get(csv_name.split('+')[0])  # Fp from base instrument
        img = make_graph(display, gfx, d['n'], d['F'], fp, fcolor)
        # Use csv_name for sordina variants desc key
        desc_key = csv_name
        if 'wah' in gfx:
            desc_key = csv_name + ('_open' if 'open' in gfx else '_closed')
        all_info[gfx] = {'csv': csv_name, 'display': display, 'tech': tech,
                          'data': d, 'fp': fp, 'img': img, 'desc_key': desc_key}
        print(f"  ✓ {display:<40s} N={d['n']:>4d}  F1={d['F'][0]:>5d}  → {img}")

# ═══════════════════════════════════════════
# BUILD HTML
# ═══════════════════════════════════════════

CSS = """
  body { font-family: 'Segoe UI', Helvetica, Arial, sans-serif; max-width: 1100px; margin: 0 auto; padding: 20px; background: #fafafa; color: #333; }
  h1 { color: #1a237e; border-bottom: 3px solid; padding-bottom: 10px; }
  h1.sax { border-color: #e65100; } h1.cuivres { border-color: #c62828; }
  h2 { margin-top: 40px; border-left: 4px solid; padding-left: 12px; }
  h2.sax { color: #e65100; border-color: #e65100; } h2.cuivres { color: #c62828; border-color: #c62828; }
  h3 { margin-top: 30px; }
  .instrument-card { background: white; border: 1px solid #ddd; border-radius: 8px; padding: 20px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
  .formant-graph { max-width: 100%; border: 1px solid #eee; border-radius: 4px; }
  .description { font-style: italic; color: #555; background: #fff3e0; padding: 10px; border-left: 3px solid #ff9800; margin: 10px 0; }
  .desc-cuivres { background: #fce4ec; border-left-color: #e91e63; }
  .fp-note { color: #1b5e20; font-weight: bold; }
  .tech-table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 0.9em; }
  .tech-table th, .tech-table td { border: 1px solid #ccc; padding: 6px 10px; text-align: center; }
  .tech-table .header th { background: #1a3a5c; color: white; }
  .tech-table tr:nth-child(even) { background: #f8f8f8; }
  .source-note { font-size: 0.85em; color: #888; margin-top: 30px; border-top: 1px solid #ddd; padding-top: 10px; }
  .section-intro { padding: 15px; border-radius: 6px; margin: 15px 0; }
  .intro-sax { background: #fff3e0; } .intro-cuivres { background: #fce4ec; }
  .yan-badge { display: inline-block; background: #ff9800; color: white; padding: 2px 8px; border-radius: 10px; font-size: 0.8em; margin-left: 8px; }
  .sordina-badge { display: inline-block; background: #7b1fa2; color: white; padding: 2px 8px; border-radius: 10px; font-size: 0.8em; margin-left: 8px; }
"""

def card(gfx_key, show_techs=True, desc_class='description'):
    info = all_info.get(gfx_key)
    if not info: return f"<!-- MISSING: {gfx_key} -->"
    d = info['data']
    desc_key = info['desc_key']
    desc = DESC.get(desc_key, '')
    is_yan = info['csv'] in ('Bass_Trombone','Contrabass_Tuba')
    is_sord = 'sordina' in info['csv'] or 'sord' in gfx_key
    badge = ''
    if is_yan: badge = '<span class="yan-badge">Yan_Adds</span>'
    if is_sord: badge += '<span class="sordina-badge">sourdine</span>'
    fp_html = f'<p class="fp-note">Fp centroïde = {info["fp"]} Hz</p>' if info['fp'] and abs(info['fp']-d['F'][0])>30 else ''
    desc_html = f'<p class="{desc_class}">{desc}</p>' if desc else ''
    tt = tech_table(info['csv']) if show_techs else ''
    return f'<div class="instrument-card"><h3>{info["display"]}{badge}</h3><img src="{info["img"]}" class="formant-graph"/>{desc_html}{fp_html}{tt}</div>'

# ── SAXOPHONE HTML ──
html_sax = f"""<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8">
<title>Référence Formantique — Section Saxophones</title>
<style>{CSS}</style></head><body>
<h1 class="sax">IV. Les Saxophones</h1>
<div class="section-intro intro-sax">
  <p><strong>Plage formantique :</strong> 398–1 798 Hz (voyelles /o/ → /e/).</p>
  <p><strong>Caractéristique :</strong> tuyau conique + anche simple = spectre riche en harmoniques pairs et impairs.
  Le saxophone alto partage la « coloration vocalique en /o/ » (Giesler) avec le basson, le cor et le trombone.</p>
  <p><strong>Note :</strong> les Saxophone ténor et Saxophone baryton ne sont pas encore disponibles dans la base de données specenv.</p>
</div>
{card('sax_alto', show_techs=True, desc_class='description')}
<p class="source-note"><strong>Source :</strong> CSV v22 — pipeline validé Δ=0 Hz.</p>
</body></html>"""

with open(os.path.join(OUT_HTML, "section_saxophones.html"), 'w', encoding='utf-8') as f:
    f.write(html_sax)
print(f"\n✓ section_saxophones.html")

# ── CUIVRES HTML ──
html_cuivres = f"""<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8">
<title>Référence Formantique — Section Cuivres</title>
<style>{CSS}</style></head><body>
<h1 class="cuivres">V. Les Cuivres</h1>
<div class="section-intro intro-cuivres">
  <p><strong>Plage formantique :</strong> 162–2 358 Hz (voyelles /u/ → /i/).</p>
  <p><strong>Caractéristique :</strong> formants bien définis, grande stabilité inter-dynamiques (sauf trompette).
  Le <em>cluster de convergence</em> 450–502 Hz (zone /o/) rassemble Cor, Trombone, Basson, Violoncelle,
  Cor anglais et Saxophone ténor.</p>
</div>

<h2 class="cuivres">Cor</h2>
{card('cuivres_horn', desc_class='desc-cuivres')}
{card('cuivres_horn_sord', show_techs=False, desc_class='desc-cuivres')}

<h2 class="cuivres">Trompette</h2>
{card('cuivres_trumpet', desc_class='desc-cuivres')}
{card('cuivres_tpt_straight', show_techs=False, desc_class='desc-cuivres')}
{card('cuivres_tpt_cup', show_techs=False, desc_class='desc-cuivres')}
{card('cuivres_tpt_wah_open', show_techs=False, desc_class='desc-cuivres')}
{card('cuivres_tpt_wah_closed', show_techs=False, desc_class='desc-cuivres')}
{card('cuivres_tpt_harmon', show_techs=False, desc_class='desc-cuivres')}

<h2 class="cuivres">Trombone</h2>
{card('cuivres_trombone', desc_class='desc-cuivres')}
{card('cuivres_trb_straight', show_techs=False, desc_class='desc-cuivres')}
{card('cuivres_trb_cup', show_techs=False, desc_class='desc-cuivres')}
{card('cuivres_trb_wah_open', show_techs=False, desc_class='desc-cuivres')}
{card('cuivres_trb_wah_closed', show_techs=False, desc_class='desc-cuivres')}
{card('cuivres_trb_harmon', show_techs=False, desc_class='desc-cuivres')}

<h2 class="cuivres">Tubas</h2>
{card('cuivres_tuba', desc_class='desc-cuivres')}
{card('cuivres_tuba_sord', show_techs=False, desc_class='desc-cuivres')}
{card('cuivres_bass_trombone', show_techs=True, desc_class='desc-cuivres')}
{card('cuivres_contrabass_tuba', show_techs=True, desc_class='desc-cuivres')}

<p class="source-note"><strong>Source :</strong> CSV v22 — pipeline validé Δ=0 Hz. Sourdines : SOL2020. Trombone basse, Tuba contrebasse : Yan_Adds.</p>
</body></html>"""

with open(os.path.join(OUT_HTML, "section_cuivres.html"), 'w', encoding='utf-8') as f:
    f.write(html_cuivres)
print(f"✓ section_cuivres.html (remplacé)")

print(f"\nTotal graphiques: {len(all_info)}")
