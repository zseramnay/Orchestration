#!/usr/bin/env python3
"""
Build Section Cordes — HTML + graphs from CSV v22
Solistes + sourdines + sourdines lourdes + ensembles
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
with open('/mnt/user-data/uploads/formants_all_techniques.csv','r') as f:
    for row in csv.DictReader(f): DATA[(row['instrument'],row['technique'])] = row
with open('/home/claude/formants_yan_adds.csv','r') as f:
    for row in csv.DictReader(f): DATA[(row['instrument'],row['technique'])] = row

def get_f(inst, tech):
    r = DATA.get((inst, tech))
    if not r: return None
    return {'n': int(r['n_samples']), 'F': [round(sf(r[f'F{i}_hz'])) for i in range(1,7)]}

# ─── Graph engine ───
VZ = [(100,400,'#DCEEFB','u (oo)\nProfondeur'),(400,600,'#D5ECD5','o (oh)\nPlénitude'),
      (600,800,'#FDE8CE','å (aw)\nTransition'),(800,1250,'#F8D5D5','a (ah)\nPuissance'),
      (1250,2600,'#E8D5F0','e (eh)\nClarté'),(2600,6000,'#FFF8D0','i (ee)\nBrillance')]
FC=['#D32F2F','#E64A19','#F57C00','#FFA000','#FBC02D','#CDDC39']
FA=[1.0,0.85,0.7,0.55,0.4,0.3]

def make_graph(display, filename, n, formants, fp=None):
    valid = [(i,f) for i,f in enumerate(formants) if f>0]
    if not valid: return None
    mf = min(max(max(f for _,f in valid)+500,3000),6500)
    fig, ax = plt.subplots(figsize=(9.6,4.8),dpi=150)
    for lo,hi,c,l in VZ:
        if lo<mf:
            ax.axvspan(lo,min(hi,mf),alpha=0.35,color=c,zorder=0)
            mid=(lo+min(hi,mf))/2
            if mid<mf*0.95:
                ax.text(mid,0.97,l,ha='center',va='top',fontsize=7,color='#666',fontweight='bold',transform=ax.get_xaxis_transform())
    bw=mf*0.012
    for i,freq in valid:
        bh=1.0*(1.0-i*0.12)
        ax.bar(freq,bh,width=bw*(1.2-i*0.05),color=FC[i],alpha=FA[i],edgecolor='#333',linewidth=0.8,zorder=3)
        ax.text(freq,bh+0.03,f"F{i+1}\n{freq} Hz",ha='center',va='bottom',fontsize=8,fontweight='bold',color='#333',zorder=5)
    f1=formants[0]
    if fp and abs(fp-f1)>30:
        ax.plot(fp,0.5,marker='D',markersize=14,color='#1B5E20',markeredgecolor='black',markeredgewidth=1.5,zorder=6)
        ax.annotate(f"Fp = {fp} Hz\n(centroïde)",xy=(fp,0.5),xytext=(fp,0.65),ha='center',fontsize=8,fontweight='bold',color='#1B5E20',arrowprops=dict(arrowstyle='->',color='#1B5E20',lw=1.5),zorder=7)
    ax.axvspan(420,550,alpha=0.12,color='red',zorder=1,linestyle='--')
    ax.text(485,0.02,'cluster /o/',ha='center',va='bottom',fontsize=7,color='#C62828',fontstyle='italic',transform=ax.get_xaxis_transform())
    ax.set_xlim(100,mf);ax.set_ylim(0,1.25)
    ax.set_xlabel("Fréquence (Hz)",fontsize=10,fontweight='bold')
    ax.set_ylabel("Importance relative du formant",fontsize=10,fontweight='bold')
    ax.set_title(f"{display}  —  Formants spectraux F1–F6 (ordinario, N={n})",fontsize=12,fontweight='bold',color='#1565C0',pad=12)
    ax.set_xscale('log')
    ticks=[t for t in [100,150,200,300,400,500,600,800,1000,1500,2000,3000,4000,5000,6000] if t<=mf]
    ax.set_xticks(ticks);ax.set_xticklabels([str(t) for t in ticks],fontsize=8)
    ax.get_xaxis().set_major_formatter(matplotlib.ticker.ScalarFormatter())
    ax.set_yticks([])
    for s in ['top','right','left']: ax.spines[s].set_visible(False)
    le=[mpatches.Patch(facecolor=FC[i],alpha=FA[i],edgecolor='#333',label=f'F{i+1} = {formants[i]} Hz') for i,f in valid]
    if fp and abs(fp-f1)>30:
        le.append(Line2D([0],[0],marker='D',color='w',markerfacecolor='#1B5E20',markeredgecolor='black',markersize=10,label=f'Fp centroïde = {fp} Hz'))
    ax.legend(handles=le,loc='upper right',fontsize=7,framealpha=0.9,edgecolor='#CCC')
    ax.text(0.01,-0.08,"Source : CSV v22 (SOL2020 + Yan_Adds)",transform=ax.transAxes,fontsize=7,color='#888')
    plt.tight_layout()
    out=os.path.join(OUT_IMG,f"{filename}.png")
    fig.savefig(out,dpi=150,bbox_inches='tight',facecolor='white');plt.close(fig)
    return f"media/{filename}.png"

def fmt_hz(v): return '—' if v==0 else f"{v:,}".replace(',',' ')

def tech_table(inst):
    techs=sorted([(t,DATA[(i,t)]) for (i,t) in DATA if i==inst],key=lambda x:x[0])
    if not techs: return ""
    rows=[]
    for tech,r in techs:
        n=r['n_samples'];fs=[round(sf(r[f'F{i}_hz'])) for i in range(1,7)]
        is_ord='ordinario' in tech.lower() and 'to_' not in tech
        bg=' style="background-color:#dff0d8;"' if is_ord else ''
        fvals=''.join(f'<td{bg}>{fmt_hz(f)}</td>' for f in fs)
        rows.append(f'<tr><td{bg}><b>{tech}</b></td><td{bg}>{n}</td>{fvals}</tr>')
    return f"""<table class="tech-table"><tr class="header"><th>Technique</th><th>N</th><th>F1</th><th>F2</th><th>F3</th><th>F4</th><th>F5</th><th>F6</th></tr>{''.join(rows)}</table>"""

# ═══════════════════════════════════════════════════════════
# Fp values
# ═══════════════════════════════════════════════════════════
FP = {
    'Violin':893, 'Violin+sordina':853, 'Violin+sordina_piombo':None,
    'Viola':1221, 'Viola+sordina':1102, 'Viola+sordina_piombo':None,
    'Violoncello':1025, 'Violoncello+sordina':909, 'Violoncello+sordina_piombo':None,
    'Contrabass':1324, 'Contrabass+sordina':None,
    'Violin_Ensemble':993, 'Violin_Ensemble+sordina':None,
    'Viola_Ensemble':1004, 'Viola_Ensemble+sordina':None,
    'Violoncello_Ensemble':1469, 'Violoncello_Ensemble+sordina':None,
    'Contrabass_Ensemble':1149,
}

# ═══════════════════════════════════════════════════════════
# Descriptions
# ═══════════════════════════════════════════════════════════
DESC = {
    'Violin': "Son brillant et expressif, grande tessiture. F1=506 Hz dans le cluster /o/. Le Hauptformant de Meyer (800–1200 Hz) se manifeste comme un plateau spectral large : Fp=893 Hz (σ=92), remarquablement stable sur 4 octaves. F2=1518 Hz (zone nasale) est un artefact statistique du peak-picking.",
    'Violin+sordina': "La sourdine abaisse F1 de 506 à 366 Hz et atténue les aigus. Le timbre devient plus doux et voilé, avec moins de projection. Fp descend de 893 à 853 Hz.",
    'Violin+sordina_piombo': "La sourdine lourde (piombo) produit un effet plus radical : F1=344 Hz, son très étouffé et distant. Le spectre est fortement comprimé, timbre quasi irréel.",
    'Viola': "Son chaud et mélancolique, intermédiaire entre violon et violoncelle. F1=377 Hz (zone /o/). Fp=1221 Hz capture la résonance de caisse caractéristique de l'alto. Tessiture expressive dans le médium.",
    'Viola+sordina': "Sourdine : F1 descend de 377 à 344 Hz. Son plus intime et voilé, perte de projection. Fp=1102 Hz.",
    'Viola+sordina_piombo': "Sourdine lourde : F1 chute à 226 Hz. Effet très prononcé, son distant et éthéré, quasi fantomatique.",
    'Violoncello': "Son profond et lyrique, grande richesse harmonique. F1=205 Hz (zone /u/). Fp=1025 Hz (σ=70). Le violoncelle converge remarquablement avec le basson (Δ=3 Hz sur F2) et le cor (Δ=42 Hz).",
    'Violoncello+sordina': "Sourdine : F1 reste à 205 Hz mais F2 monte légèrement (538 vs 506). Son plus mat et contenu. Fp=909 Hz.",
    'Violoncello+sordina_piombo': "Sourdine lourde : F1 descend à 172 Hz. Forte atténuation des harmoniques aigus, son très assourdi et lointain.",
    'Contrabass': "Son grave et puissant, fondation des cordes. F1=172 Hz (zone /u/). Fp=1324 Hz. Faible variabilité de F2 (σ=89 Hz) — les formants de la contrebasse sont très stables.",
    'Contrabass+sordina': "Sourdine : F1 descend à 162 Hz. Son plus mat, perte de résonance. Utilisé pour les passages chambristes.",
    'Violin_Ensemble': "L'ensemble de violons comprime légèrement les formants par rapport au soliste (F1: 495 vs 506 Hz). Le spectre est plus homogène, les irrégularités individuelles s'annulent. Fp=993 Hz.",
    'Violin_Ensemble+sordina': "Ensemble avec sourdines : F1=355 Hz. L'effet de la sourdine est similaire au soliste mais encore plus homogène grâce à la fusion d'ensemble.",
    'Viola_Ensemble': "Ensemble d'altos : F1=366 Hz, proche du soliste (377 Hz). La section d'altos produit un son riche et velouté. Fp=1004 Hz.",
    'Viola_Ensemble+sordina': "Ensemble d'altos avec sourdines : F1=291 Hz. Son très voilé et intime.",
    'Violoncello_Ensemble': "Ensemble de violoncelles : F1=205 Hz, identique au soliste. La section produit un son profond et enveloppant. Fp=1469 Hz.",
    'Violoncello_Ensemble+sordina': "Ensemble de violoncelles avec sourdines : F1 reste à 205 Hz, F2 descend à 463 Hz. Son contenu et mat.",
    'Contrabass_Ensemble': "Ensemble de contrebasses (non-vibrato) : F1=172 Hz. La section produit une fondation grave massive et stable. Fp=1149 Hz.",
}

# ═══════════════════════════════════════════════════════════
# Instrument list — in order
# ═══════════════════════════════════════════════════════════
SOLISTES = [
    ('Violin','Violon','cordes_violin','ordinario'),
    ('Violin+sordina','Violon — sourdine','cordes_violin_sord','ordinario'),
    ('Violin+sordina_piombo','Violon — sourdine lourde','cordes_violin_sord_piombo','ordinario'),
    ('Viola','Alto','cordes_viola','ordinario'),
    ('Viola+sordina','Alto — sourdine','cordes_viola_sord','ordinario'),
    ('Viola+sordina_piombo','Alto — sourdine lourde','cordes_viola_sord_piombo','ordinario'),
    ('Violoncello','Violoncelle','cordes_violoncello','ordinario'),
    ('Violoncello+sordina','Violoncelle — sourdine','cordes_violoncello_sord','ordinario'),
    ('Violoncello+sordina_piombo','Violoncelle — sourdine lourde','cordes_violoncello_sord_piombo','ordinario'),
    ('Contrabass','Contrebasse','cordes_contrabass','ordinario'),
    ('Contrabass+sordina','Contrebasse — sourdine','cordes_contrabass_sord','ordinario'),
]

ENSEMBLES = [
    ('Violin_Ensemble','Ensemble de violons','cordes_violin_ens','ordinario'),
    ('Violin_Ensemble+sordina','Ensemble de violons — sourdine','cordes_violin_ens_sord','ordinario'),
    ('Viola_Ensemble',"Ensemble d'altos",'cordes_viola_ens','ordinario'),
    ('Viola_Ensemble+sordina',"Ensemble d'altos — sourdine",'cordes_viola_ens_sord','ordinario'),
    ('Violoncello_Ensemble','Ensemble de violoncelles','cordes_violoncello_ens','ordinario'),
    ('Violoncello_Ensemble+sordina','Ensemble de violoncelles — sourdine','cordes_violoncello_ens_sord','ordinario'),
    ('Contrabass_Ensemble','Ensemble de contrebasses','cordes_contrabass_ens','non-vibrato'),
]

# ═══════════════════════════════════════════════════════════
# GENERATE ALL
# ═══════════════════════════════════════════════════════════
all_info = {}
for section in [SOLISTES, ENSEMBLES]:
    for csv_name, display, gfx, tech in section:
        d = get_f(csv_name, tech)
        if not d:
            print(f"  ⚠ MANQUANT: {csv_name}/{tech}")
            continue
        fp = FP.get(csv_name)
        img = make_graph(display, gfx, d['n'], d['F'], fp)
        all_info[gfx] = {'csv':csv_name,'display':display,'tech':tech,'data':d,'fp':fp,'img':img}
        print(f"  ✓ {display:<45s} N={d['n']:>4d}  F1={d['F'][0]:>5d}  → {img}")

# ═══════════════════════════════════════════════════════════
# BUILD HTML
# ═══════════════════════════════════════════════════════════
def card(gfx, show_techs=True):
    info = all_info.get(gfx)
    if not info: return ""
    d=info['data']; desc=DESC.get(info['csv'],'')
    is_yan = 'Ensemble' in info['csv']
    badge = '<span class="yan-badge">Yan_Adds</span>' if is_yan else ''
    fp_html = f'<p class="fp-note">Fp centroïde = {info["fp"]} Hz</p>' if info['fp'] else ''
    desc_html = f'<p class="description">{desc}</p>' if desc else ''
    tt = tech_table(info['csv']) if show_techs and '+sordina' not in info['csv'] else ''
    return f"""
    <div class="instrument-card">
      <h3>{info['display']}{badge}</h3>
      <img src="{info['img']}" alt="{info['display']}" class="formant-graph"/>
      {desc_html}{fp_html}{tt}
    </div>"""

html = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Référence Formantique — Section Cordes</title>
<style>
  body { font-family: 'Segoe UI', Helvetica, Arial, sans-serif; max-width: 1100px; margin: 0 auto; padding: 20px; background: #fafafa; color: #333; }
  h1 { color: #1a237e; border-bottom: 3px solid #1565c0; padding-bottom: 10px; }
  h2 { color: #1565c0; margin-top: 40px; border-left: 4px solid #1565c0; padding-left: 12px; }
  h3 { color: #0d47a1; margin-top: 30px; }
  .instrument-card { background: white; border: 1px solid #ddd; border-radius: 8px; padding: 20px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
  .formant-graph { max-width: 100%; border: 1px solid #eee; border-radius: 4px; }
  .description { font-style: italic; color: #555; background: #e3f2fd; padding: 10px; border-left: 3px solid #42a5f5; margin: 10px 0; }
  .fp-note { color: #1b5e20; font-weight: bold; }
  .tech-table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 0.9em; }
  .tech-table th, .tech-table td { border: 1px solid #ccc; padding: 6px 10px; text-align: center; }
  .tech-table .header th { background: #1a3a5c; color: white; }
  .tech-table tr:nth-child(even) { background: #f8f8f8; }
  .source-note { font-size: 0.85em; color: #888; margin-top: 30px; border-top: 1px solid #ddd; padding-top: 10px; }
  .section-intro { background: #e3f2fd; padding: 15px; border-radius: 6px; margin: 15px 0; }
  .yan-badge { display: inline-block; background: #ff9800; color: white; padding: 2px 8px; border-radius: 10px; font-size: 0.8em; margin-left: 8px; }
  .compare-table { width: 100%; border-collapse: collapse; margin: 15px 0; }
  .compare-table th, .compare-table td { border: 1px solid #ccc; padding: 8px 12px; text-align: center; }
  .compare-table th { background: #1565c0; color: white; }
  .compare-table tr:nth-child(even) { background: #e3f2fd; }
</style>
</head>
<body>

<h1>VI. Les Cordes</h1>

<div class="section-intro">
  <p><strong>Plage formantique :</strong> 162–1 518 Hz (voyelles /u/ → /e/).</p>
  <p><strong>Caractéristique :</strong> les cordes frottées possèdent des résonances de caisse larges
  (plateaux spectraux) que le peak-picking ne capture pas toujours correctement. C'est pour les cordes
  que le Fp centroïde apporte le gain de stabilité le plus significatif — notamment pour le violon 
  (Fp=893 Hz, σ=92, vs F2=1518, σ=651).</p>
  <p><strong>Cluster /o/ :</strong> le violoncelle (F2=506 Hz) et l'alto (F1=377 Hz) sont proches du 
  cluster 450–502 Hz, facilitant leur fusion avec les cuivres (cor, trombone) et les bois (basson).</p>
</div>

<h2>Cordes solistes</h2>
"""

# Violon group
html += "<h3 style='color:#1565c0; font-size:1.3em; margin-top:30px;'>Violon</h3>"
for k in ['cordes_violin', 'cordes_violin_sord', 'cordes_violin_sord_piombo']:
    html += card(k, k == 'cordes_violin')

# Alto group
html += "<h3 style='color:#1565c0; font-size:1.3em; margin-top:30px;'>Alto</h3>"
for k in ['cordes_viola', 'cordes_viola_sord', 'cordes_viola_sord_piombo']:
    html += card(k, k == 'cordes_viola')

# Violoncelle group
html += "<h3 style='color:#1565c0; font-size:1.3em; margin-top:30px;'>Violoncelle</h3>"
for k in ['cordes_violoncello', 'cordes_violoncello_sord', 'cordes_violoncello_sord_piombo']:
    html += card(k, k == 'cordes_violoncello')

# Contrebasse group
html += "<h3 style='color:#1565c0; font-size:1.3em; margin-top:30px;'>Contrebasse</h3>"
for k in ['cordes_contrabass', 'cordes_contrabass_sord']:
    html += card(k, k == 'cordes_contrabass')

# ═══════════════════════════════════════════════════════════
# ENSEMBLES
# ═══════════════════════════════════════════════════════════
html += """
<h2>Ensembles de cordes</h2>

<div class="section-intro">
  <p><strong>Comparaison soliste vs. ensemble :</strong> l'effet de section comprime légèrement les 
  formants et homogénéise le spectre. Les irrégularités individuelles s'annulent, produisant un 
  timbre plus lisse et plus fondu.</p>
</div>
"""

# Comparison table
html += """<h3>Comparaison soliste vs. ensemble (F1-F4, ordinario)</h3>
<table class="compare-table">
<tr><th>Instrument</th><th>F1 solo</th><th>F1 ens.</th><th>Δ</th><th>F2 solo</th><th>F2 ens.</th><th>Δ</th></tr>
"""
comparisons = [
    ('Violon', 'Violin', 'Violin_Ensemble'),
    ('Alto', 'Viola', 'Viola_Ensemble'),
    ('Violoncelle', 'Violoncello', 'Violoncello_Ensemble'),
    ('Contrebasse', 'Contrabass', 'Contrabass_Ensemble'),
]
for name, solo_csv, ens_csv in comparisons:
    s = get_f(solo_csv, 'ordinario')
    e = get_f(ens_csv, 'ordinario') or get_f(ens_csv, 'non-vibrato')
    if s and e:
        d1 = abs(s['F'][0]-e['F'][0]); d2 = abs(s['F'][1]-e['F'][1])
        html += f"<tr><td><b>{name}</b></td><td>{s['F'][0]}</td><td>{e['F'][0]}</td><td>{d1}</td><td>{s['F'][1]}</td><td>{e['F'][1]}</td><td>{d2}</td></tr>\n"
html += "</table>\n"

for k in ['cordes_violin_ens','cordes_violin_ens_sord',
          'cordes_viola_ens','cordes_viola_ens_sord',
          'cordes_violoncello_ens','cordes_violoncello_ens_sord',
          'cordes_contrabass_ens']:
    html += card(k, 'sord' not in k)

html += """
<p class="source-note">
  <strong>Source des données :</strong> formants_all_techniques.csv (SOL2020) + formants_yan_adds.csv (ensembles) — pipeline v22 validé.<br/>
  <strong>Ensembles :</strong> Yan_Adds (Violin-Ensemble, Viola-Ensemble, Violoncello-Ensemble, Contrabass-Ensemble).<br/>
  <strong>Sourdines lourdes (piombo) :</strong> SOL2020 — effet plus prononcé que la sourdine standard.
</p>
</body>
</html>"""

with open(os.path.join(OUT_HTML, "section_cordes.html"), 'w', encoding='utf-8') as f:
    f.write(html)

print(f"\n{'='*60}")
print(f"HTML: section_cordes.html")
print(f"Graphiques: {len(all_info)}")

