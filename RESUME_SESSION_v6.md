# Résumé de session — v5/v6 (20–21 mars 2026)

## Contexte

Repo: `https://github.com/zseramnay/Formants` (user: zseramnay)
Document en ligne: `https://zseramnay.github.io/Formants/Versions-html-and-docx/REFERENCE_FORMANTIQUE_COMPLETE.html`
Deux audits externes (Gemini + Copilot) analysés. Plan v5 défini et exécuté (points A–F), puis améliorations majeures v6 (amplitudes, bandwidths, graphiques en cloche).

## Points A–F du plan v5 (complétés)

**Point A** — Note terminologique : encadré bleu « F1–F6 = pics spectraux, pas formants vocaux (Fant 1960) ».

**Point B** — Clause "sustained ordinario only" : encadré jaune intro + rappels dans les 4 sections.

**Point C** — CSV v3 avec statistiques :
- `Resultats/formants_all_techniques_v3.csv` — 487 lignes, 39 colonnes
- `Resultats/formants_yan_adds_v3.csv` — 46 lignes, 39 colonnes
- Colonnes ajoutées : F1–F3 (mean, std, q25, q75) + F4–F6 (mean, std) + F1–F6 (db amplitude médiane)
- Validation F1_hz Δ=0 vs v2

**Point D** — Fenêtrage Hann : clarifié (fait en amont par IRCAM SuperVP). Table des paramètres corrigée.

**Point E** — Axes Bark secondaires sur fig1 et fig3 (formule Traunmüller).

**Point F** — σ(Fp) comme proxy de bandwidth/Q : encadré violet.

## Améliorations v6

### Amplitudes réelles (dB)

Le CSV v3 contient maintenant 6 colonnes `F1_db`–`F6_db` (médiane des amplitudes spectrales en dB par formant). Les graphiques reflètent l'amplitude réelle normalisée au lieu de l'ancienne « importance relative » artificielle (F1=1.0, F2=0.88...).

### Bandwidths -3dB réels depuis specenv

**Script** : calcul des BW à −3 dB pour chaque pic dans les enveloppes spectrales (specenv), médiane par instrument×technique.

**`Resultats/bandwidths_3db.csv`** — 533 profils instrument×technique, colonnes F1_bw–F6_bw.

Exemple comparatif Cor en Fa ordinario :
- σ statistique (variabilité position du pic) : F3=297 Hz, F6=608 Hz — absurde pour un formant
- BW -3dB réel (largeur de la résonance) : F3=118 Hz, F6=301 Hz — cohérent

### Graphiques en courbes de cloche

**`make_graph` réécrit** : gaussiennes colorées au lieu de barres rectangulaires.
- Hauteur = amplitude dB normalisée
- Largeur = BW_3dB / 2.355 (conversion en σ gaussien)
- Fallback si pas de BW : σ = 8% de la fréquence centrale
- Titre : « Profil formantique » au lieu de « Formants spectraux F1–F6 »

### Fp dynamique pour tous les instruments

`compute_fp_from_specenv()` cherche maintenant dans SOL2020 ET Yan_Adds (fallback avec conversion `_`→`-` et `+`→`-` pour les noms de fichiers).

Instruments qui ont maintenant leur Fp (manquaient avant) :
- Petite flûte : 1320 Hz
- Alto : 1300, Alto+sourdine : 1125, Alto+piombo : 1171
- Violoncelle : 1242, Violoncelle+sourdine : 1238, Violoncelle+piombo : 1158
- Contrebasse : 1235, Contrebasse+sourdine : 1217
- Violon+sourdine : 1218, Violon+piombo : 1345
- Ensemble de violons : 1302, +sourdine : 1208
- Ensemble d'altos : 1317, +sourdine : 1220
- Ensemble de violoncelles : 1242, +sourdine : 1219
- Ensemble de contrebasses : 1175
- Toutes les sourdines cuivres (cor, trompette, trombone, tuba)

Fp TOUJOURS affiché (suppression du seuil `abs(Fp-F1) > 30 Hz`), losange réduit (markersize 14→7), pas de flèche.

### Anti-collision des labels (v6)

Algorithme en deux passes :
1. **PASS 1 — vertical uniquement** (centre x) : essaie 10 positions au-dessus du pic, en escalier
2. **PASS 2 — horizontal** (seulement si vertical échoue) : petits décalages ±0.6×LABEL_W

Détection de densité spectrale : quand les 6 formants tiennent dans <1.5 octaves (ex: trompette harmon), police réduite automatiquement (5.5–6pt) et hitboxes réduites de 25%.

ylim élargi à 1.55 pour laisser de l'espace au-dessus des pics. Labels F4–F6 en format compact 2 lignes. Flèches pointillées pour les labels déplacés.

**Résultat : 49/49 images sans collision, vérifié systématiquement.**

### Note méthodologique sourdines

Encadré jaune dans la section « Cuivres avec sourdine » (HTML + DOCX) : explique que les sourdines créent des résonances propres qui dominent le spectre, que F1 reflète cette résonance et pas la fondamentale, et que Fp est plus fiable dans ces cas. Analyse harmon détaillée.

### Bugs corrigés

1. `build_document_complet.py` pointait vers `v4-html-docx-enriched/` → corrigé vers `v5-html-docx/`
2. `NAME_MAPPING` : `Horn`→`Cor` → `Cor en Fa`, `Trombone`→`Trombone` → `Trombone ténor`, sourdines trompette corrigées
3. Fichier Violin+sordina_piombo specenv récupéré (manquait dans le dossier per-instrument)
4. Légende déplacée de upper-right (cachait les données) vers lower-left (zone /u/ toujours vide)

## État du repo

```
Scripts/
  extract_formants_all_techniques_v2_fixed.py   ← baseline validée (v22)
  extract_formants_all_techniques_v3_stats.py   ← v3 avec mean/std/q25/q75/db
  v4-html-docx-enriched/                        ← scripts v4 (intacts)
  v5-html-docx/                                 ← scripts v5/v6 (actifs)
    common.py                                   ← charge CSV v3 + BW, make_graph gaussien, compute_fp
    build_bois_html_docx.py                     ← Fp dynamique
    build_cuivres_html_docx.py                  ← Fp dynamique + note sourdines
    build_sax_html_docx.py
    build_cordes_html_docx.py                   ← Fp dynamique
    build_synthese_html_docx.py
    build_intro_html_docx.py
    build_document_complet.py                   ← pointe vers v5-html-docx/
    make_synthese_figures.py                    ← fig1/fig3 avec axes Bark

Resultats/
  formants_all_techniques_v3.csv                ← 487 lignes, 39 colonnes (F1-F6 hz+db+stats) ★
  formants_yan_adds_v3.csv                      ← 46 lignes, 39 colonnes ★
  bandwidths_3db.csv                            ← 533 profils, BW -3dB depuis specenv ★
```

## Consignes pour la suite

- **Toujours montrer une image prototype avant de régénérer les 49 images**
- **Toujours demander l'aval avant build_document_complet et push**
- **Vérifier systématiquement les 49 images** (collision + Fp) avant de proposer le push
- Les fichiers specenv sont dans `Data/FullSOL2020_specenv par instrument/` et `Data/Yan_Adds-Divers_specenv par instrument/`

## Workflow git

```bash
git clone https://github.com/zseramnay/Formants.git
cd Formants
git config user.name "Claude" && git config user.email "claude@anthropic.com"
# Pour push : PAT fourni par Yan, nettoyé immédiatement après
git remote set-url origin https://[PAT]@github.com/zseramnay/Formants.git
git push origin main
git remote set-url origin https://github.com/zseramnay/Formants.git
```

## Ce qu'on ne change PAS

- Pas de renommage F1→P1/P2 (note terminologique suffit)
- Pas de transitions formantiques/diphthongaison (hors scope, sustained only)
- Pas de normalisation morphologique (non pertinent pour instruments)
