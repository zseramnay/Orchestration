# Résumé de session — v7 (20–21 mars 2026)

## Contexte

Repo: `https://github.com/zseramnay/Formants` (user: zseramnay)
Document en ligne: `https://zseramnay.github.io/Formants/Versions-html-and-docx/REFERENCE_FORMANTIQUE_COMPLETE.html`
Deux audits externes (Gemini + Copilot) analysés. Plan v5 défini et exécuté (points A–F), puis améliorations majeures v6/v7.

## Points A–F du plan v5 (complétés)

**Point A** — Note terminologique : encadré bleu « F1–F6 = pics spectraux, pas formants vocaux (Fant 1960) ».
**Point B** — Clause "sustained ordinario only" : encadré jaune intro + rappels dans les 4 sections.
**Point C** — CSV v3 avec statistiques (487 + 46 lignes, 39 colonnes).
**Point D** — Fenêtrage Hann : clarifié (fait en amont par IRCAM SuperVP).
**Point E** — Axes Bark secondaires sur fig1 et fig3 (formule Traunmüller).
**Point F** — σ(Fp) comme proxy de bandwidth/Q : encadré violet.

## Améliorations v6

### Bandwidths -3dB réels depuis specenv

`Resultats/bandwidths_3db.csv` — 533 profils instrument×technique, colonnes F1_bw–F6_bw.

### Graphiques en courbes de cloche

`make_graph` réécrit : gaussiennes colorées, hauteur = amplitude dB normalisée, largeur = BW_3dB / 2.355.

### Fp dynamique pour TOUS les instruments

`compute_fp_from_specenv()` cherche dans SOL2020 ET Yan_Adds (fallback `_`→`-` et `+`→`-`).
Fp TOUJOURS affiché (suppression du seuil `abs(Fp-F1) > 30 Hz`), losange réduit (markersize 7), pas de flèche.

### Anti-collision des labels (v6)

Algorithme en deux passes :
1. PASS 1 — vertical uniquement (centre x) : essaie 10 positions en escalier au-dessus du pic
2. PASS 2 — horizontal (seulement si vertical échoue) : petits décalages ±0.6×LABEL_W

Détection de densité spectrale : quand les 6 formants < 1.5 octaves (ex: trompette harmon), police réduite (5.5–6pt), hitboxes −25%.

ylim élargi à 1.55. Labels F4–F6 en format compact 2 lignes. **49/49 images sans collision.**

### Note méthodologique sourdines

Encadré jaune dans « Cuivres avec sourdine ». Analyse harmon détaillée.

## Améliorations v7

### Amplitudes dB recalculées depuis l'enveloppe spectrale moyenne

**Problème identifié** : les colonnes `F1_db`–`F6_db` du CSV v3 utilisaient la médiane des amplitudes de pics **par échantillon individuel**. Pour les instruments à grande tessiture (violon : G3–E7, 4 octaves), cela donnait des résultats contre-intuitifs : F5 plus fort que F1 au violon, F2 plus fort que F1 au trombone, etc.

**Cause** : sur une note aiguë (ex: Violon A6=1760 Hz), la fondamentale tombe dans la zone bridge hill (3500 Hz, −1.8 dB) tandis que la résonance de caisse à 400 Hz est à −20 dB. Le dataset SOL2020 contient majoritairement des notes aiguës (261 >800 Hz vs 9 <400 Hz pour le violon), ce qui biaisait la médiane.

**Correction** : les dB sont maintenant lus directement sur l'**enveloppe spectrale moyenne** (toutes notes confondues) à la fréquence de chaque formant. Résultat : les profils formantiques et les enveloppes spectrales (section VII) montrent les mêmes rapports d'amplitude.

468 profils mis à jour sur 533 (les skips = guitare, harpe, accordéon, percussions — pas de specenv).

**Limitation identifiée** : la moyenne globale sur toute la tessiture « dilue » le grave et « gonfle » l'aigu. Pour certains instruments (trombone, flûte, clarinette), F2 reste plus fort que F1 sur l'enveloppe moyenne — c'est physiquement correct car F1 est au fond du registre avec peu d'échantillons, mais ne représente pas le profil d'une note spécifique.

**Piste future** : analyse par registre (3 bandes : grave/médium/aigu) pour les instruments à grande tessiture, reflétant mieux la réalité des doublures orchestrales (qui se font dans un registre spécifique).

### Section VII — Enveloppes spectrales par instrument

**47 images individuelles** (au lieu de 3 planches groupées) :
- 13 Bois : Piccolo → Contrebasson + Saxophone alto
- 16 Cuivres : Cor → Tuba CB + toutes sourdines (cor, 4 trompette, 4 trombone, tuba)
- 18 Cordes : Violon → Contrebasse (solo + sourdine + piombo) + ensembles (+ sourdine)

Chaque image : enveloppe moyenne + ±1σ + marqueurs F1–F4 + Fp centroïde + n + σ(F2).
Taille contrôlable via CSS : `img[src*="specenv_"] { max-width: 60%; }`.

Script : `build_envelopes_by_family_html_docx.py` (réécrit : images individuelles).
Ancres ID pour sidebar : `#vii-enveloppes`, `#env-bois`, `#env-cuivres`, `#env-cordes`.

### Commentaires cordes solo — bridge hill et table d'harmonie

Tous les commentaires des 4 instruments solistes corrigés avec valeurs CSV v3 vérifiées :

**Violon** : F1=506 Hz (/o/, résonance caisse, σ=376), bridge hill F3–F5 (2347–3908 Hz), Fp=1253 Hz, convergence F1 violon ≈ F1 basson (Δ=11 Hz).

**Alto** : F1=377 Hz (/å/, résonance caisse, σ=202), bridge hill plus bas (~F3=1540 Hz), Fp=1300 Hz proche violon (Δ=47 Hz).

**Violoncelle** : F1=205 Hz (/u/, résonance table d'harmonie, σ=287), fusion violoncelle–basson via F2 vcl (506) ≈ F1 basson (495, Δ=11), Fp=1242 Hz converge avec trombone Fp=1218 (Δ=24).

**Contrebasse** : F1=172 Hz (/u/, résonance caisse, σ=36 — le plus stable), Fp=1235 Hz ≈ violoncelle (Δ=7), convergence F1 tuba CB (Δ=54).

Tableau de doublures violon mis à jour : Fp=1253 (au lieu de 893).
Section intro cordes + intro ensembles corrigées (bridge hill, valeurs Fp).

### GitHub Action — build manuel

`.github/workflows/build.yml` : `workflow_dispatch` (déclenchement manuel).
Installe numpy, matplotlib, python-docx, docxcompose. Lance `build_document_complet.py`.
Commit automatique `[auto-build]` si des changements sont détectés.

Prérequis : **Settings → Actions → General → Workflow permissions = Read and write**.

## État du repo

```
Scripts/
  extract_formants_all_techniques_v2_fixed.py   ← baseline validée (v22)
  extract_formants_all_techniques_v3_stats.py   ← v3 avec mean/std/q25/q75/db
  v5-html-docx/                                 ← scripts actifs
    common.py                                   ← CSV v3 + BW, make_graph gaussien anti-collision, compute_fp
    build_bois_html_docx.py                     ← Fp dynamique
    build_cuivres_html_docx.py                  ← Fp dynamique + note sourdines
    build_sax_html_docx.py
    build_cordes_html_docx.py                   ← Fp dynamique, bridge hill, table d'harmonie
    build_synthese_html_docx.py
    build_intro_html_docx.py
    build_envelopes_by_family_html_docx.py      ← 47 images individuelles + sourdines
    build_document_complet.py                   ← 140 TDM entries, CSS img global
    make_synthese_figures.py                    ← fig1/fig3 avec axes Bark

Resultats/
  formants_all_techniques_v3.csv                ← 487 lignes, dB depuis enveloppe moyenne ★
  formants_yan_adds_v3.csv                      ← 46 lignes, idem ★
  bandwidths_3db.csv                            ← 533 profils, BW -3dB ★

.github/workflows/build.yml                    ← GitHub Action (manual trigger)
```

## Consignes pour la suite

- **Toujours montrer une image prototype avant de régénérer les 49 images**
- **Toujours demander l'aval avant push**
- **Ne pas lancer build_document_complet** — c'est fait via GitHub Action manuellement
- **Vérifier systématiquement les 49 images** (collision + Fp) avant de proposer le push
- Fichiers specenv dans `Data/FullSOL2020_specenv par instrument/` et `Data/Yan_Adds-Divers_specenv par instrument/`

## Workflow git

```bash
git clone https://github.com/zseramnay/Formants.git
cd Formants
git config user.name "Claude" && git config user.email "claude@anthropic.com"
git remote set-url origin https://[PAT]@github.com/zseramnay/Formants.git
git push origin main
git remote set-url origin https://github.com/zseramnay/Formants.git
```

## Ce qu'on ne change PAS

- Pas de renommage F1→P1/P2 (note terminologique suffit)
- Pas de transitions formantiques/diphthongaison (hors scope, sustained only)
- Pas de normalisation morphologique (non pertinent pour instruments)

## Sur l'horizon

- **Analyse par registre** (grave/médium/aigu ou par octave) pour les instruments à grande tessiture — résoudrait le problème de F1 < F2 en dB sur l'enveloppe moyenne
- Extension au répertoire spectral contemporain (Grisey, Murail, Saariaho, Haas, Radulescu)
