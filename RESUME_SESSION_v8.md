# Résumé de session — v8 (25 mars 2026)

## Contexte

Repo: `https://github.com/zseramnay/Formants` (user: zseramnay)
Document en ligne: `https://zseramnay.github.io/Formants/Versions-html-and-docx/REFERENCE_FORMANTIQUE_COMPLETE.html`

## Historique rapide (v5–v7)

- **v5** : Points A–F (terminologie, clause sustained, CSV v3, fenêtrage Hann, axes Bark, σ(Fp))
- **v6** : Bandwidths -3dB, graphiques gaussiens, Fp dynamique, anti-collision labels, note sourdines
- **v7** : dB depuis enveloppe moyenne, 47 enveloppes individuelles, bridge hill/table d'harmonie cordes, GitHub Action

## Nouveautés v8

### 1. Amplitudes dB recalculées depuis enveloppe spectrale moyenne

Les colonnes F1_db–F6_db des CSV v3 sont lues sur l'enveloppe spectrale **moyenne** (au lieu de médiane par échantillon). Résout le problème F5 > F1 au violon. 468/533 profils mis à jour.

**Limitation** : F1 < F2 en dB pour instruments à grande tessiture → résolu par l'analyse par registre (ci-dessous).

### 2. Nouvelle représentation : « Carte spectrale vocalique »

Enveloppe specenv + zones vocaliques, échelle log, marqueurs F1–F6 + Fp, anti-collision v7 (above+below), cepstrale normalisée en violet fin. Images individuelles par registre à 51% largeur.

### 3. Analyse par registre (remplace octaves)

**Problème résolu** : l'ancienne analyse groupait par numéro d'octave MIDI (C à B), ce qui mélangeait les registres instrumentaux réels. Exemple : la flûte affichait "Oct.3 A#3–B3" au lieu de "Grave B3–A4".

**Solution** : registres définis dans `registres.md` (Yan) et encodés dans `REGISTERS` dict (common.py). Chaque instrument a ses propres registres nommés :

- **Flûte** : Grave (B3–A4), Médium (A#4–A5), Aigu (A#5–A6), Suraigu
- **Hautbois** : Grave (A#3–G4), Médium (G#4–G5), Aigu (G#5–D6), Suraigu
- **Clarinette** : Chalumeau (D3–D4), Gorge (D#4–G#4), Clairon (A4–A#5), Suraigu (B5+)
- **Basson** : Grave (A#1–A2), Bas médium (A#2–A3), Haut médium (A#3–A4), Aigu (A#4+)
- **Cor** : Pédale (F1–A1), Grave (A#1–B2), Médium (C3–E4), Aigu (F4–F5), Suraigu
- **Trompette** : Grave (F#3–B3), Médium (C4–G4), Aigu (G#4–C6), Suraigu
- **Trombone** : Pédale (E1–A1), Grave (A#1–E3), Médium (F3–E4), Aigu (F4+)
- **Tuba basse** : Grave (≤E2), Médium (F2–E3), Aigu (F3–D4), Suraigu
- **Violon** : Grave (G3–B3), Médium (C4–B4), Aigu (C5–B5), Suraigu (C6+)
- **Alto** : Grave (C3–F#3), Médium (G3–F#4), Aigu (G4–F#5), Suraigu (G5+)
- **Violoncelle** : Grave (C2–F#2), Médium (G2–F#3), Aigu (G3–F#4), Suraigu (G4+)
- **Contrebasse** : Grave (C1–B1), Médium (C2–B2), Aigu (C3–B3), Suraigu (C4+)

Ensembles partagent les registres des solistes correspondants.

Pour chaque instrument clé (16 total) :
1. **Profil formantique moyen** (courbes de cloche gaussiennes)
2. **Tableau par registre** : F1–F7 Hz/dB côte à côte + Fp
3. **Cartes spectrales vocaliques** individuelles par registre (97 images total)

### 4. Figure Bark F1×F2

`make_fig2_bark()` — échelle Bark (Traunmüller), petits marqueurs, adjustText anti-collision avec flèches. "Figure 2b" dans HTML et DOCX.

### 5. Commentaires cordes solo — bridge hill et table d'harmonie

Violon (caisse 506 Hz, bridge hill F3–F5), Alto (caisse 377 Hz, bridge hill ~F3), Violoncelle (table d'harmonie 205 Hz, fusion basson Δ=11 Hz), Contrebasse (caisse 172 Hz, σ=36 le plus stable). Doublures violon Fp corrigé (1253 Hz).

### 6. 47 enveloppes spectrales individuelles

13 Bois + 16 Cuivres (toutes sourdines) + 18 Cordes (solo+sourdine+piombo+ensembles).

### 7. Corrections

- **IRCAM → Orchidea** (specenv vient d'Orchidea, https://github.com/CarmineCella/orchidea)
- **KeyError 'color'** dans cuivres/cordes build scripts (all_info manquait la clé)
- **f-string nested quotes** (Python 3.11 : guillemets dans .replace() incompatibles)
- **GitHub Action** : pointe v6, installe adjustText

### 8. Étude pilote clarinette

Validation de l'approche par registre. LPC classique ne fonctionne pas sur spectres moyennés. Enveloppe cepstrale (ord.20) converge bien avec specenv. Fp mesure la plus robuste.

## Fonctions ajoutées dans common.py v6 (+524 lignes vs v5)

- `REGISTERS` — dict des registres pour 16 instruments (source : registres.md)
- `load_specenv_by_register()`, `load_spectrum_by_register()` (avec fallback Yan_Adds)
- `cepstral_envelope()` (troncature cepstrale ord.20)
- `make_carte_spectrale()` (v7 anti-collision above+below, échelle log, zones vocaliques)
- `compute_register_profiles()`, `make_register_table_html()`, `generate_per_register_html()`
- Alias `generate_per_octave_html = generate_per_register_html` (rétrocompatibilité)

## État du repo

```
Scripts/
  v5-html-docx/                                 ← version stable précédente
  v6-html-docx/                                 ← version active ★
    common.py                                   ← 1538 lignes
    build_bois_html_docx.py                     ← per-register : Flûte, Hautbois, Cl.Sib, Basson
    build_cuivres_html_docx.py                  ← per-register : Cor, Trompette, Trombone, Tuba
    build_cordes_html_docx.py                   ← per-register : 4 solo + 4 ensembles
    build_synthese_html_docx.py                 ← Figure 2b Bark
    build_document_complet.py                   ← CSS carte_ 51%, pointe v6
    [+ 5 autres scripts inchangés]

Resultats/
  formants_all_techniques_v3.csv                ← 487 lignes, dB enveloppe moyenne ★
  formants_yan_adds_v3.csv                      ← 46 lignes ★
  bandwidths_3db.csv                            ← 533 profils BW -3dB

.github/workflows/build.yml                    ← manual trigger, v6, adjustText
registres.md                                   ← source des registres instrumentaux
```

## Consignes pour la suite

- **Toujours montrer un prototype avant de régénérer les images**
- **Toujours demander l'aval avant push**
- **Ne pas lancer build_document_complet localement** — fait via GitHub Action
- **Ne rien faire sans demander avant**
- specenv vient d'**Orchidea** (pas IRCAM)

## Workflow git

```bash
git clone https://github.com/zseramnay/Formants.git
cd Formants && git config user.name "Claude" && git config user.email "claude@anthropic.com"
git remote set-url origin https://[PAT]@github.com/zseramnay/Formants.git
git push origin main
git remote set-url origin https://github.com/zseramnay/Formants.git
```

## Sur l'horizon

- LPC nécessiterait accès aux .wav originaux SOL2020
- Extension répertoire spectral contemporain (Grisey, Murail, Saariaho, Haas, Radulescu)
