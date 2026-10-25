# Scripts — mesure formantique v7 (octobre 2026)

Chaîne de mesure autonome. Elle remplace entièrement les pipelines v1 à v6 et l'extracteur
`extract_formants_all_techniques_*.py`, qui reposaient sur les `specenv` d'Orchidea et sur le
Fp par centroïde — l'un et l'autre invalides pour cet usage (voir
`../../Etude-Formants/Instruments/00-METHODE.md`).

## Fichiers

| Fichier | Rôle |
|---|---|
| `formants_lib.py` | **bibliothèque autonome.** Toute la méthode : spectre moyen attaque exclue, harmoniques, résidus locaux, enveloppe composite, tests de rejet, enveloppe vraie adaptative, descripteurs. Aucune dépendance au reste du projet. |
| `mesurer.py` | **point d'entrée.** Une configuration instrumentale → spectres + JSON + fiche `.md` pré-remplie. |
| `spectres_sans_attaque.py` | version autonome de l'étape 1 seule (historique ; `formants_lib` fait la même chose). |
| `enveloppe_composite.py`, `tests_de_rejet.py`, `critere_assoupli.py`, `analyse_variantes.py` | scripts de la session du 23/10, conservés pour traçabilité. **Ne pas réutiliser** : ils dépendent d'un fichier temporaire non versionné. Tout passe désormais par `mesurer.py`. |

## Prérequis

```
python3 -m pip install numpy soundfile
```

## Usage

**Un instrument :**

```bash
python3 mesurer.py \
  --racine ~/kDrive/Samples/Orchidea-dbs/OrchideaSOL_and_friends/FullSOL2020 \
  --dossier Strings/Violin \
  --technique ordinario \
  --nom Violon \
  --sortie resultats/ --fiche
```

**Une sourdine** dont la technique porte un autre nom :

```bash
python3 mesurer.py --racine .../FullSOL2020 \
  --dossier "Brass/Trumpet_C+sordina_wah" --technique ordinario_open \
  --nom "Trompette+wah ouvert" --sortie resultats/ --fiche
```

**Toute une famille d'un coup :**

```bash
python3 mesurer.py --racine .../Yan_Adds-Divers --balayer Winds \
  --technique ordinario --sortie resultats/ --fiche
```

Les instruments dépourvus de la technique demandée sont signalés et ignorés. Plusieurs
techniques peuvent être fusionnées : `--technique ordinario,non_vibrato`.

**Réanalyser sans relire les WAV** (utile après modification d'un seuil) :

```bash
python3 mesurer.py --depuis-spectres resultats/Violon.spectres.txt --fiche --sortie resultats/
```

## Sorties

- `<nom>.spectres.txt` — spectres moyens au format Orchidea `spectrum` (chemin;1024 valeurs).
  C'est le fichier lourd ; il permet de tout réanalyser sans retoucher aux WAV.
- `<nom>.json` — mesures complètes, dont l'enveloppe composite bande par bande.
- `<nom>.md` — fiche pré-remplie : en-tête YAML complet, tableaux chiffrés, et sections
  rédactionnelles marquées **À RÉDIGER**.

## Ce que le script ne décide PAS à votre place

Trois champs de l'en-tête YAML sont proposés automatiquement et doivent être **validés à la
main** :

- `classe` — A, B ou C. La règle automatique est grossière : ρ < 0,55 → A ; sinon des
  formants retenus → C ; sinon B. Elle suffit en première approche mais ne remplace pas
  l'examen du cas.
- `pertinence_formant` — pertinente, partielle ou non pertinente.
- Toute la partie rédigée.

## Pièges à ne pas refaire

1. **Ne pas utiliser les `specenv` d'Orchidea** pour une mesure de formant. Leur lissage
   cepstral est d'ordre fixe 160, codé en dur, donc valide seulement sous
   44 100 / 160 = 276 Hz. Au-dessus, le « formant » est un harmonique.
2. **Ne pas réintroduire le Fp par centroïde de bande.** Un centroïde calculé dans [a, b]
   tombe nécessairement entre a et b : il mesure la pente dans une fenêtre imposée, pas une
   résonance. Sa stabilité apparente est un effet de bornage.
3. **Ne pas redresser linéairement en log-fréquence.** Sur un spectre concave, cela fabrique
   une bosse médiane et donc de faux formants — ce qui avait produit un faux accord avec la
   littérature pour le hautbois. Le redressement local à ±3/4 d'octave est là pour ça.
4. **Ne pas abaisser `RESIDU_MIN` ni retirer le test de robustesse de fenêtre.** Si un corpus
   est trop petit, le critère assoupli (découpage au rang médian + bootstrap par note) est
   déjà conçu pour cela, et répond « non mesurable » plutôt que de produire un chiffre.
5. **Ne pas publier ρ seul comme indice d'absence de formant.** Un ρ élevé a deux causes
   opposées : aucune résonance (flûtes) ou beaucoup de résonances dont aucune ne domine
   (cordes). C'est le couple (ρ, nombre de formants retenus) qui classe.
6. **Ne pas comparer des écarts sous la résolution de mesure.** Un bin de FFT vaut 10,77 Hz.
   Un « Δ = 11 Hz » est un bin d'écart, pour des dispersions de 100 à 400 Hz.
7. **Refaire le contrôle de nomenclature sur toute nouvelle base.** `mesurer.py` l'affiche et
   alerte s'il s'écarte de 1,00 de plus de 3 %. S'il vaut 0,89, les noms de fichiers sont en
   hauteurs écrites et non réelles, et tous les f0 sont faux.
8. **Un test de permutation global sur la bande ne vaut rien ici** : la bande a été détectée
   sur la positivité de son résidu puis serait testée sur cette même positivité. Essayé,
   abandonné.

## Validation

Au 24/10/2026 : 54 configurations mesurées, 0 erreur de lecture, contrôle de nomenclature à
1,00 partout. Les mesures attaque incluse / exclue concordent à un ou deux douzièmes
d'octave, et le contrôle `ordinario` contre `non_vibrato` sur les quatre cordes solistes ne
déplace aucun formant de plus de deux douzièmes d'octave.
