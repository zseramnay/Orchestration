---
id: plan-de-reprise
titre: Plan de reprise — étude formantique v2
date: 2026-10-24
etat: proposition, à arbitrer
---

# Plan de reprise — v2

> **À lire avec** `_VERIFICATION-GRAVES.md` (ce qui survit de l'ancien), `_REPLICATION.md`
> (ce qui est établi), `_CARTE-VOCALIQUE.md` (la couche voyelles), `00-METHODE.md` (la méthode).

## Ce qui est déjà fait

| Acquis | État |
|---|---|
| 22 fiches instrument, front matter YAML exploitable | fait |
| 17 formants établis sur 9 instruments, répliqués sur 2 corpus | fait |
| Typologie en trois classes (A / B / C) | fait |
| Carte vocalique registre × dynamique, 22 instruments | fait |
| Vérification ZSamples sur 16 instruments | fait |
| Bibliothèque reproductible `formants_lib.py` + 3 CLI | fait |
| Vérification de l'ancien pipeline, 4 sources croisées | fait (24/10) |

**L'instrument par instrument n'est pas à refaire.** Il est fait. Ce qui reste est de la
correction, de la synthèse et de la mise en forme.

## Ce qui est conservé de l'ancien travail

| Couche | Sort |
|---|---|
| Chaîne de production — `Scripts/v6-html-docx/`, ~7 000 lignes : CSS, export DOCX avec signets et table des matières, builders par famille, `make_synthese_figures.py` | **conservée**, à restructurer en v2 |
| Dépouillement bibliographique — Backus, Giesler, Meyer, source par source | **conservé tel quel** |
| Descriptions de timbre, vocabulaire d'orchestration | **conservées** |
| Tessitures, effectifs, provenances des échantillons | **conservés** |
| Conclusions « flûte sans formant fixe », « clarinette Sib suiveuse » | **conservées et confirmées** (ρ mesuré, balance impair−pair mesurée) |
| Dépendance dynamique de la trompette | **phénomène conservé**, valeurs refaites |
| `Analyses_formants-mai26` — 45 CSV par note | **promu en source active** (voir `_VERIFICATION-GRAVES.md`) |
| Toutes les valeurs F1–F4 et Fp | remplacées |
| Matrices de convergence, tableaux de doublures, partie VI | refaits sur une autre base |
| 47 figures `specenv_*` en tant qu'enveloppes | écartées |
| 85 cartes `carte_*` | régénérées ; le **code** qui les produit est conservé |

## Structure v2 du site

Reprise de votre demande : une page par instrument, regroupées par famille dans un sommaire,
et une page par sujet de synthèse.

```
index.html                     sommaire général, les deux études
  formants/
    index.html                 sommaire par famille, avec statut de chaque instrument
    methode.html               méthode, limites, quand le formant n'a pas de sens
    bois/01-petite-flute.html … 12-clarinette-contrebasse.html
    cuivres/13-cor.html … 18-tuba-contrebasse.html
    cordes/19-violon.html … 22-contrebasse.html
    syntheses/
      formants-etablis.html    les 17, avec les trois tests + corroboration mai26
      typologie.html           les trois classes
      carte-vocalique.html     registre × dynamique, les quatre comportements
      dynamique.html           l'effet de la force de jeu (cuivres)
      tuyau-cylindrique.html   balance impair−pair, la famille clarinette
      cordes-et-bridge-hill.html   classe C, méthode 8 et méthode 12 ensemble
      heritage.html            ce que disait l'ancienne étude, et pourquoi
      non-mesurable.html       sourdines, ensembles, et les limites assumées
```

Une page par instrument se génère directement du front matter de sa fiche `.md`. Le sommaire
par famille affiche le statut : nombre de formants établis, classe, cellules vocaliques
assignables. Le lecteur voit d'un coup d'œil où l'étude dit quelque chose et où elle se taît.

**Point d'entrée technique** : toute la chaîne lit ses données par `load_all_csvs()` et
`get_f(inst, tech)` dans `common.py`, depuis `formants_all_techniques.csv`. Un seul point de
bascule. Produire un nouveau CSV depuis les fiches, et la chaîne se régénère.

## Les étapes, dans l'ordre

### Phase 1 — corrections imposées par la vérification du 24/10

1. **Recompter les 17 formants établis au critère strict** (même point de grille, non ≤ 1
   demi-ton). Dire combien survivent. *C'est la correction la plus urgente : elle touche le
   résultat central.*
2. Réécrire tous les énoncés de fréquence en **note + fréquence approchée** — « Ré7
   (≈ 2 350 Hz) » — puisque notre résolution est d'un demi-ton.
3. Corriger `00-METHODE.md` : la limite des 276 Hz n'est pas le discriminant ; c'est la
   typologie.
4. Ajouter la **corroboration mai26** comme quatrième test, distinct de la réplication par
   corpus, dans `00-METHODE.md` et dans les fiches concernées.
5. Corriger l'instabilité du **violon au suraigu** (méthode 12) — seul défaut de méthode
   reproductible identifié.

### Phase 2 — mesures qui manquent

6. **Appariement note par note** entre nos formants et le LPC de mai26, pour transformer
   l'indice de corroboration en test.
7. **Relancer les cuivres** avec les registres pédale ajoutés le 24/10.
8. **Mesure par registre** pour les zones médianes du hautbois et du cor anglais, les deux
   instruments les plus instables.
9. Signaler les coïncidences **F2 / formant établi** dans la sortie de `carte_vocalique.py`,
   comme indice de confiance sur les cordes.

### Phase 3 — production

10. Générer le nouveau `formants_all_techniques.csv` depuis les 22 fiches.
11. Restructurer la chaîne en **v2** selon l'arborescence ci-dessus.
12. Régénérer les figures : composites par instrument, cartes vocaliques par registre, figures
    de famille. Écarter les `specenv_*`.
13. Refaire la figure des voyelles sur **Calliope** et non Meyer.
14. Écrire les pages de synthèse.

### Suspendu, par décision

- **Sourdines et ensembles** — 28 configurations. Réplication médiane à 4,0 demi-tons sur 18
  paires : non publiables comme résultats établis. Deux issues possibles, à votre choix : un
  troisième corpus, ou publication explicite comme exploratoire.
- **Trombone basse** — registre propre à ajouter dans `registres.md` quand vous le souhaitez.
- **Méthode 12 sur mélanges réels** — fusion contre brillance/renforcement. C'est la question
  d'orchestration qui a lancé cette reprise, et elle n'est pas encore traitée. À faire une fois
  la phase 1 close, car elle a besoin d'une base stable.

## Le choix qui vous revient

La v2 **dira moins** que la v1. L'ancienne annonçait F1–F4 sur 51 configurations, des matrices
de convergence complètes et un tableau de doublures vérifiées. La v2 donnera :

- un ensemble réduit de formants établis, sur une partie des instruments seulement ;
- une typologie qui dit, pour chaque instrument, si la notion de formant s'y applique ;
- une carte vocalique par registre et dynamique, qui couvre **tous** les instruments ;
- l'effet mesuré de la dynamique et de la géométrie du tuyau ;
- et une liste explicite de ce qui n'est pas mesurable avec ces données.

Pour l'orchestration pratique, la perte est réelle : les matrices de doublures n'existeront
plus sous cette forme. Pour l'enseignement, le gain est net — la v2 explique **pourquoi** la
question est difficile, ce que la v1 ne pouvait pas faire.

La question à trancher avant la phase 3 : la partie synthèse doit-elle rester organisée par
**convergences entre instruments**, comme la v1, ou par **propriétés mesurées** (typologie,
dynamique, géométrie, voyelles), comme proposé ci-dessus ? Cela change la forme du document
final et la façon dont il s'enseigne.

## Journal

- **2026-10-24** — Première version. Vérification des graves intégrée.
