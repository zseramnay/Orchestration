# Étude formantique — où se trouve quoi

> Rangement du 25/10/2026. Ce fichier existe pour qu'on n'ait pas à deviner.

## Le principe

Deux choses étaient mélangées et sont maintenant séparées :

| | |
|---|---|
| **`Instruments/`** | **la rédaction** — ce qui s'écrit, se lit, et finira dans le document |
| **`Mesures/`** | **les données** — ce que les scripts produisent, un dossier par corpus |

Rien d'autre ne doit entrer dans `Instruments/`.

## Arborescence

```
Etude-Formants/
│
├── Instruments/                 LA RÉDACTION
│   ├── 00-METHODE.md            méthode, critères de rejet, typologie
│   ├── _AMORCE-SESSION.md       à coller au début d'une conversation vierge
│   ├── _CORPUS-CROISES.md       les 5 corpus, la règle de provenance, les 20 formants
│   ├── _REPLICATION.md          le critère à deux routes, le test de permutation
│   ├── _VERIFICATION-GRAVES.md  pourquoi l'ancien pipeline est faux
│   ├── _CARTE-VOCALIQUE.md      la couche voyelles, 22 instruments
│   ├── _LARGEURS-DE-BANDE.md    largeurs LPC et enveloppe confrontées
│   ├── _PLAN-DE-REPRISE.md      où en est le projet, dans quel ordre avancer
│   ├── FullSOL2020/             12 fiches — instruments venus de l'IRCAM
│   └── Yan_Adds/                10 fiches — instruments venus de VSL/ConTimbre/perso
│
├── Mesures/                     LES DONNÉES, un dossier par corpus
│   ├── SOL/                     json et csv des mesures d'octobre
│   ├── ZSamples/                64 json · 40 fiches · 20 spectres
│   ├── Xsamples/                json · fiches · spectres
│   ├── VO/                      idem
│   └── VN/                      idem
│
├── Largeurs/                    LARGEURS DE BANDE
│   ├── LPC/                     45 configurations, segmentation mai26
│   ├── LPC-v7/                  idem, segmentation du projet (onset + 200 ms)
│   ├── Enveloppe/{SOL,SOL-v7,ZSamples,Xsamples,VO,VN}/
│   └── Comparaison/             confrontation et contrôle de résolution
│
├── Analyses_formants-mai26/     45 CSV par note : LPC, cepstre 40/80/160
├── Formants-Francais/           voyelles Calliope (Tubach 1989)
├── Resultats/ · Resultats-anciens/   l'ancien pipeline — archive
├── Versions-html-and-docx/      les versions publiées, v1 à v6
├── Conversations/               transcriptions — archéologie, pas amorce
└── Archives/                    périmé, à supprimer quand vous voulez
```

## La règle de provenance — pourquoi `Instruments/` est scindé en deux

Ce n'est pas cosmétique. **`Yan_Adds-Divers` contient du VSL, et le corpus VN est du VSL.**
Vérifié sur trois paires de fichiers : nombre d'échantillons identique au sample près,
corrélation jusqu'à **1,00000**. Ce sont les mêmes enregistrements.

> Une concordance entre `Yan_Adds` et `VN` sur un même instrument **ne compte pas** comme une
> réplication. Elle compte pour les douze instruments de `FullSOL2020`, qui viennent de l'IRCAM.

D'où la séparation des fiches : elle rend la provenance visible au moment où l'on écrit.
Détail et conséquences dans `Instruments/_CORPUS-CROISES.md`.

## Ce qui est dans `Archives/`

| Dossier | Pourquoi |
|---|---|
| `ZSamples-contamine-2026-10-25/` | régénération ZSamples faite avec un `corpus.py` dont le **filtre de technique avait disparu** : 597 fichiers de cor au lieu de 71, sourdines et multiphoniques inclus |
| `Largeurs-Enveloppe-ZSamples-contamine/` | les largeurs calculées sur ces spectres |
| `ZSamples-refait-json/` | la régénération corrigée ; ses spectres ont été versés dans `Mesures/ZSamples/`, ses json font doublon avec les originaux |
| `zsamples-2026-10-24/` | 11 des 12 fichiers déjà présents dans `Mesures/ZSamples/` |
| `_a_supprimer/` | trois json écrasés par un bug de nommage le 23/10 |

**Aucun chiffre publié ne vient de ces fichiers.** `_CORPUS-CROISES.md` s'appuyait sur la mesure
originale, et la régénération corrigée la reproduit à l'identique sur les vingt configurations —
cor 71 fichiers [1 759], violon 164 [2 216 · 2 792 · 3 948 · 5 583], trompette 74 [1 479].

Vous pouvez supprimer `Archives/` entièrement.

## Les scripts

Tous dans `../Scripts/v7-formants-2026/` :

| Script | Rôle |
|---|---|
| `formants_lib.py` | la bibliothèque — spectres, enveloppe composite, tests de rejet |
| `corpus.py` | verser un nouveau corpus, quelle que soit sa nomenclature |
| `carte_vocalique.py` | méthode 12, espace vocalique en Bark |
| `largeurs_lpc.py` | largeurs par pôles LPC |
| `largeurs_enveloppe.py` | largeurs à −3 dB sur l'enveloppe vraie |
| `mesurer.py` | mesure d'une configuration unique |

Ne pas utiliser les versions v1 à v6, ni `extract_formants_all_techniques_*.py` : elles
reposent sur les `specenv` d'Orchidea et sur le Fp par centroïde, tous deux invalides.

**Un piège à connaître** : dans un descripteur `corpus.py`, le champ `techniques` est
**obligatoire** dès qu'un dossier mélange les techniques. Sans lui, les sourdines et les modes
de jeu entrent dans la mesure. C'est ce qui a produit la régénération contaminée ci-dessus.

## Journal

- **2026-10-25** — Séparation rédaction / mesures. Fiches scindées par provenance.
  `Instruments-ZSamples/` versé dans `Mesures/ZSamples/`. Largeurs regroupées sous `Largeurs/`.
  Correction du filtre de technique de `corpus.py` et mise à l'écart des données contaminées.
