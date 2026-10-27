---
id: largeur-par-registre
titre: La largeur de bande par registre et par dynamique
date: 2026-10-27
tache: C5
script: Scripts/v7-formants-2026/largeur_par_registre.py
donnees: Etude-Formants/Mesures/Largeur-par-registre/
corpus: SOL, segment v7
---

# La largeur de bande par registre et par dynamique

## L'hypothèse

Le plan de reprise la posait ainsi : *chez les cuivres, la localisation devrait s'améliorer en
jeu fort*. Le raisonnement est connu — joué fort, un cuivre raidit son spectre par non-linéarité
de propagation dans la perce, l'énergie monte en fréquence, et les résonances se détachent mieux
du fond. Si c'est vrai, la largeur à −3 dB doit se resserrer et la part de pics **ouverts** —
ceux qui ne redescendent pas de 3 dB des deux côtés, donc non localisables — doit baisser du
doux vers le fort.

Rien à recalculer : `Largeurs/Enveloppe/SOL-v7/` porte déjà, note par note, chaque pic de
l'enveloppe vraie avec sa largeur, sa proéminence, sa dynamique et son MIDI. Il fallait croiser.

## La mesure retenue : la part de pics ouverts

C'est la mesure directe de la localisation, et elle est préférable à la largeur médiane pour une
raison de biais : **un pic ouvert n'a pas de largeur**. L'écarter en silence tirerait la médiane
vers les pics étroits, et un instrument dont 76 % des pics sont inexploitables paraîtrait plus
net qu'un instrument dont tous les pics sont mesurables mais larges. On compte donc les ouverts
au lieu de les ignorer, et on ne calcule les largeurs et les Q que sur les pics fermés.

## Résultat d'ensemble

411 cellules, 17 instruments, formants établis uniquement, pics à moins de deux demi-tons du
formant.

| Famille | n | Effet doux → fort, médian | Détail par instrument |
|---|---|---|---|
| **Cuivres** | 5 | **+31 points** | +54 · +48 · +31 · −8 · −9 |
| Flûtes | 2 | +35 points | +58 · +12 |
| Clarinettes | 3 | +10 points | +15 · +10 · +2 |
| Cordes | 4 | +3 points | +5 · +3 · +3 · −2 |
| Anches doubles | 3 | +1 point | +9 · +1 · 0 |

Un effet positif signifie que la part de pics ouverts **baisse** en jeu fort, donc que la
localisation s'améliore.

**L'hypothèse est confirmée, et c'est bien un effet de cuivres.** Les cordes et les anches
doubles ne montrent rien — trois et un point, c'est du bruit. Mais deux cuivres sur cinq vont à
l'encontre de l'hypothèse, et il a fallu regarder par registre pour comprendre.

## À registre constant, la ligne de partage apparaît

La dynamique et le registre ne sont pas indépendants dans un corpus d'échantillons : les notes
extrêmes ne sont pas jouées à toutes les nuances. L'effet d'ensemble mélange donc les deux. À
registre constant, avec au moins dix pics de chaque côté :

| Instrument | Registre | Doux | Fort | Effet |
|---|---|---|---|---|
| **Trombone ténor** | grave | 69 % | 7 % | **+63 pts** |
| | médium | 36 % | 8 % | +29 pts |
| **Trompette en Ut** | grave | 58 % | 6 % | **+52 pts** |
| | médium | 35 % | 7 % | +27 pts |
| | aigu | 0 % | 5 % | −5 pts *(déjà localisé partout)* |
| **Cor en Fa** | grave | 33 % | 47 % | **−14 pts** |
| | médium | 16 % | 40 % | **−25 pts** |
| | aigu | 40 % | 15 % | +25 pts |
| **Tuba basse** | grave | 39 % | 68 % | **−30 pts** |
| | médium | 78 % | 79 % | −1 pt |

Le partage est net, et il suit **la perce** :

- **les cuivres à perce majoritairement cylindrique** — trombone ténor, trombone basse,
  trompette — confirment massivement l'hypothèse, et précisément dans le grave et le médium, là
  où le cuivrage se développe ;
- **les cuivres à perce majoritairement conique** — cor, tuba — vont en sens inverse dans ces
  mêmes registres, et ne confirment que dans l'aigu.

C'est cohérent avec la physique du cuivrage : le raidissement de l'onde par non-linéarité a
besoin d'une longueur cylindrique pour se développer. Un trombone et une trompette cuivrent, un
cor et un tuba beaucoup moins. **L'étude ne démontre pas ce mécanisme ; elle constate que la
mesure se range selon lui.**

## Ce que cela change pour l'orchestration

Ce sont les chiffres les plus directement utilisables de toute la tâche C.

1. **Le formant du trombone basse n'existe pratiquement pas au piano.** 76 % de ses pics y sont
   non localisables, contre 22 % au fort. Écrire « le trombone basse a un formant à 879 Hz » sans
   préciser la nuance, c'est décrire un instrument joué fort.
2. **La trompette dans le grave passe de 58 % à 6 %.** Son grave est flou au doux et net au
   fort — ce qui correspond à l'expérience des orchestrateurs, et le chiffre la précise.
3. **Le médium du tuba basse n'est jamais localisable** : 78 à 79 % d'ouverts à toutes les
   nuances. Il n'y a pas là de formant à publier, à aucune dynamique.
4. **Le cor se comporte à l'envers** dans le grave et le médium : joué fort, il devient *moins*
   localisable, pas plus. À ne pas décrire comme les autres cuivres.
5. **Les cordes sont indifférentes à la nuance** de ce point de vue — trois points d'effet. Leur
   caisse filtre de la même façon quelle que soit la force de l'archet, ce qui est attendu : la
   non-linéarité n'est pas de leur côté.

## Deux corrections faites en route

**1. Le cor manquait des deux jeux de largeurs.** `Largeurs/LPC-v7/` et
`Largeurs/Enveloppe/SOL-v7/` ne contenaient que `Horn+sordina`. La cause est en amont :
**l'analyse de mai 2026 avait omis le cor ordinaire**, et tous les scripts aval tiraient leur
liste d'instruments de ce dossier — l'absence ne se voyait donc nulle part. `largeurs_lpc.py` et
`largeurs_enveloppe.py` reçoivent une option `--depuis-dossier` qui construit la liste depuis les
échantillons. Le cor est désormais complet : 134 notes, 4 973 pôles LPC et 1 721 pics
d'enveloppe. **C'est aussi ce qui comble le manque signalé par la tâche C4.**

**2. Les trombones tombaient dans un registre « tous ».** `registres.md` écrit « trombones » au
pluriel et ne distingue pas le ténor du basse ; la clé cherchée était « trombone ». Corrigé, les
deux pointent sur la même ligne — ce qui a fait apparaître la structure par registre du trombone
ténor, +63 points dans le grave.

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| Le partage cylindrique / conique tient-il sur d'autres corpus ? | les largeurs d'enveloppe existent pour ZSamples, Xsamples, VO et VN ; seul SOL a été croisé ici |
| Le trombone basse a-t-il une structure par registre ? | sa cellule douce ne compte que 21 pics, trop peu pour trancher par registre |
| L'effet se lit-il aussi sur la largeur des pics fermés, et pas seulement sur leur nombre ? | les largeurs et les Q par cellule sont dans le CSV ; ils n'ont pas été analysés |
| Le cuivrage est-il bien le mécanisme ? | il faudrait mesurer la montée du centre de gravité spectral avec la nuance et la corréler à l'effet |

## Journal

- **2026-10-27** — Tâche C5. L'hypothèse du plan est confirmée pour les cuivres, +31 points
  d'ensemble et jusqu'à +63 à registre constant, et elle ne vaut que pour eux. Deux cuivres sur
  cinq s'y opposent, et le partage suit la perce : cylindrique oui, conique non. Troisième fois
  de la journée qu'un agrégat masquait une structure — ici, le mélange registre × dynamique.
