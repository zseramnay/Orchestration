---
id: verification-graves
titre: Vérification des instruments graves — l'ancien pipeline est-il valide en dessous de 276 Hz ?
date: 2026-10-24
sources: Analyses_formants-mai26 (45 CSV par note) · formants_*_v3.csv · fiches 01–22
---

# Vérification des instruments graves

## L'hypothèse testée

Le lissage cepstral d'Orchidea utilise `NUM_SMOOTH = 160`, qui ne retire le peigne
harmonique que si f0 < 44 100 / 160 = **276 Hz**. D'où l'hypothèse : l'ancien F1 serait une
mesure d'enveloppe **valide pour les instruments graves** et un suiveur de fondamentale pour
les aigus. Une dizaine d'instruments auraient alors été mesurés correctement.

**Cette hypothèse est réfutée comme règle**, et le test a livré autre chose de plus utile.

## L'outil : mai26 contient un proxy exact de l'ancien pipeline

Les 45 CSV de `Analyses_formants-mai26` portent une colonne **`cep_F1_order160`** — le même
ordre que le `NUM_SMOOTH` d'Orchidea, mais mesuré **note par note**, avec la note, la
dynamique et le MIDI. L'ancien pipeline ne publiait qu'une médiane par instrument.

Et le proxy est fidèle là où il peut l'être. Médianes :

| Instrument | Ancien F1 (specenv) | `cep_F1_order160` (mai26) |
|---|---|---|
| Contrebasson | 226,1 Hz | **226 Hz** |
| Tuba contrebasse | 226,1 Hz | **226 Hz** |
| Contrebasse | 172,3 Hz | **172 Hz** |

Identiques au Hz. **mai26 nous donne donc l'ancienne mesure note par note**, ce que nous
n'avions jamais eu. C'est le premier résultat : ce répertoire est l'archive la plus utile du
projet, et non un essai abandonné.

## Découverte 1 — les 2 099 anciennes valeurs sont des centres de bin FFT

Test sur toutes les valeurs F1 à F4 de `formants_all_techniques_v3.csv` et
`formants_yan_adds_v3.csv` :

> **2 099 valeurs sur 2 099 tombent exactement sur un centre de bin FFT**
> (largeur de bin = 44 100 / 4 096 = 10,7666 Hz ; écart médian 0,0025 bin).

| Valeur publiée | Bin |
|---|---|
| 226,1 Hz | 21,00 |
| 323,0 Hz | 30,00 |
| 172,3 Hz | 16,00 |
| 495,3 Hz | 46,00 |
| 742,9 Hz | 69,00 |
| 1 109,0 Hz | 103,00 |

Conséquence immédiate et définitive : **toutes les convergences annoncées à « Δ = 11 Hz »
étaient des bins adjacents**, c'est-à-dire un écart d'un pas de grille. Il n'y avait aucune
mesure derrière cette précision. La « convergence quasi-parfaite violon–basson » (506 contre
495 Hz) est bin 47 contre bin 46.

Cela ne disqualifie pas les écarts plus grands. À 226 Hz un bin vaut 0,8 demi-ton, donc trois
instruments dans le même bin s'accordent réellement à moins d'un demi-ton. La question reste
de savoir si la valeur mesure quelque chose.

## Découverte 2 — l'ordre 160 trouve l'harmonique le plus fort, pas l'enveloppe

J'ai corrélé `cep_F1_order160` à la fondamentale, note par note, sur les 45 configurations
(Spearman). Si la mesure suivait la hauteur, ρ serait élevé.

| ρ faible — la mesure ne suit **pas** la hauteur | ρ élevé — la mesure **suit** la hauteur |
|---|---|
| Cor anglais −0,01 · Trombone +0,04 · Tuba contrebasse +0,08 | Petite flûte **+0,75** · Alto-ens. +0,83 |
| Trombone basse +0,12 · Trompette +0,14 · Hautbois +0,15 | Violon-ens. +0,82 · Clarinette Sib **+0,67** |
| Clarinette basse +0,15 · Tuba basse +0,19 | Flûte +0,58 · Alto +0,60 |
| Contrebasson **−0,23** | Violoncelle-ens. +0,71 |

**Le classement ne suit pas du tout la proportion de notes sous 276 Hz.** Le hautbois a 8 %
de ses notes sous la limite et ρ = 0,15 ; la contrebasse en a 86 % et ρ = 0,44.

Ce qu'il suit, c'est **notre typologie**. ρ est faible exactement pour les instruments de
classe A — résonance fixe dominante — et élevé pour la classe B, à spectre suiveur.

L'explication est mécanique. Quand le peigne harmonique fuit dans le lissage, le pic détecté
tombe sur **l'harmonique le plus fort**. Si l'instrument possède une résonance dominante, cet
harmonique est celui qui la traverse, et la mesure se trouve *accidentellement* dans la bonne
région. Sinon, il tombe sur la fondamentale ou un harmonique bas.

C'est pourquoi le cor était le seul instrument dont les anciens chiffres se confirmaient : non
parce que sa tessiture est grave, mais parce qu'il a une résonance dominante. Et cela confirme
rétrospectivement la « découverte majeure » de l'ancienne étude sur la clarinette en Sib —
« le pic suit la fondamentale ou le 3ᵉ harmonique » : ρ = 0,67, elle avait raison.

## Verdict instrument par instrument

| Instrument | Ancien F1 | cep-160 | Env. vraie (méth. 6) | Formants établis (méth. 8) | Verdict |
|---|---|---|---|---|---|
| **Basson** | 495 Hz | 538 | — | **554 Hz** | **CONFIRMÉ** — 4 méthodes à ≤ 2 demi-tons |
| Contrebasson | 226 | 226 | 301 | 329 · 698 · 1 174 | non confirmé (6,5 dt sous méth. 8) |
| Tuba contrebasse | 226 | 226 | 269 | 370 · 879 · 1 396 | non confirmé (8,5 dt) |
| Tuba basse | 226 | 275 | 280 | 466 · 659 · 987 · 1 479 | non confirmé (12,5 dt) |
| Trombone ténor | 237 | 668 | 538 | 830 · 1 567 · 2 091 · 3 134 | non confirmé — le proxy lui-même divorce |
| Trombone basse | 258 | 775 | 646 | 879 · 1 759 | non confirmé |
| Contrebasse | 172 | 172 | 151 | 466 · 1 396 · 2 635 · 3 948 | non confirmé — zone de f0 (ρ 0,90) |
| Violoncelle | 205 | 1 109 | 312 | 554 · **1 108** · 2 216 · 3 948 | non confirmé, mais voir ci-dessous |
| Clarinette basse | 323 | 743 | 323 | 1 244 · 2 487 | non confirmé (classe B, ρ 0,88) |
| Clarinette contrebasse | 323 | 1 249 | 172 | 1 759 | non confirmé |
| Flûte basse | 302 | 603 | 420 | — | non confirmé (classe B) |
| Flûte contrebasse | 334 | 759 | 382 | 1 759 | non confirmé (classe B) |

**Un seul instrument sur douze voit son ancienne valeur confirmée : le basson.** Et c'est
cohérent — c'est le seul instrument sur lequel tout concorde, y compris la littérature
(Backus 470, Giesler 500, Meyer 500) et la carte vocalique (F1 470–590 Hz).

### L'« unisson formantique parfait » à 226 Hz est écarté

Contrebasson, tuba basse et tuba contrebasse donnaient tous trois 226,1 Hz, présenté comme un
unisson entre trois familles. Les trois tombent dans le bin 21. Mais nos propres méthodes leur
donnent **329, 466 et 370 Hz** — trois valeurs distinctes. L'unisson est une propriété de la
mesure d'ordre 160 sur trois spectres à peigne très dense, non une propriété des instruments.

Même conclusion pour les 323,0 Hz identiques de la clarinette basse et de la clarinette
contrebasse (bin 30).

### Un défaut distinct, sur le violoncelle

Pour le violoncelle, `cep_F1_order160` donne **1 109 Hz** et notre formant établi est
**1 108 Hz** — accord à 0,0 demi-ton. Mais l'ancien pipeline publiait **205 Hz**.

Le lissage avait donc trouvé la bonne résonance ; c'est la **sélection du pic** qui a pris le
plus grave au lieu du plus saillant. C'est un défaut indépendant de celui du lissage, et il
affecte potentiellement tous les instruments où l'ancien pipeline a publié une valeur basse.

## Corroboration de nos formants par une famille de méthodes indépendante

mai26 contient du **LPC** (F1, F2, F3), qui ne partage aucune étape avec nos méthodes 6, 8
et 12. Confrontation des 39 formants retenus des fiches aux médianes LPC et cepstre-80 :

> **12 sur 39 corroborés à ≤ 1 demi-ton ; 22 sur 39 à ≤ 2 demi-tons.**

Les accords les plus nets :

| Formant établi | Méthode mai26 | Écart |
|---|---|---|
| Violoncelle 1 108 Hz | cep-160 : 1 109 Hz | **0,0 dt** |
| Alto 2 091 Hz | cep-160 : 2 089 Hz | **0,0 dt** |
| Contrebasse 1 396 Hz | cep-80 : 1 389 Hz | 0,1 dt |
| Tuba contrebasse 370 Hz | cep-80 : 366 Hz | 0,2 dt |
| Trombone ténor 2 091 Hz | LPC F2 : 2 120 Hz | 0,2 dt |
| Clarinette en Sib 3 517 Hz | LPC F3 : 3 625 Hz | 0,5 dt |
| Violoncelle 3 948 Hz | LPC F3 : 3 843 Hz | 0,5 dt |
| Contrebasse 3 948 Hz | LPC F3 : 3 823 Hz | 0,6 dt |
| Cor anglais 2 216 Hz | LPC F2 : 2 312 Hz | 0,7 dt |
| Clarinette basse 1 244 Hz | cep-80 : 1 195 Hz | 0,7 dt |
| **Violon 2 348 Hz** (*bridge hill*) | cep-80 : 2 455 Hz | 0,8 dt |
| Trombone basse 879 Hz | cep-80 : 926 Hz | 0,9 dt |

**Réserve à ne pas oublier** : comparer une médiane sur toutes les notes à une position de
formant est un indice, pas un test — et retenir la plus proche de cinq colonnes gonfle le
taux. Un vrai test exigerait un appariement note par note. À faire ; les données le permettent.

Noter aussi que le **cor anglais** est corroboré ici (2 216 Hz par LPC, à 0,7 dt) alors qu'il
ne se réplique pas sur ZSamples. Corroboration par méthode et réplication par corpus sont deux
exigences distinctes, et il ne passe que la première.

## Découverte 3 — notre propre méthode est quantifiée au demi-ton

En vérifiant ces écarts, un fait saute aux yeux : **tous nos formants établis sont des
fréquences de notes tempérées.**

| Notre valeur | Note | Notre valeur | Note |
|---|---|---|---|
| 554 Hz | Do♯5 = 554,4 | 1 108 Hz | Do♯6 = 1 108,7 |
| 1 396 Hz | Fa6 = 1 396,9 | 466 Hz | La♯4 = 466,2 |
| 2 348 Hz | Ré7 = 2 349,3 | 4 183 Hz | Do8 = 4 186,0 |
| 879 Hz | La5 = 880,0 | 370 Hz | Fa♯4 = 370,0 |

C'est une conséquence directe de la méthode 8 : l'enveloppe composite est construite sur un axe
fait des fréquences harmoniques des notes, donc d'une grille tempérée. **La résolution de notre
méthode est d'un demi-ton**, pas mieux.

Deux conséquences à acter :

1. Nous ne pouvons annoncer aucune précision supérieure au demi-ton. Écrire « 2 348 Hz »
   suggère une précision au Hz que nous n'avons pas ; il faut écrire **Ré7 (≈ 2 350 Hz)**.
2. Notre critère de réplication « ≤ 1 demi-ton » est **au grain de la grille** : il accepte un
   pas de grille adjacent. C'est le même type d'indulgence que le « Δ = 11 Hz » de l'ancienne
   étude, un niveau plus haut. Un critère strict dirait : **même point de grille**.

Il faut donc recompter les 17 formants établis avec le critère strict, et dire combien
survivent. C'est la première correction à apporter à `_REPLICATION.md`.

## Ce que cette vérification change

1. **L'hypothèse des 276 Hz est abandonnée.** Le discriminant est la typologie, non la
   tessiture. À corriger dans `00-METHODE.md`, qui présente la limite comme la cause unique.
2. **Les instruments graves ne sont pas réhabilités** — un seul l'est, le basson.
3. **mai26 devient une source active** et non une archive : proxy note par note de l'ancien
   pipeline, et famille de méthodes indépendante pour corroborer.
4. **Notre propre résolution est d'un demi-ton**, et nos énoncés doivent le dire.
5. Deux défauts de l'ancien pipeline sont désormais distingués : le lissage d'ordre fixe, et
   la sélection du pic le plus grave (cas du violoncelle).

## Reproduire

Les quatre tests sont des scripts courts sur données déjà présentes, sans retour aux
échantillons :

| Test | Entrée |
|---|---|
| Quantification en bins | `Etude-Formants/Resultats/formants_*_v3.csv` |
| ρ(cep-160, f0) par instrument | `Etude-Formants/Analyses_formants-mai26/*_formant_results.csv` |
| Croisement des quatre sources | les deux ci-dessus + front matter des fiches |
| Corroboration LPC / cepstre-80 | idem |

À verser dans `Scripts/v7-formants-2026/` sous le nom `verifier_anciennes_mesures.py`.
