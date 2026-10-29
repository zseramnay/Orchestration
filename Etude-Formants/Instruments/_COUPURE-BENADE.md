---
id: coupure-benade
titre: La coupure du réseau de trous — elle se mesure, mais pas chez les flûtes
date: 2026-10-28
tache: C10
script: Scripts/v7-formants-2026/coupure_benade.py
donnees: Etude-Formants/Mesures/Coupure-Benade/
---

# La coupure du réseau de trous latéraux

## Ce qu'on en attendait

`00-METHODE.md` l'annonce depuis le 23 octobre sans l'avoir jamais mesurée : « pour la classe B,
l'invariant pertinent n'est pas un formant mais la fréquence de coupure du réseau de trous
latéraux (Benade) et la pente spectrale ». La fiche de la petite flûte va plus loin : « pas de
formant. Le descripteur à développer est la fréquence de coupure de Benade. »

C'était une promesse faite aux flûtes, qui n'ont aucun formant retenu. **Elle ne peut pas être
tenue**, et cette page explique pourquoi — puis montre où le descripteur fonctionne quand même.

## La mesure

Le réseau de trous latéraux d'un bois se comporte en filtre : sous la coupure il réfléchit,
au-dessus il laisse passer. Les harmoniques sous la coupure sont portés par des résonances d'air
fortes, ceux au-dessus ne le sont plus et l'enveloppe y tombe plus vite. La signature
observable est donc **un coude dans l'enveloppe spectrale**.

Note par note, sur les amplitudes harmoniques en dB contre le log de la fréquence, on ajuste un
**bâton brisé** : deux droites se rejoignant en un point libre. Le coude n'est retenu que si la
pente du haut est plus raide que celle du bas d'au moins 4 dB par octave.

**Le test qui décide** est celui de la tâche C1 : une coupure est une propriété du tuyau, pas de
la note. La pente de log2(coude) contre log2(f0) doit donc être nulle. Si elle vaut 1, le coude
n'est qu'un rang harmonique déguisé.

## Premier essai, et un témoin négatif qui s'effondre

| | Pente contre f0 | Verdict |
|---|---|---|
| **Petite flûte** | 0,79 · 1,08 | suit la note |
| **Flûte** | 1,03 · 1,01 | suit la note |
| **Violon** | **0,13 · 0,12** | **coupure stable** |

**Les flûtes, qui ont un réseau de trous, suivent la note. Le violon, qui n'en a pas, donne une
coupure stable.** C'est l'inverse exact de l'hypothèse, et cela condamne la première version de
la mesure.

Le diagnostic est arithmétique. Une coupure vers 2 kHz n'est mesurable que si **plusieurs
harmoniques l'encadrent**. Or la petite flûte a **sept harmoniques par note** et sa note la plus
grave est à 587 Hz : la moitié de sa tessiture a sa *fondamentale* au-dessus de toute coupure
plausible. Le bâton brisé n'a alors plus qu'un rang harmonique fixe à trouver, qui suit f0 par
construction.

## Second essai, en ne gardant que les notes graves

En restreignant aux notes dont la fondamentale est sous 600 Hz, le tableau se range par
famille — et il se range d'une façon qu'on n'attendait pas.

| Famille | Instruments | Pente médiane contre f0 |
|---|---|---|
| **Clarinettes** | 4 | **0,07** |
| Sans trous latéraux *(cordes, cuivres)* | 10 | 0,47 |
| Anches doubles | 4 | 0,65 |
| **Flûtes** | 2 | **1,22** |

**Seules les clarinettes donnent un coude qui ne bouge pas avec la note.** Et il réplique :

| Instrument | Coude par corpus (Hz) | Étendue | Pente |
|---|---|---|---|
| **Clarinette en Si♭** | 2 094 · 2 078 · 1 970 · 2 013 | **1,1 demi-ton** | 0,27 |
| **Clarinette contrebasse** | 2 116 · 2 218 | **0,8 demi-ton** | 0,07 |
| Clarinette basse | 1 669 · 1 852 · 2 358 · 1 314 | 10,1 dt | **0,07** |
| Clarinette en Mi♭ | 3 348 · 1 986 · 3 607 | 10,3 dt | −0,02 |

**La clarinette en Si♭ à 2 000 Hz, sur quatre corpus de provenances distinctes et à un demi-ton
près, est une grandeur neuve de cette étude.** Et le fait que les quatre clarinettes se tiennent
entre 1 300 et 2 400 Hz **sans s'ordonner par taille** est ce que le modèle des trous latéraux
prédit : la coupure dépend de la géométrie des trous relativement à la perce, et les facteurs
visent une coupure voisine quelle que soit la taille de l'instrument.

## Ce qu'il ne faut pas en conclure

**Un coude stable n'est pas une preuve de coupure de Benade.** Le témoin négatif le dit sans
ambiguïté : le **violon**, qui n'a aucun trou latéral, donne le coude le plus stable de tout le
corpus — pente 0,02, à 3 100–4 200 Hz selon le corpus. C'est vraisemblablement le *bridge hill*
ou la transition de rayonnement de la caisse, mais la mesure ne peut pas le dire.

Ce que la méthode trouve est donc **une transition spectrale fixe**, pas une coupure de réseau
de trous. L'attribution demanderait un levier indépendant, que cette étude n'a pas :

- la coupure devrait disparaître quand on ferme tous les trous (note la plus grave) ;
- elle devrait se déplacer entre un instrument et sa version à perce modifiée ;
- ou se comparer à une mesure d'impédance d'entrée, qui n'est pas dans nos moyens.

Tant que l'un des trois n'est pas fait, le chiffre doit être publié comme **« coude spectral,
compatible avec une coupure de réseau de trous »** et non comme une coupure mesurée.

## La promesse faite aux flûtes doit être retirée

C'est la conclusion qui compte, parce qu'elle corrige deux documents.

| Flûte | Pente contre f0 | Notes utilisables |
|---|---|---|
| Petite flûte | 0,79 · 1,08 | **4** sous 600 Hz, sur 111 |
| Flûte | 0,71 · 1,70 | 39 et 14 |
| Flûte basse | — | **2** coudes nets sur 27 notes graves |
| Flûte contrebasse | 1,23 | 4 et 14 |

Les flûtes cumulent les trois handicaps : **peu d'harmoniques** (7 par note pour la petite
flûte, contre 22 pour la clarinette en Si♭), **une tessiture haute** qui place leurs
fondamentales au-dessus de la région de coupure, et **une enveloppe trop lisse** pour qu'un
coude s'y détache — la flûte basse n'en montre que deux sur vingt-sept notes graves.

Autrement dit : **le descripteur promis aux instruments sans formant est précisément celui
qu'ils ne permettent pas de mesurer.** Il faut cesser de l'annoncer dans `00-METHODE.md` et dans
la fiche de la petite flûte, et chercher ailleurs ce qui décrira les flûtes.

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| **Que publier pour les flûtes, alors ?** | la pente spectrale et son écart-type sont déjà mesurés et répliqués ; c'est peu, mais c'est honnête. La tâche C9 leur a ajouté un trait positif : **elles ne fusionnent avec rien** (14 à 18 %), ce qui est un fait d'orchestration utilisable |
| Attribuer le coude de la clarinette | le mesurer sur la note la plus grave, tous trous fermés : la coupure devrait y disparaître |
| Le coude du violon à 3 100–4 200 Hz | est-ce le *bridge hill* ? Le comparer entre violons avec et sans sourdine, qui charge le chevalet |
| Pourquoi les anches doubles échouent (pente 0,65) | hypothèse à tester : la résonance propre de l'anche fait partie de la source et impose sa propre cassure — ce qui recoupe l'échec du LPC sur cette famille (tâche C4) |

## Journal

- **2026-10-28** — Tâche C10. Le descripteur **existe et se mesure** — la clarinette en Si♭ à
  2 000 Hz sur quatre corpus, à un demi-ton près — mais **pas chez les flûtes, à qui il avait été
  promis**. Le témoin négatif a servi deux fois : d'abord en condamnant la première version de la
  mesure, ensuite en interdisant l'attribution du coude au réseau de trous, puisque le violon,
  qui n'en a pas, donne le coude le plus stable de l'étude.
