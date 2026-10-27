---
id: resonance-de-caisse
titre: Résonances de caisse des cordes — ce que la méthode 12 voit, et où elle ne voit rien
date: 2026-10-27
tache: C1
instruments: alto, violoncelle, contrebasse, violon
corpus: SOL, ZSamples, VO, VN
script: Scripts/v7-formants-2026/resonance_caisse.py
donnees: Etude-Formants/Mesures/Resonance-caisse/
figure: Site-v2/figures/synthese-resonance-caisse.png
---

# Résonances de caisse des cordes

## La question

La carte vocalique avait livré, sur l'alto, un résultat inattendu : au grave, son premier
maximum auditif n'est pas sa fondamentale, et il se tient entre 347 et 498 Hz sur trois corpus.
Nous en avions conclu à la **résonance principale de caisse**, que la littérature situe vers
350–450 Hz pour un alto, et noté que la méthode 8 l'avait entièrement manquée — son
redressement local à ±3/4 d'octave étant aveugle aux résonances plus larges que ~1,5 octave.
**Cette explication s'est révélée fausse le lendemain** : la tâche C2 a montré que la méthode 8
ne mesure simplement rien en dessous de 3,2 à 5,7 fois la fondamentale la plus grave du corpus.
La résonance n'est pas redressée, elle est hors du domaine. Voir `_FENETRE-DE-REDRESSEMENT.md`.

D'où la tâche C1 : *les autres cordes cachent-elles la même chose ?*

## Pourquoi il fallait un test plus dur

L'argument de l'alto reposait sur **un seul contrôle**, la marque `F1=f0` : le premier maximum
n'est pas la fondamentale. Ce contrôle ne suffit pas aux cordes graves.

Au violoncelle, la zone de caisse attendue (A0 vers 100 Hz, B1 vers 180–250 Hz) tombe
exactement sur les harmoniques 2 à 4 des notes du grave : 2 × 98 = 196, 3 × 65 = 195. Un
« F1 stable vers 200 Hz » peut donc n'être que **l'harmonique le plus fort**, qui suit la note
au lieu de rester en place. La marque `F1=f0` ne le verrait pas : elle ne teste que
l'harmonique 1.

## Le test qui tranche

Une résonance ne bouge pas quand la note bouge ; un harmonique, si. On mesure donc, **note par
note**, la pente de log2(F1) contre log2(f0) à l'intérieur d'un registre :

| pente | lecture |
|---|---|
| ≈ 0 | F1 reste en place pendant que f0 se déplace → **résonance** |
| ≈ +1 | F1 suit la note → **harmonique**, pas une propriété de l'instrument |

Le test n'a de force que si f0 se déplace assez : on exige une étendue d'au moins six demi-tons
dans le registre. Ici elle vaut 7 à 11 demi-tons partout.

### Deux garde-fous, posés avant de lire les résultats

**1. F1 est-il un vrai maximum, ou un centroïde déguisé ?** L'échelle de Bark est très
comprimée dans le grave : 3,5 Bark, le seuil de fusion de Chistovich, y vaut environ 0 à
350 Hz. On pouvait craindre que les harmoniques d'une note de contrebasse à 41 Hz (41, 82,
123, 164…) se fusionnent tous en chaîne, auquel cas « F1 » ne serait que le centre de gravité
énergétique des premiers harmoniques, c'est-à-dire un descripteur de pente spectrale.

Le soupçon est levé : le nombre médian de maxima fusionnés dans le groupe qui produit F1 vaut
**1** presque partout, et F1 se tient **2,3 à 5,6 demi-tons sous le centroïde** du motif
d'excitation en dessous de 900 Hz. Ce n'est donc pas le centroïde. Attention toutefois : ce
contrôle ne vaut que dans un sens. Un F1 *éloigné* du centroïde est certainement un trait
distinct ; un F1 *proche* du centroïde n'est pas disqualifié pour autant — il peut simplement
s'agir d'une résonance qui domine toute la bande.

**2. La pente sur les corpus réunis peut mentir.** Trois corpus qui ne s'accordent pas sur le
niveau peuvent produire, réunis, une pente plate qu'aucun ne donne séparément. **C'est le
tableau par corpus qui tranche, jamais la droite de la figure.** La contrebasse en est le cas
d'espèce : pente réunie −0,01, pentes par corpus de −0,70 à +0,84.

## Résultats

Quatre corpus de provenances distinctes : SOL (IRCAM), ZSamples, VO, VN (VSL).

| Instrument | Registre | Corpus | F1 médians (Hz) | Étendue | Pente médiane | F1 − centroïde | Conclusion |
|---|---|---|---|---|---|---|---|
| Alto | grave | 4 | 374 · 375 · 387 · 395 | **0,9 dt** | +0,01 | −1,5 dt | résonance, accord le plus serré de l'étude |
| Alto | médium | 4 | 351 · 423 · 450 · 470 | 5,1 dt | **+0,95** | +0,8 dt | **F1 suit la note** |
| **Violoncelle** | **médium** | **4** | **209 · 217 · 223 · 223** | **1,1 dt** | **−0,04** | −3,7 dt | **RÉSONANCE ÉTABLIE** |
| Violoncelle | grave | 4 | 178 · 193 · 196 · 238 | 5,0 dt | −0,96 | −4,0 dt | non établi |
| Contrebasse | grave | 3 | 106 · 112 · 122 | 2,4 dt | −0,41 | −5,6 dt | non établi |
| Contrebasse | médium | 3 | 106 · 184 · 192 | 10,3 dt | +1,05 | −3,2 dt | F1 suit la note |
| Violon | grave | 4 | 363 · 461 · 467 · 532 | 6,6 dt | −0,07 | −0,3 dt | non établi, dispersion inter-corpus |
| Violon | médium | 4 | 495 · 496 · 525 · 554 | 1,9 dt | +0,03 | −0,3 dt | compatible, mais F1 ≈ centroïde |

> Ces valeurs ne sont **pas** sur la grille tempérée. Contrairement à la méthode 8, dont l'axe
> composite est tempéré par construction et dont la résolution est d'un demi-ton, la méthode 12
> lit un maximum sur une échelle continue. On écrit donc un intervalle, pas une note.

## Ce qui est acquis

### Le violoncelle a une résonance de caisse vers 215–220 Hz

**Quatre corpus de provenances distinctes, F1 entre 209 et 223 Hz — 1,1 demi-ton d'écart —
pente médiane −0,04 sur onze demi-tons de fondamentale.** Dans le registre médium (Sol2–Sol3,
f0 de 104 à 196 Hz), F1 ne bouge pas pendant que la note parcourt presque une octave. Le rang
harmonique médian passe de 2,1 en bas du registre à 1,1 en haut : la même résonance est
excitée par l'harmonique 2 puis par l'harmonique 1. C'est exactement la signature attendue.

La valeur tombe dans la zone B1 que la littérature acoustique donne au violoncelle
(180–250 Hz).

**Réserve, à écrire avec le résultat.** 220 Hz est aussi la corde de La à vide du violoncelle,
et les luthiers accordent volontiers le mode B1+ au voisinage de cette corde. La mesure établit
qu'il y a là une résonance qui ne suit pas la note ; elle **ne sépare pas** le mode de caisse
de la résonance sympathique de la corde à vide. Pour l'orchestrateur la conséquence pratique
est la même — une zone de renforcement fixe vers 215–220 Hz quelle que soit la note du médium —
mais l'attribution physique reste ouverte.

### Et la méthode 8 ne voyait rien là non plus

Le violoncelle a quatre formants retenus par la méthode 8, dont 1 108 Hz par route A. Aucun
dans la zone de 215 Hz. Le second cas, après l'alto, où **la méthode 12 voit une résonance que
la méthode 8 manque par construction** : la zone de caisse est large et plate, donc redressée
avec le reste par `residu_local()`. Cela confirme la hiérarchie établie sur l'alto, et renforce
l'intérêt de la tâche C2 — élargir la fenêtre de redressement.

## Ce qui est corrigé

### L'alto : le grave tient, le médium tombe

`_CARTE-VOCALIQUE.md` affirmait « F1 entre 347 et 498 Hz **au grave et au médium** sur trois
corpus … compatible avec la résonance principale de caisse ». **La moitié médium est fausse.**
Au médium, les quatre corpus donnent des pentes de +0,79, +0,85, +1,04 et +1,11 : F1 y est
l'harmonique le plus fort, qui suit la note. La médiane de la cellule n'était qu'une moyenne de
hauteurs jouées.

Le grave, lui, sort **renforcé** du test :

- quatre corpus et non trois ;
- un intervalle de **374 à 395 Hz** et non de 347 à 498 — 0,9 demi-ton, l'accord inter-corpus
  le plus serré de toute l'étude ;
- pente médiane +0,01, et les deux corpus qui s'écartent le font **en sens opposés** (−0,58 et
  +0,67), ce qui est du bruit d'estimation et non une tendance ;
- 0 % de cellules marquées `F1=f0` dans les quatre corpus.

La formulation juste est donc : *l'alto présente, dans son registre grave seulement, une
résonance fixe vers 375–395 Hz*, compatible avec son mode B1+.

## Ce qui n'est pas établi, et pourquoi

### La contrebasse — trois raisons, toutes identifiables

1. **Les corpus ne s'accordent pas sur le niveau.** Au grave : 106, 112, 122 Hz, soit
   2,4 demi-tons. Au médium, SOL donne 106 Hz avec une pente de +1,41 et 70 % de cellules
   `F1=f0` — chez SOL, F1 au médium *est* la fondamentale — pendant que ZSamples et VO donnent
   184 et 192 Hz.
2. **Un tiers du registre grave est sous le plancher du calcul.** Le masque du motif
   d'excitation commence à 50 Hz ; or **31 des 91 notes du registre grave, soit 34 %, ont une
   fondamentale en dessous de 50 Hz** (Mi1 = 41 Hz). Pour ces notes la fondamentale est
   partiellement exclue du motif, ce qui pousse F1 vers le haut. Le registre grave de la
   contrebasse est donc hors du domaine de validité de la méthode 12 telle qu'elle est réglée.
3. **Les modes attendus coïncident avec les fondamentales.** La littérature place le corps de
   la contrebasse vers 70–110 Hz, c'est-à-dire *à l'intérieur* de son propre registre médium.
   Résonance et fondamentale ne sont alors pas séparables par un détecteur de maximum.

### Le violon — compatible, mais non concluant

Au médium, quatre corpus entre 495 et 554 Hz (1,9 dt), pente médiane +0,03 : c'est compatible
avec le mode B1+ que la littérature situe vers 530–550 Hz. Mais F1 n'y est qu'à 0,3 demi-ton du
centroïde, et 62 % des cellules sont marquées `F1=f0`. On ne peut pas distinguer la résonance
de la hauteur jouée. Au grave, les quatre corpus se dispersent sur 6,6 demi-tons (363 à
532 Hz) — à rapprocher de l'instabilité du violon déjà signalée au suraigu (**tâche C3**).

## La régularité qui se dégage

La méthode 12 détecte une résonance de caisse **quand le registre grave de l'instrument se
trouve en dessous de son mode principal**, et seulement là.

| Instrument | Registre grave (f0) | Mode principal attendu | Le grave est-il sous le mode ? | Résultat |
|---|---|---|---|---|
| Alto | 131–196 Hz | B1+ ~ 350–450 | **oui**, d'une octave | résonance vue au grave |
| Violoncelle | 65–98 Hz · médium 104–196 | B1 ~ 180–250 | **oui**, d'une octave | résonance vue au médium |
| Violon | 196–294 Hz | B1+ ~ 530–550 | à peine | compatible, non concluant |
| Contrebasse | 41–65 Hz | corps ~ 70–110 | le mode est *dans* le médium | rien |

C'est une propriété du rapport entre tessiture et taille de la caisse, pas une propriété du
corpus. Elle prédit où la méthode marchera, et c'est à ce titre qu'elle est utile : inutile de
chercher une résonance de caisse par cette voie là où la tessiture la recouvre.

## Conséquences pour les documents

- `_CARTE-VOCALIQUE.md` : retirer « et au médium » de l'affirmation sur l'alto, et resserrer
  l'intervalle à 374–395 Hz sur quatre corpus.
- Fiche de l'alto : même correction.
- Fiche du violoncelle : ajouter la résonance vers 215–220 Hz, avec la réserve sur la corde de
  La à vide.
- `00-METHODE.md` : noter le plancher de 50 Hz du masque comme limite de validité, et la
  régularité tessiture/caisse ci-dessus.

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| Mode de caisse ou corde de La à vide, au violoncelle ? | mesurer sur des notes de la corde de Do et de Sol uniquement, corde de La étouffée — ou comparer avec et sans sourdine, qui déplace les modes de caisse mais pas la corde |
| La contrebasse a-t-elle une résonance mesurable ? | abaisser le plancher du masque sous 50 Hz et refaire ; sans quoi la question restera sans réponse par cette voie |
| Le violon ? | la **tâche C3** d'abord : tant que le détecteur bascule entre fondamentale et caisse, la ligne du violon n'est pas publiable |
| La méthode 8 verrait-elle ces zones avec un redressement plus large ? | **tâche C2** — c'est maintenant le test le plus intéressant du bloc C, puisque deux résonances réelles lui échappent |

## Journal

- **2026-10-27** — Première version. Tâche C1 exécutée sur quatre cordes et quatre corpus,
  29 à 30 lignes de mesure selon le lot. Un résultat acquis (violoncelle), une affirmation
  publiée corrigée de moitié (alto), deux non-résultats dont les causes sont identifiées.
