---
id: chapitre-synthese
titre: Synthèse — ce que l'orchestre partage, et ce qu'on peut en dire
date: 2026-10-27
revision: >
  27/10 — intégration des six travaux du bloc C : résonances de caisse,
  plancher du domaine, plancher de la fondamentale, corroboration LPC,
  effet de la nuance sur la localisation, et la partie III sur les mélanges.
---

# Synthèse

## Ce que cette étude peut dire, et ce qu'elle ne peut pas

Elle mesure **où se trouvent les résonances fixes des instruments**, à partir de cinq corpus
d'échantillons indépendants, et elle ne retient un résultat que si deux corpus de provenances
différentes le donnent au même point de grille. Vingt-trois formants satisfont ce critère, sur
treize instruments.

Elle mesure aussi les **mélanges**, et c'est neuf depuis le 27 octobre. Non pas en simulant :
le motif d'excitation auditive est linéaire en puissance, donc le motif d'un mélange est
exactement la somme des motifs de ses parents. Deux cent quinze paires d'instruments ont été
calculées de cette façon. Mais ce qui est mesuré là est le mélange **spectral**, qui est une
condition nécessaire de la fusion timbrale et non sa totalité : il ignore la phase, l'espace,
et le fait que des musiciens qui jouent ensemble s'ajustent.

Elle a deux planchers, tous deux mesurés plutôt que supposés, et tous deux découverts en
corrigeant une explication fausse.

**Le plancher du domaine.** La méthode 8 ne mesure rien en dessous de **3,2 à 5,7 fois la
fondamentale la plus grave du corpus**, parce qu'un harmonique n'entre dans l'enveloppe
composite que si sa fenêtre contient trois harmoniques de la même note. Nous avions longtemps
attribué ses angles morts à sa largeur de fenêtre ; c'est son plancher qui est en cause. C'est
lui, et non le redressement, qui lui fait manquer les résonances de caisse de l'alto et du
violoncelle.

**Le plancher de la fondamentale.** Un spectre harmonique n'a aucun partiel sous f0 : un maximum
situé là est du bruit d'archet ou la jupe de la fonction d'étalement, jamais un formant. La
méthode 12 ne l'excluait pas, et basculait entre fondamentale et caisse dès que la note montait
au-dessus de la zone de résonance — un tiers des notes du suraigu du violon avaient un « F1 »
sous leur propre fondamentale.

Cette réserve n'est pas de la prudence rhétorique. L'étude antérieure annonçait des doublures
« vérifiées » à partir de coïncidences entre des valeurs qui se sont révélées être des
fréquences médianes de tessiture tombant dans des bins de FFT voisins. Le risque n'est pas
d'avoir de mauvais chiffres : c'est d'avoir de bons chiffres et d'en tirer une conclusion
qu'ils ne portent pas.

Ce qui est écrit ci-dessous comme convergence reste **un fait de mesure**, et ce qui en est tiré
pour l'orchestration une **hypothèse**. La partie III, sur les mélanges, est la seule à porter
sur des combinaisons ; elle ne valide pas les convergences, elle répond à une autre question.

---

# I. Les convergences

<!-- TABLEAUX-CONVERGENCES -->

## Ce que signifie partager un formant

Un formant est une région de fréquence que l'instrument renforce **quelle que soit la note
jouée**. Quand deux instruments en partagent un, deux choses se produisent.

À l'**unisson ou à l'octave**, les partiels qui tombent dans cette région sont renforcés chez
les deux. Le spectre résultant n'est pas la somme de deux timbres : c'est un seul timbre dont
l'énergie est doublement appuyée au même endroit. C'est la condition acoustique classiquement
associée à la **fusion**.

À des **hauteurs différentes**, le formant colore quand même les deux de la même façon,
puisqu'il ne dépend pas de la note. Il en résulte une parenté de timbre qui ne tient pas à
l'unisson — ce que les traités appellent, sans l'expliquer, une « bonne association ».

## Basson et violoncelle à 1 108 Hz — le cas le plus net

C'est la convergence la mieux établie de l'étude, et c'est aussi la doublure la plus
anciennement pratiquée du répertoire.

| | Formant | Établi par | Corroboré par |
|---|---|---|---|
| **Basson** | 1 108 Hz (Do♯6) | Xsamples + ZSamples, écart nul | largeur mesurée : Q ≈ 10 |
| **Violoncelle** | 1 108 Hz (Do♯6) | SOL + VN, écart nul | largeur mesurée : Q ≈ 6 |

Deux instruments, deux familles, quatre corpus différents, **la même valeur au point de grille
près**. Et ce sont deux de nos formants les mieux localisés : l'enveloppe vraie y montre un pic
réel dans les deux cas, ce qui n'est pas vrai partout.

Un point à ne pas escamoter : **la coïncidence est dans le résultat, pas dans la cause.** Chez
le basson, 1 108 Hz relève de la perce et de l'anche ; chez le violoncelle, d'un mode de caisse.
Rien ne les relie physiquement. Ce qui les relie est ce que l'oreille reçoit — et pour
l'orchestration, c'est précisément ce qui compte.

Le basson a par ailleurs un second formant établi à **1 974 Hz**, et le violoncelle à
**3 948 Hz** : au-delà de la zone commune, les deux instruments divergent.

## Clarinette basse et tuba basse à 1 244 Hz

Même situation formelle — deux instruments, deux familles, écart nul — mais corroborée
différemment. La clarinette basse tient son 1 244 Hz de SOL et Xsamples ; le tuba basse de VN et
ZSamples, sans que SOL le voie.

Les deux sont des instruments graves dont on se sert comme **fondation**, et la zone partagée
est haute par rapport à leur tessiture : vers 1 250 Hz, on est dans leurs harmoniques supérieurs,
non dans leur fondamentale. Un renforcement là ne pèse pas sur la fondation : il ajoute de la
définition. Hypothèse à tester, mais elle est cohérente avec l'usage.

## Clarinette contrebasse et trombone basse à 1 759 Hz

Le troisième unisson, et le plus fragile des trois — non par la qualité de la mesure, mais parce
que la combinaison est rare au répertoire et qu'on manque donc de corroboration par l'usage.
La clarinette contrebasse le tient de SOL + VO, le trombone basse de SOL + ZSamples + VO.

## Les régions partagées

Cinq zones réunissent plusieurs instruments à un demi-ton près. **Ce ne sont pas des unissons**,
et il faut le dire : la résolution de la méthode étant d'un demi-ton, deux valeurs qui en
diffèrent d'un sont indistinguables de deux valeurs identiques *ou* de deux valeurs réellement
distinctes. Les présenter comme des unissons serait refaire, un cran plus haut, l'erreur des
convergences à Δ = 11 Hz.

Deux d'entre elles méritent qu'on s'y arrête.

### 2 216–2 348 Hz — clarinette basse, cor, violon

Trois instruments, trois familles. Le membre le plus intéressant est le violon : **2 348 Hz est
son *bridge hill***, la résonance du chevalet que la littérature acoustique situe entre 2 et
3 kHz. C'est la seule de nos mesures qui dispose d'une validation externe publiée.

Que le cor s'y trouve n'est pas anodin : cor et cordes est un mariage que tous les traités
recommandent et qu'aucun n'explique. Cette région commune en est une explication possible.

### 3 748–3 959 Hz — alto, violon, violoncelle

Toute la famille des cordes aiguës, dans une région de moins d'un demi-ton d'étendue. Trois
instruments de tailles très différentes y trouvent un formant établi.

C'est acoustiquement remarquable, et c'est à verser au dossier plutôt qu'à conclure : une
résonance qui ne se déplace pas avec la taille de l'instrument ne peut pas être un mode de la
caisse, qui lui est proportionnel. Elle pointe vers quelque chose que le violon, l'alto et le
violoncelle ont en commun indépendamment de leurs dimensions — le chevalet, l'épaisseur des
tables, ou la nature du bois. **Nous ne pouvons pas trancher avec les mesures en main.**

Cette région n'apparaissait pas avant le 26 octobre, l'alto étant alors cru sans formant.

---

# II. Les propriétés mesurées

Les convergences disent ce que des instruments ont en commun. Cette section dit ce que nous
avons mesuré **de chaque instrument**, indépendamment des autres — et c'est souvent plus solide.

## 1. La typologie, et pourquoi ρ est le descripteur le plus fiable

ρ mesure la corrélation entre la position du pic de l'enveloppe et la note jouée.

| ρ | Lecture | Classe |
|---|---|---|
| proche de 0 | le spectre ne suit pas la hauteur : il y a une résonance fixe | **A** |
| proche de 1 | le spectre suit la hauteur : il n'y a pas de résonance fixe | **B** |
| élevé, mais avec des formants à l'agrégation | résonances multiples dont aucune ne domine | **C** |

Sur les cinq corpus et vingt et un instruments, **l'étendue médiane de ρ entre corpus n'est que
de 0,14**. C'est beaucoup plus stable que les positions de formants, qui changent souvent de
plusieurs demi-tons d'un corpus à l'autre.

Autrement dit : **nous savons mieux dire si un instrument a un formant que dire où il est.**

Deux cas opposés valent la peine :

**Le contrebasson** est l'instrument au ρ le plus stable du corpus — 0,31 à 0,43 sur quatre
corpus, tous le classent A. Et pourtant **aucune de ses fréquences ne réplique**. Son
comportement réplique, ses valeurs non.

**Le hautbois** est l'inverse : son ρ varie de 0,26 à 0,80, l'étendue la plus forte mesurée. Le
seul corpus qui en fait un instrument à formant est celui de l'IRCAM ; les trois autres en font
un spectre suiveur.

## 2. La dynamique déplace le spectre des cuivres de près d'une octave

C'est le résultat le plus robuste de la famille.

| Instrument · registre | doux | moyen | fort | facteur |
|---|---|---|---|---|
| **Trombone ténor · grave** | 355 Hz | 531 Hz | **718 Hz** | ×2,02 |
| Tuba basse · médium | 129 Hz | 214 Hz | 350 Hz | ×2,71 |
| Trompette · médium | 639 Hz | 880 Hz | 1 073 Hz | ×1,68 |

Le trombone ténor a été remesuré sur un corpus indépendant : **359 → 707 Hz**, contre 355 → 718.
Les deux concordent à moins d'un demi-Bark.

Conséquence pour l'orchestration : **chez un cuivre, la nuance n'est pas un réglage de volume,
c'est un changement de timbre.** Un *forte* ne sonne pas comme un *piano* plus fort ; il sonne
une octave plus haut dans sa région d'énergie. Une doublure accordée sur le timbre en *piano*
ne l'est plus en *forte*.

Conséquence pour la méthode : toute mesure de cuivre agrégée sur plusieurs nuances est
nécessairement étalée. C'est ce qui explique que le cor à 466 Hz et la trompette à 1 479 Hz
soient longtemps restés « fragiles » alors qu'ils correspondent à la littérature.

### La nuance ne déplace pas seulement le spectre, elle décide s'il y a un formant

Mesuré le 27/10 sur la part de pics **non localisables** — ceux dont l'enveloppe ne redescend
pas de 3 dB des deux côtés, donc sans largeur mesurable.

| Famille | Effet doux → fort |
|---|---|
| **Cuivres** | **+31 points** |
| Clarinettes | +10 |
| Cordes | +3 |
| Anches doubles | +1 |

C'est un effet de cuivres, et de cuivres seulement. Mais deux sur cinq s'y opposent, et **le
partage suit la perce** : à registre constant, le trombone ténor passe de 69 % à 7 % de pics non
localisables dans le grave (+63 points) et la trompette de 58 % à 6 % (+52), tandis que le cor
va de 33 % à 47 % et le tuba basse de 39 % à 68 % dans ce même registre. Les perces
majoritairement cylindriques confirment, les coniques non — ce qui est la géométrie même du
cuivrage, qui a besoin d'une longueur cylindrique pour se développer.

Trois chiffres à retenir pour l'orchestration :

- **le formant du trombone basse n'existe pratiquement pas au piano** — 76 % de ses pics y sont
  non localisables contre 22 % au forte. Écrire « le trombone basse a un formant à 879 Hz » sans
  la nuance, c'est décrire un instrument joué fort ;
- **le médium du tuba basse n'est localisable à aucune nuance** : 78 à 79 %. Il n'y a pas là de
  formant à publier ;
- **le cor se comporte à l'envers** dans le grave et le médium : joué fort, il devient *moins*
  localisable. À ne pas décrire comme les autres cuivres.

Détail dans `_LARGEUR-PAR-REGISTRE.md`.

## 3. La géométrie du tuyau se lit dans le spectre

La balance entre harmoniques impairs et pairs est positive chez toutes les clarinettes, et
**elle croît en descendant la famille** :

| Clarinette | Mib | Sib | basse | contrebasse |
|---|---|---|---|---|
| Balance impair − pair | +0,7 dB | +4,2 dB | +4,5 dB | +5,1 dB |

C'est la signature du **tuyau cylindrique fermé**, qui ne produit en théorie que les harmoniques
impairs. Plus l'instrument est grand, plus la perce est longue devant son diamètre, plus la
signature est nette. Aucun autre instrument de l'orchestre ne présente ce profil : tous les
autres sont à valeur négative, c'est-à-dire à série complète.

C'est, avec le ρ, le descripteur le mieux ancré dans la physique, et il ne demande aucune
hypothèse sur les formants.

## 4. Un formant étroit sélectionne, un formant large colore

Nous avons mesuré les largeurs de bande par deux méthodes indépendantes. Les deux ne
s'accordent qu'une fois sur dix-neuf, et cette discordance est instructive — mais l'ordre de
grandeur, lui, se dégage :

| | Q | Largeur | Instruments |
|---|---|---|---|
| **Formants étroits** | 10 à 29 | 0,03–0,11 octave | violon · violoncelle aigu · clarinette en Sib · basson |
| **Formants larges** | 3,4 à 6,1 | 0,21–0,34 octave | tuba basse · les deux trombones |

Un formant de Q = 29 — le violon vers 3 950 Hz — **sélectionne** : seuls les partiels qui
tombent dedans sont renforcés, et c'est la note jouée qui décide lesquels. Un formant de
Q = 3,4 — le tuba basse — **colore** : il renforce une région entière quelle que soit la note.

D'où un critère pratique : **un formant large donne un effet stable sur toute la tessiture, un
formant étroit donne un effet qui dépend de la note, donc irrégulier le long d'une phrase.**

Et une mise en garde. Chez la **trompette**, la mesure sur l'enveloppe vraie donne une largeur
égale à la fréquence centrale — Q ≈ 1. Cela signifie qu'il n'y a **aucun pic localisé** : ce
qui réplique à 2 792 Hz est une région d'énergie, pas une résonance. Le modèle LPC y place
pourtant un pôle étroit, parce qu'il y est contraint. Même situation au trombone basse vers
1 670 Hz.

## 5. L'espace vocalique couvre ce que les formants ne couvrent pas

Les formants n'établissent quelque chose que pour treize instruments sur vingt-deux. La carte
vocalique les couvre tous, parce qu'elle ne suppose aucune résonance fixe : elle lit la
saillance du motif d'excitation auditive et la compare aux voyelles du français.

**Quatre comportements :**

| | Instruments |
|---|---|
| Voyelle unique et stable sur toute la tessiture | **le basson**, seul cas |
| Traverse l'espace en montant, puis en sort par le haut | flûtes · clarinettes · cordes · hautbois |
| Hors de l'espace dès le médium ou l'aigu | **petite flûte** · flûte · clarinettes Mib et Sib au suraigu |
| Hors de l'espace par le bas, ou sans second maximum | contrebasson · les deux tubas · cor et trombones hors de l'aigu |

**La frontière tombe vers 700 Hz de F1**, qui est exactement le plafond de /a/ chez l'homme —
684 Hz. Au-delà, aucune voyelle française n'est atteignable : il n'y a rien au-dessus. C'est
arithmétique, et c'est ce qui explique que le piccolo ne ressemble à aucune voyelle.

Les clarinettes offrent la régularité la plus frappante : **chacune monte de /u/ vers /a/** en
gravissant ses registres, puis sort par le haut. Et plus l'instrument est grave, moins il monte.
Une même trajectoire, décalée selon la taille — ce qui recoupe exactement l'écoute : chalumeau
sombre et fermé, clarine ouvert.

## 6. Les cordes ont une résonance de caisse, et la méthode 8 ne la voit pas

C'est le cas où les deux méthodes ne sont pas seulement complémentaires mais **hiérarchisées**.

| Instrument | Résonance | Corpus | Écart | Pente |
|---|---|---|---|---|
| **Violoncelle**, médium | **209 · 217 · 223 · 223 Hz** | quatre | 1,1 demi-ton | −0,04 |
| **Alto**, grave | **374 · 375 · 387 · 395 Hz** | quatre | 0,9 demi-ton | +0,01 |

La pente est celle de log2(F1) contre log2(f0) à l'intérieur du registre : **plate, elle dit que
F1 ne bouge pas pendant que la note parcourt presque une octave.** C'est le test qui distingue
une résonance d'un harmonique, et il manquait à la première version de ce résultat : la marque
`F1=f0` ne teste que l'harmonique 1, or au violoncelle la zone de caisse attendue tombe
exactement sur les harmoniques 2 à 4 du grave.

Appliqué sérieusement, ce test a **corrigé de moitié** ce que nous avions publié sur l'alto : au
médium, les quatre corpus donnent des pentes de +0,79 à +1,11 — F1 y suit la note. Le grave, lui,
en sort renforcé : quatre corpus au lieu de trois, et 374–395 Hz au lieu de 347–498.

Aucun des formants retenus par la méthode 8 n'est dans ces zones, et on sait maintenant pourquoi :
elles sont **sous le plancher de son domaine**. Au violoncelle, 215 Hz pour un plancher à
261–329 Hz selon le corpus.

Une réserve à publier avec le résultat du violoncelle : 220 Hz est aussi sa corde de La à vide, et
les luthiers accordent volontiers le mode de caisse au voisinage de cette corde. La mesure établit
qu'il y a là une résonance qui ne suit pas la note ; elle ne sépare pas le mode de caisse de la
résonance sympathique. Pour l'orchestrateur la conséquence est la même.

La contrebasse et le violon ne donnent rien, et leurs raisons sont nommées dans
`_RESONANCE-DE-CAISSE.md`. Il s'en dégage une règle utile : **la méthode 12 voit la caisse quand
le registre grave de l'instrument se trouve sous son mode principal, et seulement là.** Alto et
violoncelle, d'une octave ; violon, à peine ; contrebasse, pas du tout, ses modes étant à
l'intérieur de son propre médium. C'est un rapport entre tessiture et taille de caisse, pas une
propriété de corpus — donc cela prédit où la méthode marchera.

## 7. La corroboration par le LPC tient pour quatre familles sur cinq

La route B disait « corroboré par deux familles de méthodes indépendantes ». C'était un indice,
pas un test : avec six à douze pôles répartis sur 6 kHz, un pôle tombe près de n'importe quelle
fréquence assez souvent. En appariant **note par note** — les deux analyses portent sur les mêmes
fichiers — et en comparant chaque formant au rang qu'il occupe dans son propre spectre :

| Famille | Rang médian du formant |
|---|---|
| Flûtes | 82,6 % |
| **Cordes** | **79,2 %** |
| Clarinettes | 78,3 % |
| Cuivres | 69,3 % |
| **Anches doubles** | **15,4 %** |

**Les six formants des anches doubles sont tous sous la médiane**, quatre sous le 30ᵉ centile. Le
basson place trois de ses quatre formants dans les 6 % les plus bas de son propre spectre.

C'est ce que le modèle du LPC prédit. Il suppose une source à spectre plat filtrée par une
cavité — vrai d'une corde, où corde et caisse sont séparables ; faux d'une anche double, où
l'anche et la perce sont couplées et où la résonance de l'anche fait partie de la source. La
limite était écrite dans nos notes de méthode ; elle est maintenant mesurée.

**Conséquence à tenir** : les formants du basson, du hautbois, du cor anglais et du contrebasson
restent établis par la route A, qui est le critère le plus exigeant — mais ils ne peuvent pas
revendiquer de corroboration LPC, et les fiches ne doivent pas le laisser croire.

Appliqué aux trois unissons, membre par membre :

| Unisson | Membres | Rangs |
|---|---|---|
| **1 244 Hz** | clarinette basse · tuba basse | 75,6 % · 74,2 % |
| 1 108 Hz | violoncelle · **basson** | 67,6 % · **29,4 %** |
| 2 216 Hz | violon · **cor** | 70,8 % · **30,8 %** |

**Celui de 1 244 Hz est le seul dont les deux membres sont corroborés.** Les deux autres sont
asymétriques, et le membre faible s'explique chaque fois.

---

# III. Les mélanges

C'est la question qui a lancé la reprise : **quand deux instruments jouent ensemble, le mélange
produit-il une couleur nouvelle, ou l'un ne fait-il que renforcer l'autre ?**

## La question posait deux issues ; la mesure en impose quatre

| Issue | Lecture |
|---|---|
| **homogène** | les deux ont **déjà** la même couleur ; le mélange la garde |
| **fusion** | la couleur du mélange n'appartient à aucun des deux — elle est neuve |
| **renforcement** | le mélange se lit comme l'un des deux ; l'autre apporte de l'énergie |
| **juxtaposition** | le mélange porte plus de maxima que son parent le plus riche : on entend les deux |

La plus fréquente manquait à la question. Deux instruments de même couleur, c'est **la forme la
plus forte du fondu**, par identité et non par émergence — flûte et clarinette en sont l'exemple
d'école, et notre premier classement les rangeait avec les mesures manquantes.

## Une seule grandeur ordonne presque tout

L'écart de couleur entre les deux instruments **pris séparément**, mesuré dans le plan (F1, F2)
en Bark.

| Écart (Bark) | Fusion | Homogène | Renforcement |
|---|---|---|---|
| 0,0 – 0,5 | 8 % | **66 %** | 13 % |
| 1,0 – 1,5 | 14 % | 39 % | 28 % |
| 2,0 – 2,5 | 30 % | 22 % | **32 %** |
| 3,5 et plus | **55 %** | 14 % | 16 % |

Trois mouvements monotones : la fusion croît avec l'écart, l'homogénéité décroît symétriquement,
et le renforcement **culmine au milieu** — il faut que les deux couleurs soient assez proches
pour que l'une absorbe l'autre, et assez distinctes pour qu'il y ait domination.

**Il n'y a donc pas deux familles d'alliages mais un continuum.** Qui cherche un fondu choisit
deux instruments de même couleur ; qui cherche une couleur neuve en choisit deux éloignées ; le
renforcement est ce qu'on obtient entre les deux.

## Les paires

| Les plus fusionnantes | | Les plus homogènes | |
|---|---|---|---|
| trompette + tuba basse | 89 % | flûte + petite flûte | 75 % |
| hautbois + trombone ténor | 60 % | clarinette basse + clarinette | 72 % |
| trombone basse + violoncelle | 55 % | contrebasson + cor | 71 % |
| cor + trompette | 54 % | clarinette + cor | 61 % |

Les familles se fondent avec elles-mêmes, ce qui rassure sur la mesure. Mais la liste contient
aussi des paires que rien ne rapproche sur le papier — **clarinette en Mi♭ et tuba contrebasse à
0,15 Bark**, aux deux extrémités de l'orchestre.

Et là où il y a renforcement, la mesure dit qui l'emporte : **le violon sur la flûte, le cor
anglais sur la flûte**. On le lit dans tous les traités ; on le mesure ici.

## Deux faits de famille

**Les anches doubles fusionnent avec tout (30 à 34 %), les flûtes avec rien (14 à 18 %).** Le
hautbois et le basson sont les meilleurs agents de couleur nouvelle de l'orchestre ; la flûte se
fait absorber plutôt qu'elle ne transforme. Cela rejoint par une tout autre voie ce que la
partie IV dira des flûtes : pas de formant propre. **Un instrument sans couleur propre ne peut
pas en faire naître une nouvelle.**

**Les cuivres fusionnent fortement entre eux (31 %)** alors qu'ils forment une famille. C'est le
résultat le plus contre-intuitif de la table, et l'explication est dans l'écart : trompette et
tuba sont à 4,3 Bark l'un de l'autre, plus loin que beaucoup de paires inter-familles. **La
famille d'orchestration n'est pas une famille de couleur.**

## Ce que cette partie ne dit pas

**Unisson seulement.** Les mélanges sont calculés à hauteur identique. Le doublage à l'octave,
qui est le cas le plus fréquent en orchestration, n'est pas couvert — et il changerait
probablement beaucoup, l'instrument aigu ne recouvrant alors plus les maxima graves de l'autre.
Un seul corpus ; deux instruments à la fois, un tutti n'étant pas une somme de paires ; et le
spectral n'est pas le perceptif.

Détail et table complète dans `_MELANGES.md`.

---

# IV. Ce qui n'est pas mesurable, et c'est un résultat

Une étude qui ne dit pas où elle échoue n'est pas vérifiable.

## La petite flûte n'a aucun formant, dans trois corpus

C'est la réponse à l'anomalie qui a déclenché toute cette reprise. L'ancienne étude s'étonnait
que le piccolo ne converge avec rien, alors que le répertoire regorge de doublures à l'octave
qui lui doivent leur brillance.

La raison est qu'**il n'a pas de formant à faire converger.** Trois corpus indépendants n'en
trouvent aucun, avec un ρ de 0,96 et une pente de −21 dB/octave de part et d'autre. Son spectre
suit entièrement la hauteur.

Un résultat négatif répliqué ne souffre pas du problème des comparaisons multiples : il n'y a
rien à apparier par hasard. C'est, avec le basson, l'acquis le plus sûr de l'étude.

Ce que le piccolo apporte à une doublure n'est donc pas un formant commun, mais de l'énergie
dans une région où les autres n'en ont plus. **Le mécanisme de la brillance n'est pas celui de
la fusion**, et c'est précisément la distinction que la méthode 12 sur mélanges réels doit
établir.

## Le hautbois non plus, et c'est contre-intuitif

Quatre corpus, quatre réponses incompatibles : 4 183 Hz, aucun, 2 635 et 4 974, 3 320. Pas deux
valeurs à moins de quatre demi-tons. Et un ρ de 0,72 à 0,80 dans trois corpus sur quatre.

Il faut l'assumer : le hautbois est l'instrument auquel la littérature assigne un formant avec
le plus d'assurance — Backus 1 300 Hz, Meyer et Giesler ~1 100 Hz. **Nos quatre corpus ne le
confirment pas.** Sa description passe par la carte vocalique, qui elle réplique : /a/–/œ/ au
médium dans trois corpus, et hors de l'espace vocalique dès l'aigu, à 991 Hz au Hz près dans
trois corpus.

## Les sourdines ne sont pas génériques

Dix-huit paires comparables, **écart médian de 4,0 demi-tons**. Ce n'est probablement pas un
échec de mesure mais un fait : une sourdine est un objet physique dont le matériau, la forme et
l'ajustement au pavillon varient d'un exemplaire à l'autre.

Il faut donc parler de « la sourdine sèche de cet instrument-là », non de « la sourdine sèche ».

## Et deux instruments sur lesquels nous ne pouvons rien dire

Le **tuba contrebasse** et le **contrebasson** n'ont aucun second corpus indépendant : les deux
banques qui les contiennent se sont révélées être la même bibliothèque. Tout ce que l'étude en
dirait reposerait sur une source unique.

---

# V. Ce que cela change pour l'enseignement

L'ancienne version de cette étude donnait des matrices de convergence complètes et un tableau
de doublures vérifiées. Celle-ci donne moins, et autrement.

**Ce qu'on perd** : les matrices. Elles reposaient sur des valeurs qui n'étaient pas des
formants, et leurs convergences étaient des bins de FFT voisins.

**Ce qu'on gagne** : trois unissons formantiques et cinq régions partagées, chacun validé
indépendamment sur des corpus de provenances distinctes ; une typologie qui dit, pour chaque
instrument, si la notion de formant s'y applique ; un effet de la dynamique chiffré — en position
*et* en localisation ; une signature de la géométrie du tuyau, qui se retrouve jusque dans l'effet
de la nuance chez les cuivres ; une carte vocalique qui couvre les vingt-deux instruments ; la
distinction entre formant étroit et formant large ; les résonances de caisse de l'alto et du
violoncelle, que la méthode 8 ne pouvait pas voir ; un chiffre de corroboration par famille, et
la famille où elle échoue.

**Et la table des mélanges**, qui est d'un autre ordre : elle ne décrit plus des instruments mais
des combinaisons, et elle se résume à une règle qu'on peut enseigner en une phrase — *l'écart de
couleur entre deux instruments décide si leur mélange fond, renforce ou crée*.

**Et surtout** : une liste explicite de ce qui n'est pas mesurable. Pour un enseignement sur les
technologies appliquées à la composition, savoir *pourquoi* la question est difficile vaut
probablement autant que la réponse.

## La question qui reste ouverte

La question qui a lancé la reprise — pourquoi telle doublure fusionne et telle autre brille —
**a reçu une réponse** le 27 octobre, et elle tient en une grandeur : l'écart de couleur. Ce n'est
pas la réponse que nous attendions. Nous cherchions des doublures expliquées par un formant
partagé ; ce sont au contraire les instruments **éloignés** qui produisent une couleur nouvelle,
et les proches qui se fondent sans rien créer.

La question qui la remplace est plus étroite et plus urgente : **tout cela est calculé à
l'unisson, alors que l'orchestration double à l'octave.** Un instrument aigu doublant à l'octave
ne recouvre plus les maxima graves de l'autre ; la table en serait sans doute profondément
changée. Le calcul est en place et ne demande qu'un décalage de douze demi-tons.

Et une réserve qui ne se lèvera pas par le calcul : ce qui est mesuré est le mélange **spectral**.
La fusion timbrale en a besoin, mais elle demande aussi la phase, l'espace, et des musiciens qui
s'écoutent. Aucun de ces trois-là n'est dans un spectre moyen.
