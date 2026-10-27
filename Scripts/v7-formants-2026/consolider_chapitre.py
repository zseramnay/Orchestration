# -*- coding: utf-8 -*-
"""Intègre au chapitre de synthèse les six résultats du 27/10/2026."""
import io, sys

P = '_CHAPITRE-SYNTHESE.md'
t = io.open(P, encoding='utf-8').read()
n = 0

def rep(a, b, marq):
    global t, n
    if marq in t:
        print('  = déjà :', marq[:46]); return
    if t.count(a) != 1:
        print('  !! ancre en %d exemplaires : %s' % (t.count(a), a[:60])); sys.exit(1)
    t = t.replace(a, b, 1); n += 1
    print('  +', marq[:60])

# ── A. le préambule affirmait que les mélanges ne sont pas mesurés ──────────
rep("""Elle ne mesure **pas** les mélanges. Toutes les mesures portent sur des notes jouées seules,
puis agrégées. Lorsque deux instruments partagent un formant, on peut en déduire une
**prédiction** sur ce que leur superposition devrait produire — on ne peut pas l'affirmer.""",
"""Elle mesure aussi les **mélanges**, et c'est neuf depuis le 27 octobre. Non pas en simulant :
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
sous leur propre fondamentale.""",
'préambule : les mélanges sont mesurés')

rep("""Ce qui est écrit ci-dessous comme convergence est donc **un fait de mesure**. Ce qui en est
tiré pour l'orchestration est **une hypothèse**, que la méthode 12 appliquée à des mélanges
réels pourra confirmer ou défaire.""",
"""Ce qui est écrit ci-dessous comme convergence reste **un fait de mesure**, et ce qui en est tiré
pour l'orchestration une **hypothèse**. La partie III, sur les mélanges, est la seule à porter
sur des combinaisons ; elle ne valide pas les convergences, elle répond à une autre question.""",
'préambule : la partie III porte sur les combinaisons')

# ── B. la dynamique : ajouter l'effet sur la localisation ──────────────────
rep("""Conséquence pour la méthode : toute mesure de cuivre agrégée sur plusieurs nuances est
nécessairement étalée. C'est ce qui explique que le cor à 466 Hz et la trompette à 1 479 Hz
soient longtemps restés « fragiles » alors qu'ils correspondent à la littérature.""",
"""Conséquence pour la méthode : toute mesure de cuivre agrégée sur plusieurs nuances est
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

Détail dans `_LARGEUR-PAR-REGISTRE.md`.""",
'la nuance décide s\'il y a un formant')

# ── C. deux sections neuves à la fin de la partie II ───────────────────────
rep("""---

# III. Ce qui n'est pas mesurable, et c'est un résultat""",
"""## 6. Les cordes ont une résonance de caisse, et la méthode 8 ne la voit pas

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

# IV. Ce qui n'est pas mesurable, et c'est un résultat""",
'partie III : les mélanges')

# ── D. renuméroter la partie IV ────────────────────────────────────────────
rep("""# IV. Ce que cela change pour l'enseignement""",
    """# V. Ce que cela change pour l'enseignement""",
    '# V. Ce que cela change')

rep("""**Ce qu'on gagne** : trois unissons formantiques et cinq régions partagées, chacun validé
indépendamment sur des corpus de provenances distinctes ; une typologie qui dit, pour chaque
instrument, si la notion de formant s'y applique ; un effet de la dynamique chiffré ; une
signature de la géométrie du tuyau ; une carte vocalique qui couvre les vingt-deux instruments ;
et la distinction entre formant étroit et formant large.""",
"""**Ce qu'on gagne** : trois unissons formantiques et cinq régions partagées, chacun validé
indépendamment sur des corpus de provenances distinctes ; une typologie qui dit, pour chaque
instrument, si la notion de formant s'y applique ; un effet de la dynamique chiffré — en position
*et* en localisation ; une signature de la géométrie du tuyau, qui se retrouve jusque dans l'effet
de la nuance chez les cuivres ; une carte vocalique qui couvre les vingt-deux instruments ; la
distinction entre formant étroit et formant large ; les résonances de caisse de l'alto et du
violoncelle, que la méthode 8 ne pouvait pas voir ; un chiffre de corroboration par famille, et
la famille où elle échoue.

**Et la table des mélanges**, qui est d'un autre ordre : elle ne décrit plus des instruments mais
des combinaisons, et elle se résume à une règle qu'on peut enseigner en une phrase — *l'écart de
couleur entre deux instruments décide si leur mélange fond, renforce ou crée*.""",
'ce qu\'on gagne : mis à jour')

# ── E. la question ouverte a changé ────────────────────────────────────────
rep("""## La question qui reste ouverte

Tout ce chapitre décrit des instruments **seuls**. La question qui a lancé la reprise —
pourquoi telle doublure fusionne et telle autre brille — demande de mesurer des **mélanges**.
C'est la seule tâche du programme qui touche directement à l'orchestration, et elle n'est pas
faite.

Les convergences ci-dessus en sont les hypothèses de départ. Elles ne sont pas la réponse.""",
"""## La question qui reste ouverte

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
s'écoutent. Aucun de ces trois-là n'est dans un spectre moyen.""",
'La question qui la remplace')

t = t.replace('date: 2026-10-26', 'date: 2026-10-27\nrevision: >\n'
              '  27/10 — intégration des six travaux du bloc C : résonances de caisse,\n'
              '  plancher du domaine, plancher de la fondamentale, corroboration LPC,\n'
              '  effet de la nuance sur la localisation, et la partie III sur les mélanges.', 1)
io.open(P, 'w', encoding='utf-8').write(t)
print('\n%d sections intégrées · %d lignes' % (n, t.count('\n') + 1))
