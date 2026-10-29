---
id: verification-vocalique
titre: La vérification vocalique étendue — F1 réplique, le nom de la voyelle non
date: 2026-10-29
tache: C13
instruments: les 22
corpus: SOL, ZSamples, Xsamples, VO, VN
script: Scripts/v7-formants-2026/verification_vocalique.py
donnees: Etude-Formants/Mesures/Verification-Vocalique/
---

# Étendre la vérification vocalique aux instruments non comparés

## Le trou n'était pas un manque de données

Le plan portait depuis le 24 octobre trois instruments « sans second corpus » : clarinette
basse, flûte contrebasse, clarinette contrebasse. **Les spectres existaient pourtant dans
deux à trois autres corpus chacun.** Et ZSamples — vingt instruments mesurés — n'avait
**aucune carte vocalique du tout**.

La cause n'est pas la donnée mais le nom. Les corpus n'écrivent pas les instruments comme les
fiches : accents absents (« Flute basse », « Trombone tenor », « Petite flute ») et
transposition ajoutée (« Clarinette basse en Sib », « Cor en Fa », « Trompette en Ut »). La
comparaison était une **égalité exacte**. Le script imprimait « absent » et passait.

| | Avant | Après |
|---|---|---|
| Couples instrument × corpus cartographiés, hors SOL | 21 | **56** |
| ZSamples | **0** | 20 |
| Xsamples | 5 | 10 |
| VO | 8 | 15 |
| VN | 8 | 11 |
| **Instruments ayant au moins deux corpus** | 19 | **22 sur 22** |
| dont deux corpus de **provenances distinctes** | — | **21** |

Seul le **tuba contrebasse** reste sans paire valide : ses deux corpus, SOL (`Yan_Adds`) et
VN, sont la même bibliothèque, comme la règle de provenance du 25 octobre l'établit.

> **Une absence de données et une absence de correspondance se ressemblent trop** pour qu'on
> s'en remette à l'œil sur une ligne de sortie. La résolution est désormais dans
> `carte_vocalique.resoudre_fichier`, avec sa garde : le nom suivi de « en » est accepté,
> un préfixe libre non — sans quoi « Flûte » attraperait « Flûte basse » et « Clarinette
> basse » attraperait « Clarinette contrebasse ». Et le script liste maintenant à la fin
> les instruments sans spectres, au lieu de le dire ligne à ligne au milieu du reste.

Vérifié avant de conclure : sur les 21 cartes qui existaient déjà, **les valeurs mesurées
sont identiques au recalcul**. Seul le champ de coïncidence de la tâche C8 bougeait au
centième, parce qu'il partait du Hz arrondi et non du Bark — corrigé.

## La vérification elle-même : un partage net

L'unité est la cellule — un registre à une nuance — ce qui apparie la nuance, comme la tâche
C7 a montré qu'il le fallait chez les cuivres. 176 cellules comparables entre corpus de
provenances distinctes.

| Grandeur | Écart médian entre corpus | 90ᵉ centile |
|---|---|---|
| **F1** | **0,48 Bark** | 1,54 |
| **F2** | **0,92 Bark** | 2,77 |
| distance dans le plan (F1, F2) | 1,16 Bark | 2,93 |

Les seuils de lecture de la carte sont **2,0 Bark** pour qu'une cellule cesse d'être
assignable à une voyelle, et **3,5 Bark** — le seuil de Chistovich — pour qu'elle sorte de
l'espace vocalique.

**F1 réplique : son écart entre corpus est quatre fois plus petit que le seuil qui déciderait
quoi que ce soit.** F2 ne réplique pas aussi bien, et dans un dixième des cellules les deux
corpus placent l'instrument à près de 3 Bark l'un de l'autre.

### Et le nom de la voyelle, qui a besoin des deux, ne réplique pas

| | Cellules |
|---|---|
| comparées, avec voyelle dans au moins deux corpus | 124 |
| **les corpus nomment la même voyelle** | **60 — soit 48 %** |
| désaccord entre voyelles **voisines** (≤ 2 Bark) | 22 |
| **désaccord entre voyelles éloignées** (> 2 Bark) | **42** |

Les deux tiers des désaccords ne sont donc pas des effets de frontière mais des
contradictions :

| Instrument | Cellule | Les corpus disent | Écart |
|---|---|---|---|
| **Violoncelle** | grave, fort | /i/ et /u/ | **6,43 Bark** |
| **Alto** | grave, fort | /e/, /o/, /ø/ | 5,89 Bark |
| **Cor anglais** | grave, fort | /a/, /u/, /ɔ/, /ɛ/ | 5,58 Bark |
| Violoncelle | médium, fort | /u/ et /y/ | 5,36 Bark |
| Clarinette basse | gorge, fort | /e/, /y/, /ɔ/ | 4,72 Bark |

Par instrument, l'accord de voyelle va de **75 %** (clarinette en Mi♭) et 73 % (hautbois) à
**0 %** (flûte contrebasse, trombone basse). Le basson, pourtant l'instrument le mieux
corroboré de l'étude par la tâche C8, n'est qu'à 22 % — et c'est cohérent : son F1 s'accorde
à 1,7 demi-ton entre corpus, son F2 à 2,5.

**Conséquence : un nom de voyelle ne doit pas être publié sur la foi d'un seul corpus.**
C'est la troisième tâche à conclure ainsi, après C6 pour la localisation et C12 pour la
référence vocalique. La règle est générale et devrait figurer dans la méthode.

## Une erreur de tolérance, qu'il faut écrire

La première lecture de ces mêmes données donnait un désastre : écart médian de F1 de
**1,81 demi-ton** entre corpus, et **4 instruments sur 21** seulement sous le demi-ton.

C'était appliquer à la carte la tolérance de la **route A**, qui est celle des formants —
des points d'une grille au douzième d'octave. Une cellule de carte n'est pas cela : c'est une
médiane sur 5 à 49 notes, dont l'**IQR interne vaut 5,1 demi-tons en médiane**. Exiger que
deux telles médianes s'accordent à un demi-ton n'a pas de sens.

| | Valeur médiane |
|---|---|
| dispersion **interne** d'une cellule (IQR de F1) | **5,07 demi-tons** |
| écart **entre corpus** | **1,92 demi-ton** |
| rapport | **0,38** — et sous 1 dans **80 %** des cellules |

**L'écart entre deux corpus vaut 38 % de la dispersion propre d'une cellule.** La carte est
donc reproductible à sa propre échelle ; c'est la tolérance empruntée qui ne l'était pas.

> Le piège a une forme connue dans cette étude : **un test de non-régression doit employer la
> tolérance du résultat qu'il protège.** On l'avait appris à l'envers le 27 octobre, quand un
> test de dégâts au tiers d'octave avait fait recommander une migration que la route A, au
> demi-ton, a démentie. Ici l'erreur est symétrique — une tolérance trop fine au lieu d'une
> tolérance trop lâche — et elle produit un faux négatif au lieu d'un faux positif.

Et un dernier piège, qu'il faut signaler parce qu'il se lit à l'envers : **les cellules
marquées « F1=f0 » sont celles qui répliquent le mieux** (0,96 demi-ton contre 1,61 pour les
cellules sans marque). Ce n'est évidemment pas un gage de qualité — ces cellules mesurent la
note jouée, qui est la même dans tous les corpus par construction. Un taux de réplication
calculé sans écarter ces cellules-là flatterait la méthode exactement là où elle ne mesure
rien.

## Ce que la tâche laisse

- **C13 est close sur la couverture** : 22 instruments sur 22 ont au moins deux corpus,
  21 en ont deux de provenances distinctes. Le tuba contrebasse reste seul, et seule une
  sixième banque y changerait quelque chose.
- Le **contrebasson** gagne deux corpus indépendants pour la carte — Xsamples et ZSamples —
  alors que le chapitre le range parmi les « deux instruments sur lesquels nous ne pouvons
  rien dire ». Cette phrase vise les **formants** et reste vraie pour eux ; elle ne vaut plus
  pour la carte vocalique.
- **F1 est publiable par corpus ; la voyelle ne l'est pas.**

## Ce qui reste ouvert

| Question | Ce qui la trancherait |
|---|---|
| **Pourquoi F2 réplique-t-il deux fois moins bien que F1 ?** | F2 est le second maximum survivant après la fusion de Chistovich : il dépend du contenu aigu, donc de la prise de son. La tâche C6 a déjà montré qu'un corpus (Xsamples) rend les pics 37 % plus étroits — même famille de cause, à vérifier sur les mêmes cellules |
| **Faut-il publier la voyelle en médiane des corpus** plutôt qu'en valeur d'un corpus ? | les données sont là, cellule par cellule, dans `par-cellule.csv`. C'est une décision de présentation, mais elle supprimerait 42 contradictions |
| Les 35 cartes neuves changent-elles un résultat déjà publié ? | **oui, et les trois tâches qui lisent la carte ont été recalculées le jour même.** C8 : conclusion inchangée et renforcée — le basson passe de ×4,7 à ×5,7 sur F1, le reste reste au niveau du hasard. C12 : inchangée, sauf que la dilatation optimale passe à ×1,15 et bat désormais les trois voix. **C11 : trois instruments changent de classe** — le cor quitte « hors espace dans le grave », le trombone ténor passe à « sans second maximum au doux seulement ». Les quatre instruments réellement sous l'espace n'ont pas bougé. Les documents correspondants portent leurs nouveaux chiffres |
| La règle « jamais sur un seul corpus » | elle vaut maintenant pour la localisation (C6), la référence vocalique (C12) et le nom de voyelle (C13). Elle mérite une entrée propre dans `00-METHODE.md` plutôt que trois mentions éparses |

## Journal

- **2026-10-29** — Tâche C13. **Le trou n'était pas un manque de données mais un échec
  d'appariement de noms** : 35 cartes mesurables écartées en silence, dont les vingt de
  ZSamples, parce que les corpus écrivent « Flute basse » et « Clarinette basse en Sib ». Avec
  la résolution tolérante, **22 instruments sur 22 ont au moins deux corpus** et 21 en ont
  deux de provenances distinctes. La vérification elle-même partage net : **F1 réplique**
  (0,48 Bark d'écart médian, contre 2,0 Bark pour changer un verdict), **F2 non** (0,92, et
  2,77 au 90ᵉ centile), et **le nom de voyelle, qui a besoin des deux, ne s'accorde que dans
  48 % des cellules** — avec deux tiers de désaccords entre voyelles distantes de plus de
  2 Bark, jusqu'à /i/ contre /u/ sur le grave fort du violoncelle. **Et une erreur de
  tolérance a dû être corrigée en cours de route** : lue au demi-ton de la route A, la carte
  paraissait ne jamais répliquer ; or la dispersion interne d'une cellule vaut 5,1 demi-tons,
  et l'écart entre corpus 38 % de cela.
