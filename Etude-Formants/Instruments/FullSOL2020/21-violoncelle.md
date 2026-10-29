---
id: violoncelle
nom: Violoncelle
nom_en: Violoncello
famille: cordes
sous_famille: cordes solistes
ordre: 21
source_echantillons: FullSOL2020/Strings/Violoncello/ordinario
provenance: IRCAM
technique: ordinario
n_echantillons: 291
tessiture_mesuree: C2–C6
f0_min_hz: 65.4
f0_max_hz: 1046.5

classe: C
pertinence_formant: pertinente
rho_hauteur: 0.82

formants:
  - { hz: 1108, statut: "route A", residu_db: 4.26, etendue_residu: [932, 1174],  verdict: retenu,  notes: 15 }
  - { hz: 2216, statut: "1 corpus", residu_db: 6.55, etendue_residu: [1863, 2635], verdict: retenu,  notes: 28 }
  - { hz: 3948, statut: "route A", residu_db: 4.14, etendue_residu: [1759, 5915], verdict: retenu,  notes: 38 }
  - { hz: 554, statut: "1 corpus",  residu_db: 3.38, etendue_residu: [494, 622],   verdict: retenu-cr, notes: 6,
      ic_bas: [2.0, 3.9], commentaire: "CONFIRME le 23/10 ; coincide exactement avec le basson" }

pente_db_octave: -13.3
pente_sd: 1.4
harmoniques_par_note: 47
harmoniques_min_max: [10, 90]
balance_impair_pair_db: -2.6
enveloppe_vraie_hz: 312
enveloppe_vraie_sd: 554

version_mesure: 2026-10-23-wav-v1
---

# Violoncelle

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **IRCAM (FullSOL2020)**.
>
> | Formant | Statut |
> |---|---|
> | **554 Hz** | présent dans le seul corpus SOL |
> | **1108 Hz** | **route A** — SOL + VN au même point de grille, 1108 Hz · zone vue par SOL, VN · pic localisé, Q = 6.4 |
> | **2216 Hz** | présent dans le seul corpus SOL |
> | **3948 Hz** | **route A** — SOL + VN au même point de grille, 3948 Hz · zone vue par SOL, VN · pic localisé, Q = 18.3 |
>
> `etendue_residu` dans l'en-tête **n'est pas une largeur de bande** : c'est l'étendue du résidu de la méthode 8, souvent une butée de plage. Les largeurs mesurées sont dans `_LARGEURS-DE-BANDE.md`.

## Classe C — trois formants retenus, bien circonscrits

> **Méthode 8 — trois formants retenus, un fragile.** Voir `00-METHODE.md`, section
> « Typologie — trois classes ».

Classe C, ρ = 0,82, enveloppe vraie dispersée (312 Hz, σ = 554 Hz). Mais trois formants
passent les deux tests, et deux d'entre eux sont **remarquablement bien circonscrits** :
932–1 174 Hz pour le premier, 1 863–2 635 Hz pour le second, soit moins d'une demi-octave
chacun. Pour une corde, c'est la signature nette de résonances de caisse.

Conditions de mesure favorables : 291 échantillons, une médiane de **47 harmoniques par
note**, et la **pente la plus régulière de tout le corpus** — σ = 1,4 dB/octave.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **1 108 Hz** | +4,3 dB | 932–1 174 Hz | 15 | **retenu** |
| **2 216 Hz** | +6,6 dB | 1 863–2 635 Hz | 28 | **retenu** |
| **3 948 Hz** | +4,1 dB | 1 759–5 915 Hz | 38 | **retenu** |
| 554 Hz | +3,4 dB | 494–622 Hz | 6 | fragile |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 1 108 Hz | +0,9 | +4,6 | +4,4 | +6,4 | oui |
| 2 216 Hz | — | +4,1 | +6,0 | +4,9 | oui |
| 3 948 Hz | — | +1,2 | +4,2 | +6,6 | oui |
| 554 Hz | +4,3 | **+7,3** | +9,4 (2 notes) | — | oui |

Le cas de **554 Hz** mérite attention. Il passe le test de robustesse de fenêtre et ses
résidus sont forts et croissants (+4,3 puis +7,3 puis +9,4 dB), mais il n'est alimenté que
par 6 notes distinctes et son groupe k = 8–15 ne compte que 2 notes. Le verdict « fragile »
résulte donc d'un manque de couverture, non d'une contradiction.

Or cette zone est précisément celle où l'on attendrait la résonance principale de caisse du
violoncelle, et elle coïncide **exactement** avec le premier formant du basson (554 Hz). Sa
bande à −3 dB est l'une des plus étroites du corpus (494–622 Hz). **Mon avis est qu'il est
réel et qu'il sera confirmé** : à vérifier par une mesure restreinte au registre grave, où
les notes alimentant cette bande par leurs harmoniques 2 à 7 sont les plus nombreuses.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| **Pente spectrale médiane** | −13,3 dB/octave (**σ 1,4**) | **la plus régulière du corpus** |
| Harmoniques par note | 47 [10–90] | spectre riche |
| Balance impair − pair | −2,6 dB | corde frottée |
| Enveloppe vraie (méthode 6) | 312 Hz, σ 554 | inexploitable note par note |

La régularité de la pente (σ = 1,4) est le record du corpus, devant l'alto (1,8) et le
basson (1,7). Elle signifie que le comportement spectral du violoncelle varie très peu d'une
note et d'une dynamique à l'autre, ce qui est cohérent avec sa réputation d'homogénéité sur
toute la tessiture.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Lecture |
|---|---|---|---|
| Backus, Giesler, Meyer | pas de valeur propre au violoncelle dans les tableaux de l'étude | — | à chercher aux sources primaires |
| Ancien pipeline `specenv` | F1 = 205 Hz, zone /u/ | — | zone de la fondamentale (rapport 0,78) |
| Cepstre mai 2026 (ordre 80) | 1 991 Hz, σ 915 | 2 216 Hz | proche |
| Cepstre mai 2026, **ensemble** | 355 Hz, σ **16** | — | voir ci-dessous |

Le même phénomène que pour l'alto se retrouve, et de façon encore plus spectaculaire : le
cepstre de mai 2026 donnait **1 991 Hz avec σ = 915 Hz pour le violoncelle soliste** mais
**355 Hz avec σ = 16 Hz pour l'ensemble de violoncelles**. Un même corps de résonance ne peut
pas avoir un formant qui saute de deux octaves et demie entre soliste et pupitre, ni une
dispersion divisée par 57.

J'avais relevé cette incohérence comme un défaut de l'estimateur. Elle peut aussi être un
résultat : dans un pupitre, les légers désaccords entre instrumentistes densifient le spectre
et révèlent le filtre, là où un soliste ne l'éclaire qu'aux positions de ses harmoniques. À
tester en priorité, car si l'hypothèse se vérifie, **les ensembles sont de meilleurs témoins
formantiques que les solistes** et il faudra les privilégier.

## Registres

`Data/registres.md` : grave C2–G2 · médium G2–G3 · aigu G3–G4 · suraigu G4 et au-dessus.

Mesure par registre : non effectuée. **Prioritaire pour cet instrument**, afin de confirmer
la zone 554 Hz.

## Résonance de caisse vers 215–220 Hz — méthode 12, quatre corpus

**Acquis le 27/10/2026, tâche C1.** Dans le registre médium (Sol2–Sol3, f0 de 104 à 196 Hz),
le premier maximum du motif d'excitation auditive se tient entre **209 et 223 Hz sur quatre
corpus de provenances distinctes** — 1,1 demi-ton d'écart — avec une pente de log2(F1) contre
log2(f0) de **−0,04** sur onze demi-tons de fondamentale. **F1 ne bouge pas pendant que la note
parcourt presque une octave** : c'est une résonance, pas un harmonique.

| Corpus | n | F1 médian | Pente | F1 − centroïde |
|---|---|---|---|---|
| SOL | 90 | 223 Hz | +0,50 | −2,5 dt |
| ZSamples | 72 | 209 Hz | 0,00 | −3,4 dt |
| VO | 30 | 223 Hz | −0,38 | −4,1 dt |
| VN | 36 | 217 Hz | −0,08 | −4,0 dt |

Le rang harmonique médian passe de 2,1 en bas du registre à 1,1 en haut : la même résonance est
excitée par l'harmonique 2 puis par l'harmonique 1. La valeur tombe dans la zone B1 que la
littérature acoustique donne au violoncelle (180–250 Hz).

**Aucun des quatre formants retenus par la méthode 8 n'est dans cette zone.** C'est le second
cas, après l'alto, où la méthode 12 voit une résonance que la méthode 8 manque par
construction : la zone de caisse est large et plate, donc redressée avec le reste par
`residu_local()`.

**Réserve à publier avec le résultat.** 220 Hz est aussi la corde de La à vide, et les luthiers
accordent volontiers le mode B1+ au voisinage de cette corde. La mesure établit qu'il y a là une
résonance qui ne suit pas la note ; elle **ne sépare pas** le mode de caisse de la résonance
sympathique de la corde à vide. Pour l'orchestrateur la conséquence est la même ; l'attribution
physique reste ouverte. Voir `../_RESONANCE-DE-CAISSE.md`.

Le registre **grave** (Do2–Sol2), lui, ne donne rien : pentes de −1,33 à +0,36 selon le corpus,
médianes de 178 à 238 Hz sur 5 demi-tons d'écart.

<!-- DEBUT-TRAVAUX-27-10 -->
## Ce que les travaux des 27 et 28/10/2026 disent de cet instrument

### Corroboration par le LPC, note par note

Rang du formant dans son propre spectre : 50 % signifie que les deux méthodes ne s'accordent pas mieux là qu'ailleurs. Voir `../_APPARIEMENT-LPC.md`.

| Formant | Statut | Rang |
|---|---|---|
| 554 Hz | 1 corpus | 91,2 % |
| 1108 Hz | route A | 67,6 % |
| 2216 Hz | 1 corpus | 88,2 % |
| 2631 Hz | route A | 91,2 % |
| 3948 Hz | route A | 91,2 % |

### Effet de la nuance sur la localisation

Part de pics **non localisables** — l'enveloppe n'y redescend pas de 3 dB des deux côtés, donc le formant n'a pas de largeur mesurable. Voir `../_LARGEUR-PAR-REGISTRE.md`.

| | doux | moyen | fort |
|---|---|---|---|
| **tous registres** | 8 % | 7 % | 10 % |
| grave | 6 % | 4 % | 11 % |
| médium | 7 % | 7 % | 12 % |
| aigu | 11 % | 8 % | 8 % |
| suraigu | 0 % | 0 % | — |

**Effet doux → fort : -2 points.** Effet négligeable.

### La carte vocalique corrobore-t-elle les formants ?

**Non.** F1 0,0 % contre 1,7 % au témoin · F2 10,0 % contre 9,5 % au témoin — soit le hasard, ou moins. Les deux méthodes ne mesurent pas la même région du spectre : F1 se situe en médiane à 15,5 demi-tons du formant établi le plus proche, parce que la méthode 8 ne mesure rien sous 3,2 à 5,7 fois f0_min.

Témoin : les propres formants de l'instrument transposés de ±2 à ±6 demi-tons, ce qui conserve leur nombre et leur espacement. Voir `../_COINCIDENCES-F2.md`.

### Dans la table des mélanges

21 paires calculées. **Fusion moyenne : 23 %.** Voir `../_MELANGES.md`.

| Couleur nouvelle avec | | Se fond avec | |
|---|---|---|---|
| Trombone basse | 52 % | Tuba basse | 52 % |
| Basson | 42 % | Tuba contrebasse | 51 % |
| Hautbois | 41 % | Flûte basse | 49 % |

<!-- FIN-TRAVAUX-27-10 -->

## Observations pour l'orchestration

À compléter. Éléments acquis :

- La zone **554 Hz (fragile) coïncide exactement avec le premier formant du basson**
  (554 Hz, retenu). Si elle se confirme, ce serait la convergence la plus précise de toute
  la reprise, et elle rendrait compte acoustiquement d'une parenté timbrale
  basson–violoncelle que les orchestrateurs connaissent bien. **À confirmer en priorité.**
- La zone **2 216 Hz est commune au violoncelle, au cor (2 216 Hz), au cor anglais
  (2 216 Hz) et au trombone ténor (2 091 Hz)** — quatre instruments de trois familles, tous
  validés indépendamment.
- L'ancienne étude proposait des doublures à l'octave avec le cor anglais (Δ = 247 Hz) et le
  cor (Δ = 183 Hz), décrites comme « complémentarité /o/–/u/ ». Ces écarts portaient sur des
  zones de fondamentale. Elles sont à reprendre, mais la correspondance réelle à 2 216 Hz
  pourrait les justifier autrement — et mieux.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `violoncelle_composite.png` | enveloppe composite — formants bien circonscrits |
| `violoncelle_554.png` | zone 554 Hz : résidus par rang et couverture, pour trancher le cas fragile |
| `violoncelle_soliste_vs_ensemble.png` | **piste méthodologique** : 1 991 Hz σ 915 contre 355 Hz σ 16 |
| `basson_violoncelle_554.png` | la coïncidence à 554 Hz, si confirmée |

## Révision du 23/10/2026 — critère assoupli

**La zone à 554 Hz est confirmée** : elle passe de *fragile* à **retenu-cr**. Le bootstrap
par note donne +2,4 dB [IC +2,0] en dessous du rang 4 et +4,7 dB [IC +3,9] au-dessus, sur 70
et 50 points. Les deux bornes basses sont nettement positives.

Ce fichier avançait que ce candidat « est réel et sera confirmé ». Il l'est. Son statut
fragile résultait bien d'un manque de couverture dans les groupes de rang étroits, non d'une
contradiction.

Conséquence : **la coïncidence exacte avec le premier formant du basson (554 Hz, retenu)
devient un résultat et non une hypothèse.** Deux instruments de familles différentes, chacun
validé indépendamment, partagent une résonance à la même fréquence au Hz près. C'est la
convergence la plus précise de toute la reprise, et elle rend compte acoustiquement d'une
parenté timbrale basson–violoncelle que les orchestrateurs connaissent de longue date. À
examiner en priorité par la méthode 12.

Les trois autres formants (1 108, 2 216, 3 948 Hz) sont tous confirmés.

## Journal

- **2026-10-23** — Trois formants retenus (1 108, 2 216, 3 948 Hz), bien circonscrits.
- **2026-10-23** — Pente la plus régulière du corpus (σ 1,4 dB/octave).
- **2026-10-23** — Zone 554 Hz fragile par manque de couverture, mais coïncidant exactement
  avec le premier formant du basson. Jugée probable, à confirmer par registre.
- **2026-10-23** — L'incohérence soliste/ensemble des données de mai 2026 (1 991 Hz σ 915
  contre 355 Hz σ 16) est reclassée : d'un défaut d'estimateur en hypothèse à tester.
- **2026-10-23 (soir)** — Critère assoupli appliqué (voir `00-METHODE.md`).
- **2026-10-27** — Tâche C1 : résonance de caisse vers 215–220 Hz établie au médium sur quatre
  corpus par la méthode 12, invisible à la méthode 8. Réserve sur la corde de La à vide.
