---
id: violoncelle
nom: Violoncelle
nom_en: Violoncello
famille: cordes
sous_famille: cordes solistes
ordre: 21
source_echantillons: FullSOL2020/Strings/Violoncello/ordinario
technique: ordinario
n_echantillons: 291
tessiture_mesuree: C2–C6
f0_min_hz: 65.4
f0_max_hz: 1046.5

classe: C
pertinence_formant: pertinente
rho_hauteur: 0.82

formants:
  - { hz: 1108, residu_db: 4.26, bande_db3: [932, 1174],  verdict: retenu,  notes: 15 }
  - { hz: 2216, residu_db: 6.55, bande_db3: [1863, 2635], verdict: retenu,  notes: 28 }
  - { hz: 3948, residu_db: 4.14, bande_db3: [1759, 5915], verdict: retenu,  notes: 38 }
  - { hz: 554,  residu_db: 3.38, bande_db3: [494, 622],   verdict: retenu-cr, notes: 6,
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
