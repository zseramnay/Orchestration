---
id: cor
nom: Cor en Fa
nom_en: Horn in F
famille: cuivres
sous_famille: cuivres principaux
ordre: 13
source_echantillons: FullSOL2020/Brass/Horn/ordinario
provenance: IRCAM
technique: ordinario
n_echantillons: 134
tessiture_mesuree: G1–F5
f0_min_hz: 49.0
f0_max_hz: 698.5

classe: A/B
pertinence_formant: pertinente
rho_hauteur: 0.71

formants:
  - { hz: 739, statut: "1 corpus",  residu_db: 5.36, etendue_residu: [554, 1174],  verdict: retenu, notes: 12 }
  - { hz: 2216, statut: "route A", residu_db: 3.71, etendue_residu: [349, 5270],  verdict: retenu, notes: 32 }
  - { hz: 4183, statut: "1 corpus", residu_db: 3.61, etendue_residu: [349, 5270],  verdict: retenu, notes: 41 }

candidats_rejetes:
  - { hz: 466, residu_db: 2.93, verdict: fragile-cr, ic_bas: [2.4, 2.7],
      raison: "echec de robustesse de fenetre ; reclasse fragile le 23/10",
      commentaire: "zone Backus 400-500, Giesler 250-500, Meyer ~450" }

pente_db_octave: -16.8
pente_sd: 2.9
harmoniques_par_note: 56
harmoniques_min_max: [15, 90]
balance_impair_pair_db: -3.8
enveloppe_vraie_hz: 388
enveloppe_vraie_sd: 131

version_mesure: 2026-10-23-wav-v1
---

# Cor en Fa

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **IRCAM (FullSOL2020)**.
>
> | Formant | Statut |
> |---|---|
> | **739 Hz** | présent dans le seul corpus SOL |
> | **2216 Hz** | **route A** — SOL + VN au même point de grille, 2216 Hz · zone vue par SOL, VN, Xsamples |
> | **4183 Hz** | présent dans le seul corpus SOL |
>
> `etendue_residu` dans l'en-tête **n'est pas une largeur de bande** : c'est l'étendue du résidu de la méthode 8, souvent une butée de plage. Les largeurs mesurées sont dans `_LARGEURS-DE-BANDE.md`.

## Méthode retenue, et pourquoi

> **Méthode 8 — trois formants retenus.** Voir `00-METHODE.md`.

Le cor présente un profil singulier qu'il faut signaler d'emblée : **ρ = 0,71**, une valeur
élevée, qui le classerait parmi les instruments suiveurs. Pourtant, son enveloppe vraie
note par note donne **388 Hz avec σ = 131 Hz seulement** — une stabilité comparable à celle
du basson. Les deux indicateurs semblent se contredire.

L'explication tient à la nature du spectre des cuivres : leur enrichissement harmonique
dépend fortement de la **dynamique** plus que du registre. Le pic de l'enveloppe vraie se
déplace donc avec la force de jeu, et comme les échantillons couvrent toutes les dynamiques
sur toute la tessiture, la corrélation avec f0 capte en partie cet effet. Le ρ élevé n'indique
pas ici un spectre suiveur au sens des flûtes, mais une dépendance à la dynamique.

C'est pourquoi la classe est notée **A/B** et la pertinence formantique maintenue
« pertinente » : les trois formants détectés passent les deux tests, et l'instrument offre
une médiane de **56 harmoniques par note**, ce qui rend l'agrégation très fiable.

Conséquence méthodologique : pour les cuivres, une analyse **par dynamique** est plus
nécessaire qu'une analyse par registre. C'est l'inverse des clarinettes.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **739 Hz** | +5,4 dB | 554–1 174 Hz | 12 | **retenu** |
| **2 216 Hz** | +3,7 dB | large | 32 | **retenu** |
| **4 183 Hz** | +3,6 dB | large | 41 | **retenu** |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 739 Hz | +6,9 | +5,3 | +7,2 | — | oui |
| 2 216 Hz | — | +3,8 | +2,3 | +3,9 | oui |
| 4 183 Hz | — | +6,8 | +4,9 | +4,1 | oui |
| 466 Hz | +4,5 | +2,5 | +3,3 | — | **non** |

Le formant à 739 Hz est le mieux circonscrit : bande à −3 dB de 554 à 1 174 Hz, résidu
supérieur à +5 dB sur trois groupes de rang.

Le rejet de 466 Hz est un cas limite : ses résidus sont positifs et cohérents sur trois
groupes, mais il ne survit pas au changement de fenêtre de redressement. Il se situe par
ailleurs dans la bande du formant à 739 Hz, dont il pourrait être le flanc. **À revoir lors
de l'analyse par dynamique**, car c'est la zone où l'ancien pipeline plaçait le F1 du cor
(388 Hz) et où toutes les sources classiques le situent.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −16,8 dB/octave (σ 2,9) | régulière |
| Harmoniques par note | 56 [15–90] | spectre très riche |
| Balance impair − pair | −3,8 dB | tuyau conique à pavillon |
| **Enveloppe vraie (méthode 6)** | **388 Hz, σ 131** | très stable, malgré un ρ élevé |

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Accord |
|---|---|---|---|
| Backus (1969) | 400–500 Hz | env. vraie 388 Hz ; candidat 466 rejeté | **plausible** |
| Giesler (1985) | 250–500 Hz, *o-ähnlich* | env. vraie 388 Hz | **bon** |
| Meyer (2009) | ~450 Hz, /o/ | env. vraie 388 Hz ; candidat 466 rejeté | **plausible** |
| Ancien pipeline `specenv` | F1 = 388 Hz | env. vraie 388 Hz | **identique** |
| Ancien Fp | 738 Hz | **formant retenu à 739 Hz** | **identique** |

Ce cas est le plus favorable de tout le corpus aux mesures antérieures, et il faut le dire :
l'ancien F1 (388 Hz) coïncide exactement avec notre enveloppe vraie, et l'ancien Fp (738 Hz)
coïncide exactement avec notre premier formant retenu (739 Hz).

L'explication est cohérente avec le diagnostic général. D'une part, la tessiture du cor
descend à 49 Hz et sa majeure partie reste sous la limite de validité du lissage Orchidea
(276 Hz), contrairement aux instruments aigus. D'autre part, le rapport entre l'ancien F1 et
le milieu géométrique de la tessiture valait **1,87** pour le cor, l'un des plus élevés du
corpus : c'était le signe que la valeur mesurait une résonance réelle et non la hauteur
médiane. Les anciens chiffres du cor étaient donc justes — par chance de configuration, non
par validité de méthode.

Les formants à 2 216 et 4 183 Hz sont en revanche inédits.

## Registres

`Data/registres.md` : sons pédale F1–A1 · grave A#1–B2 · médium C3–E4 · aigu F4–F5 ·
suraigu à partir de F#5.

Mesure par registre : non effectuée. **Priorité moindre que l'analyse par dynamique** pour
cet instrument (voir ci-dessus).

<!-- DEBUT-TRAVAUX-27-10 -->
## Ce que les travaux des 27 et 28/10/2026 disent de cet instrument

### Corroboration par le LPC, note par note

Rang du formant dans son propre spectre : 50 % signifie que les deux méthodes ne s'accordent pas mieux là qu'ailleurs. Voir `../_APPARIEMENT-LPC.md`.

| Formant | Statut | Rang |
|---|---|---|
| 739 Hz | 1 corpus | 84,6 % |
| 2216 Hz | route A | 30,8 % |
| 4183 Hz | 1 corpus | 0,0 % |

### Effet de la nuance sur la localisation

Part de pics **non localisables** — l'enveloppe n'y redescend pas de 3 dB des deux côtés, donc le formant n'a pas de largeur mesurable. Voir `../_LARGEUR-PAR-REGISTRE.md`.

| | doux | moyen | fort |
|---|---|---|---|
| **tous registres** | 28 % | 31 % | 36 % |
| sons pédale | — | 24 % | — |
| grave | 33 % | 32 % | 47 % |
| médium | 16 % | 38 % | 40 % |
| aigu | 40 % | 25 % | 15 % |

**Effet doux → fort : -8 points.** Elle se dégrade en jeu fort — comportement inverse de celui des trombones et de la trompette.

### La couleur dépend de la nuance, et ne se lit que jusqu'au médium

F1 du motif d'excitation, notes **pédale + grave + médium** réunies : au-dessus du médium la mesure donne la fondamentale et non une couleur (F1/f0 = 1,0, pente contre f0 de 1,03 au suraigu). Voir `../_REGISTRES-PEDALE.md`.

| Corpus | doux | moyen | fort | doux → fort |
|---|---|---|---|---|
| SOL | 291 Hz | 351 Hz | 439 Hz | **+7,1 dt** |
| Xsamples | — | 434 Hz | 508 Hz | **—** |
| VN | 402 Hz | 553 Hz | 592 Hz | **+6,7 dt** |

**Aucune paire de corpus ne s'accorde à moins d'un demi-ton**, à aucune nuance. La couleur de cet instrument n'est pas répliquée.

### Dans la table des mélanges

21 paires calculées. **Fusion moyenne : 17 %.** Voir `../_MELANGES.md`.

| Couleur nouvelle avec | | Se fond avec | |
|---|---|---|---|
| Trompette en Ut | 54 % | Contrebasson | 71 % |
| Hautbois | 40 % | Tuba contrebasse | 67 % |
| Cor anglais | 37 % | Clarinette en Sib | 61 % |

<!-- FIN-TRAVAUX-27-10 -->

## Observations pour l'orchestration

À compléter. Éléments acquis :

- Le formant à **4 183 Hz est commun au cor, au hautbois et au cor anglais** — trois
  instruments de deux familles différentes, chacun validé indépendamment par les deux tests.
  C'est la convergence inter-familles la mieux établie de cette reprise, et elle pourrait
  fonder acoustiquement la fusion cor–bois, bien connue des orchestrateurs.
- L'ancienne étude annonçait une convergence « quasi-parfaite » cor–alto à Δ = 11 Hz
  (388 contre 377 Hz) et un « super-cluster à 323 Hz » entre cor, trompette et alto dans le
  registre médium. Ces valeurs tombaient dans des bins voisins ; la convergence cor–alto
  reste plausible puisque le F1 du cor était ici une mesure valide, mais elle doit être
  reprise avec les cordes mesurées par la même méthode.
- Le formant à 739 Hz, confirmé par deux voies, est un bon candidat pour expliquer le
  caractère « rond et chaleureux » : il se situe dans la zone /o/–/a/ et la bande est large.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `cor_composite.png` | enveloppe composite avec les trois formants |
| `cor_dynamiques.png` | **prioritaire** : enveloppes composites par dynamique (p / mf / ff) |
| `cuivres_rho.png` | figure de famille : ρ des six cuivres, montrant l'effet dynamique |
| `convergence_4183.png` (ÉCARTÉE 24/10) | figure inter-familles : cor, hautbois, cor anglais à 4 183 Hz |

## Révision du 23/10/2026 — critère assoupli

La zone à **466 Hz** passe de *rejetée* à **fragile-cr**, et c'est important pour cet
instrument. Le bootstrap donne +3,0 dB [IC +2,4] en dessous du rang 5 et +3,2 dB [IC +2,7]
au-dessus : le test d'indépendance au rang est **franchement passé des deux côtés**. Seule la
robustesse de fenêtre échoue encore.

Or c'est précisément la zone où **les trois sources classiques situent le formant du cor** —
Backus 400–500 Hz, Giesler 250–500 Hz, Meyer ~450 Hz — et où l'ancien pipeline donnait
F1 = 388 Hz, valeur que notre enveloppe vraie confirme. Le rejet initial était donc trop
sévère.

Ce fichier signalait déjà ce candidat comme « cas limite » à revoir lors de l'analyse par
dynamique. Cette analyse reste à faire, et elle devrait trancher : si la résonance est
réelle mais sa position dépend de la force de jeu, la bande globale est nécessairement
étalée, ce qui expliquerait l'échec de robustesse sans infirmer le formant.

## Journal

- **2026-10-23** — Trois formants retenus (739, 2 216, 4 183 Hz).
- **2026-10-23** — Contradiction apparente entre ρ = 0,71 et σ(enveloppe vraie) = 131 Hz
  résolue : pour les cuivres, la dépendance est à la dynamique plus qu'au registre. Analyse
  par dynamique à mener en priorité pour toute la famille.
- **2026-10-23** — Seul instrument du corpus dont les deux valeurs antérieures (F1 = 388 et
  Fp = 738) sont confirmées. Configuration favorable, non validité de la méthode ancienne.
- **2026-10-23 (soir)** — Critère assoupli appliqué (voir `00-METHODE.md`).
- **2026-10-24** — Confrontation au corpus ZSamples : voir la bannière en tête et `_REPLICATION.md`.
