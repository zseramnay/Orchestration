---
id: cor-anglais
nom: Cor anglais
nom_en: English horn
famille: bois
sous_famille: anches doubles
ordre: 6
source_echantillons: Yan_Adds-Divers/Winds/EnglishHorn/ordinario
provenance: VSL
technique: ordinario
n_echantillons: 128
tessiture_mesuree: E3–B5
f0_min_hz: 164.8
f0_max_hz: 987.8

classe: A
pertinence_formant: pertinente
rho_hauteur: 0.22

formants:
  - { hz: 2216, statut: "indéterminé", residu_db: 4.36, etendue_residu: [1863, 2635], verdict: retenu, notes: 11 }
  - { hz: 4183, statut: "1 corpus", residu_db: 3.54, etendue_residu: [3517, 5915], verdict: retenu, notes: 20 }

question_ouverte: "zone 1100-1200 Hz : +10 dB sur la moitié grave et les dynamiques douces, absente de l'agrégation globale"

pente_db_octave: -16.5
pente_sd: 2.5
harmoniques_par_note: 27
harmoniques_min_max: [11, 66]
balance_impair_pair_db: -4.3
enveloppe_vraie_hz: 1066
enveloppe_vraie_sd: 242

version_mesure: 2026-10-23-wav-v1
---

# Cor anglais

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **VSL / ConTimbre / perso (Yan_Adds)**.
>
> | Formant | Statut |
> |---|---|
> | **2216 Hz** | **indéterminé** — concordance à un demi-ton avec VN mais aucun couple au même point de grille. Environ une concordance sur quatre à ce seuil est fortuite. |
> | **4183 Hz** | présent dans le seul corpus SOL |
>
> `etendue_residu` dans l'en-tête **n'est pas une largeur de bande** : c'est l'étendue du résidu de la méthode 8, souvent une butée de plage. Les largeurs mesurées sont dans `_LARGEURS-DE-BANDE.md`.

## Méthode retenue, et pourquoi

> **Méthode 8 — deux formants retenus. Le résultat le plus solide des anches doubles
> aiguës.** Voir `00-METHODE.md`.

Le cor anglais réunit les conditions favorables : **ρ = 0,22**, deuxième valeur la plus
basse du corpus après le hautbois, donc classe A nettement établie ; 128 échantillons ;
une médiane de 27 harmoniques par note ; et une pente très régulière (σ = 2,5 dB/octave).

Point remarquable, qui le distingue du hautbois : l'enveloppe vraie note par note donne
**1 066 Hz avec σ = 242 Hz seulement**. C'est, après le basson et le contrebasson, la mesure
note par note la plus stable des bois. Les deux méthodes 6 et 8 sont donc toutes deux
exploitables pour cet instrument, ce qui est rare.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **2 216 Hz** | +4,4 dB | 1 863–2 635 Hz | 11 | **retenu** |
| **4 183 Hz** | +3,5 dB | 3 517–5 915 Hz | 20 | **retenu** |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 2 216 Hz | +6,1 | +4,8 | +6,7 | — | oui |
| 4 183 Hz | — | +3,0 | +5,4 | +4,9 | oui |

Les deux formants passent les deux tests sans réserve. Celui à 2 216 Hz est le mieux
circonscrit des deux : sa bande à −3 dB ne fait que 1 863–2 635 Hz, soit moins d'une demi-
octave, ce qui est la signature d'une résonance bien définie et non d'une tendance large.

## La question ouverte : la zone 1 100–1 200 Hz

Comme pour le hautbois, une zone médiane apparaît en sous-ensembles sans s'imposer
globalement :

| Sous-ensemble | Zone détectée | Résidu |
|---|---|---|
| Moitié grave de la tessiture | 1 174 Hz | **+10,0 dB** |
| Dynamiques douces (p, pp, mp) | 1 174 Hz | **+10,0 dB** |
| Dynamiques fortes (mf, f, ff) | 1 108 Hz | **+10,1 dB** |
| Agrégation globale | — | non détectée |

Un résidu de +10 dB est le plus fort observé pour cet instrument, et il apparaît dans trois
sous-ensembles indépendants, y compris dans les deux moitiés du partage par dynamique. Cela
rend l'hypothèse d'un artefact peu probable. Le fait qu'il ne ressorte pas globalement
s'explique vraisemblablement par la dilution : la zone se situe près de la fondamentale pour
les notes aiguës, où elle n'est donc pas mesurable comme résidu harmonique.

Cette zone correspond exactement à l'enveloppe vraie mesurée (1 066 Hz, σ 242) et au Fp de
l'ancien pipeline (1 135 Hz), et elle est compatible avec Backus (~930 Hz).

**À faire :** mesure par registre. Mon avis provisoire est que ce formant est réel et qu'il
sera confirmé, contrairement au cas du hautbois où le doute est plus sérieux.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −16,5 dB/octave (σ 2,5) | intermédiaire entre hautbois et basson |
| Harmoniques par note | 27 [11–66] | spectre riche |
| Balance impair − pair | −4,3 dB | tuyau conique |
| Enveloppe vraie (méthode 6) | 1 066 Hz, **σ 242** | mesure note par note exploitable |

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Accord |
|---|---|---|---|
| Backus (1969) | ~930 Hz | 1 066 Hz (env. vraie) ; 1 174 Hz (grave) | **plausible** |
| Meyer (2009) | ~600 Hz, /å/ | non retrouvé | désaccord |
| Yan_Adds (ancien) | F1 = 452 ± 163 Hz, /o/ | — | hauteur médiane (rapport 1,12) |
| Ancien Fp | 1 135 Hz | 1 066 Hz (env. vraie) | **convergent** |

Le cas est intéressant : c'est le seul instrument pour lequel l'ancien **Fp** tombe près
d'une valeur que la nouvelle méthode retrouve. Cela ne réhabilite pas le Fp comme mesure de
formant — il reste une mesure de pente bornée — mais indique qu'ici la bande choisie
contenait effectivement une résonance réelle.

Les formants à 2 216 et 4 183 Hz sont en revanche inédits : aucune source consultée ne les
mentionne.

## Registres

Non défini dans `Data/registres.md`. À établir — et nécessaire pour trancher la question
ouverte.

<!-- DEBUT-TRAVAUX-27-10 -->
## Ce que les travaux des 27 et 28/10/2026 disent de cet instrument

### Corroboration par le LPC, note par note

Rang du formant dans son propre spectre : 50 % signifie que les deux méthodes ne s'accordent pas mieux là qu'ailleurs. Voir `../_APPARIEMENT-LPC.md`.

| Formant | Statut | Rang |
|---|---|---|
| 4183 Hz | 1 corpus | 25,0 % |

> **Cet instrument est une anche double, et le LPC y échoue par construction** : il suppose une source filtrée par une cavité séparable, alors que l'anche et la perce sont couplées et que la résonance de l'anche fait partie de la source. Les six formants des anches doubles sont tous sous la médiane. **Ses formants restent établis par la route A — le critère le plus exigeant — mais ne peuvent pas revendiquer de corroboration LPC.**

### Effet de la nuance sur la localisation

Part de pics **non localisables** — l'enveloppe n'y redescend pas de 3 dB des deux côtés, donc le formant n'a pas de largeur mesurable. Voir `../_LARGEUR-PAR-REGISTRE.md`.

| | doux | moyen | fort |
|---|---|---|---|
| **tous registres** | 5 % | 2 % | 5 % |
| grave | 5 % | 5 % | 6 % |
| médium | 5 % | 0 % | 6 % |
| aigu | 7 % | — | — |

**Effet doux → fort : +0 point.** Effet négligeable.

### La carte vocalique corrobore-t-elle les formants ?

**Non.** F1 0,0 % contre 0,0 % au témoin · F2 0,0 % contre 0,0 % au témoin — soit le hasard, ou moins. Les deux méthodes ne mesurent pas la même région du spectre : F1 se situe en médiane à 15,5 demi-tons du formant établi le plus proche, parce que la méthode 8 ne mesure rien sous 3,2 à 5,7 fois f0_min.

Témoin : les propres formants de l'instrument transposés de ±2 à ±6 demi-tons, ce qui conserve leur nombre et leur espacement. Voir `../_COINCIDENCES-F2.md`.

### Dans la table des mélanges

21 paires calculées. **Fusion moyenne : 28 %.** Voir `../_MELANGES.md`.

| Couleur nouvelle avec | | Se fond avec | |
|---|---|---|---|
| Clarinette contrebasse en Sib | 53 % | Flûte basse | 47 % |
| Tuba contrebasse | 47 % | Trompette en Ut | 47 % |
| Alto | 40 % | Cor en Fa | 40 % |

<!-- FIN-TRAVAUX-27-10 -->

## Observations pour l'orchestration

À compléter. Éléments acquis :

- Le formant à 4 183 Hz est **commun au cor anglais et au hautbois** (même valeur mesurée).
  C'est une parenté de famille mesurée, qui pourrait fonder acoustiquement la fusion
  naturelle des anches doubles entre elles — là où l'ancienne étude l'attribuait à une zone
  commune autour de 900–1 200 Hz jamais vérifiée.
- L'ancienne étude annonçait une convergence « quasi-parfaite » cor anglais–clarinette Sib à
  Δ = 11 Hz (452 contre 463 Hz). Les deux valeurs étant des médianes de tessiture dans des
  bins voisins, cette convergence est écartée. Les formants réels des deux instruments
  (2 216 Hz contre 1 567 et 3 517 Hz) ne coïncident pas.
- La convergence cor anglais–basson, également annoncée à Δ = 43 Hz, est à reprendre : les
  formants mesurés sont 2 216 Hz contre 554, 1 108 et 1 974 Hz. Aucune coïncidence directe,
  mais une complémentarité possible, à examiner par la méthode 12.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `cor_anglais_composite.png` | enveloppe composite globale |
| `cor_anglais_registres.png` | enveloppes par registre — pour la zone 1 100–1 200 Hz |
| `anches_doubles_communes.png` (à revoir, non répliqué) | figure de famille : superposition des quatre anches doubles, mettant en évidence 4 183 Hz |

## Journal

- **2026-10-23** — Deux formants retenus sans réserve (2 216 et 4 183 Hz), tous deux inédits.
- **2026-10-23** — Zone 1 100–1 200 Hz détectée à +10 dB dans trois sous-ensembles
  indépendants mais absente du global. Jugée probable, à confirmer par registre.
- **2026-10-23** — Formant à 4 183 Hz identique à celui du hautbois : première parenté de
  famille mesurée dans cette reprise.
- **2026-10-24** — Confrontation au corpus ZSamples : voir la bannière en tête et `_REPLICATION.md`.
