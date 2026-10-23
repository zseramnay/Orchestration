---
id: clarinette-basse
nom: Clarinette basse en Sib
nom_en: Bass clarinet
famille: bois
sous_famille: clarinettes
ordre: 11
source_echantillons: Yan_Adds-Divers/Winds/Bass-Clarinet-Bb/ordinario
technique: ordinario
n_echantillons: 173
tessiture_mesuree: A#1–E6
f0_min_hz: 58.3
f0_max_hz: 1318.5

classe: B
pertinence_formant: partielle
rho_hauteur: 0.88

formants:
  - { hz: 1244, residu_db: 3.15, bande_db3: [783, 5915],  verdict: retenu, notes: 19 }
  - { hz: 2487, residu_db: 5.49, bande_db3: [1759, 3726], verdict: retenu, notes: 31 }
  - { hz: 5270, residu_db: 2.13, bande_db3: [415, 5915],  verdict: fragile-cr, notes: 42,
      ic_bas: [4.2, -0.6], commentaire: "declasse le 23/10 : moitie de rang eleve a +0,0 dB, IC negatif" }

candidats_rejetes:
  - { hz: 466, residu_db: 5.82, verdict: rejete, raison: "+11,3 dB aux rangs 2-3 mais 0,0 dB aux rangs 4-7 — effet de hauteur" }

pente_db_octave: -16.5
pente_sd: 5.7
harmoniques_par_note: 44
harmoniques_min_max: [8, 90]
balance_impair_pair_db: 4.5
enveloppe_vraie_hz: 323
enveloppe_vraie_sd: 504

version_mesure: 2026-10-23-wav-v1
---

# Clarinette basse en Sib

## Méthode retenue, et pourquoi

> **Méthode 8 — trois formants retenus. Corpus le plus fourni de la famille.**
> Voir `00-METHODE.md`.

Avec **173 échantillons**, c'est le corpus le plus riche des douze bois, et avec une médiane
de 44 harmoniques par note, l'agrégation est bien alimentée. L'instrument reste de classe B
— ρ = 0,88, le plus élevé des clarinettes — mais la combinaison d'un grand corpus et d'une
tessiture descendant à 58 Hz permet de détecter des résonances réelles. D'où la mention
« partielle ».

Cet instrument fournit en outre l'illustration la plus nette du fonctionnement du filtre de
rejet, avec le cas du candidat à 466 Hz détaillé plus bas.

## Formants

| Zone | Résidu | Largeur −3 dB | Notes distinctes | Verdict |
|---|---|---|---|---|
| **1 244 Hz** | +3,2 dB | 783–5 915 Hz | 19 | **retenu** |
| **2 487 Hz** | +5,5 dB | 1 759–3 726 Hz | 31 | **retenu** |
| **5 270 Hz** | +2,1 dB | très large | 42 | **retenu** (faible) |

## Tests de rejet

| Zone | k = 2–3 | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse |
|---|---|---|---|---|---|
| 1 244 Hz | — | +6,3 | +7,7 | +2,7 | oui |
| 2 487 Hz | +8,5 | +10,0 | +5,4 | +4,4 | oui |
| 5 270 Hz | — | +6,8 | +4,1 | +1,8 | oui |
| 466 Hz | **+11,3** | **0,0** | — | — | **non** |

Le cas de **466 Hz est exemplaire** et mérite d'être conservé comme illustration de la
méthode. Son résidu atteint +11,3 dB lorsque la bande est éclairée par les harmoniques 2 et
3, et tombe exactement à **0,0 dB** lorsqu'elle est éclairée par les harmoniques 4 à 7.
Un formant est une propriété du filtre : il ne peut pas dépendre du rang harmonique qui le
révèle. Ce pic était donc un effet de hauteur — les notes dont l'harmonique 2 tombe vers
466 Hz, soit le registre autour de D3–E3 — et non une résonance.

C'est précisément ce type de pic que l'ancienne méthode retenait sans contrôle. Le F1 publié
pour cet instrument était d'ailleurs 323 Hz, dans la même région.

Le formant à 2 487 Hz est en revanche le mieux établi : positif sur les quatre groupes de
rang, avec un maximum de +10,0 dB, et 31 notes distinctes.

## Descripteurs globaux

| Descripteur | Valeur | Lecture |
|---|---|---|
| Pente spectrale médiane | −16,5 dB/octave (σ 5,7) | forte dispersion, comme toute la famille |
| Harmoniques par note | 44 [8–90] | spectre riche |
| **Balance impair − pair** | **+4,5 dB** | dominance impaire marquée |
| Enveloppe vraie (méthode 6) | 323 Hz, σ 504 | σ supérieur à la valeur : inexploitable note par note |

L'écart-type de l'enveloppe vraie (504 Hz) dépasse la valeur médiane elle-même (323 Hz).
C'est la démonstration la plus brutale du caractère suiveur : note par note, la mesure n'a
aucun sens pour cet instrument. Seule l'agrégation en a.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Lecture |
|---|---|---|---|
| Yan_Adds (ancien) | F1 = 323 ± 87 Hz, /u/ | 466 Hz **rejeté** | effet de hauteur, écarté |
| Ancien Fp | 1 204 Hz | 1 244 Hz | **convergent** |
| Cepstre mai 2026 (ordre 80) | 1 195 Hz, 24 % sur un harmonique | 1 244 Hz | proche |
| Backus, Giesler, Meyer | aucune donnée | — | instrument non étudié |

Le Fp de 1 204 Hz et le cepstre de 1 195 Hz encadrent notre formant retenu à 1 244 Hz. Trois
méthodes différentes convergent donc sur cette zone, ce qui la rend solide malgré un résidu
modeste de +3,2 dB.

## Registres

`Data/registres.md` : grave A#1–D3 · médium D#3–G#3 · aigu A3–A#4 · suraigu à partir de B4.

Mesure par registre : **non effectuée.** Elle est d'autant plus souhaitable que le candidat
rejeté à 466 Hz correspond au registre médium : une analyse par registre dirait si une
résonance y existe réellement, ce que l'agrégation globale ne peut pas trancher.

## Observations pour l'orchestration

À compléter. Éléments acquis :

- Le formant à 2 487 Hz, bien établi, n'a pas d'équivalent connu dans la littérature pour
  cet instrument.
- Les doublures annoncées antérieurement — trombone basse à Δ = 65 Hz, contrebasson à
  Δ = 97 Hz, flûte basse à Δ = 22 Hz décrite comme « excellente » — reposaient toutes sur des
  zones de fondamentale et sont à reprendre intégralement.
- La dominance impaire de +4,5 dB en fait, avec la clarinette contrebasse, le meilleur
  candidat pour la doublure de comblement par un instrument à spectre complet (flûte,
  basson) : le partiel pair manquant y est le plus manquant.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `clarinette_basse_composite.png` | enveloppe composite avec les trois formants |
| `clarinette_basse_466.png` | **figure méthodologique** : résidu à 466 Hz par groupe de rang, +11,3 dB puis 0,0 dB |
| `clarinette_basse_registres.png` | enveloppes par registre |
| `clarinettes_balance.png` | figure de famille partagée |

La deuxième figure est la meilleure illustration pédagogique du test d'indépendance au rang
harmonique de tout le corpus : elle montre en un coup d'œil la différence entre un formant
et un effet de hauteur.

## Révision du 23/10/2026 — critère assoupli

Deux résultats, de sens opposé.

**Le cas de référence tient.** La zone à 466 Hz reste **rejetée** sous le critère assoupli :
+9,4 dB [IC +3,6] en dessous du rang 5, mais **−0,3 dB [IC −2,0]** au-dessus. Ce cas est
devenu le test de validité de l'assouplissement lui-même, et il le passe. Un critère qui
l'aurait retenu aurait été à rejeter.

**Un déclassement.** La zone à 5 270 Hz passe de *retenue* à **fragile-cr** : +4,7 dB
[IC +4,2] en dessous du rang 47, mais +0,0 dB [IC −0,6] au-dessus. Le découpage adaptatif au
rang médian révèle ce que les quatre groupes fixes masquaient. Le critère assoupli est donc
ici **plus** discriminant que le strict.

## Journal

- **2026-10-23** — Trois formants retenus (1 244, 2 487, 5 270 Hz). Trois méthodes
  indépendantes convergent sur la zone 1 200–1 250 Hz.
- **2026-10-23** — Candidat à 466 Hz rejeté : +11,3 dB aux rangs 2–3, 0,0 dB aux rangs 4–7.
  Cas retenu comme illustration de référence du filtre de rejet.
- **2026-10-23 (soir)** — Critère assoupli appliqué (voir `00-METHODE.md`).
