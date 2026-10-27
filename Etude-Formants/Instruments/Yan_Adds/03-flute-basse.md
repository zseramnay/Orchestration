---
id: flute-basse
nom: Flûte basse
nom_en: Bass flute
famille: bois
sous_famille: flûtes
ordre: 3
source_echantillons: Yan_Adds-Divers/Winds/Bass-Flute/ordinario
provenance: VSL
technique: ordinario
n_echantillons: 39
tessiture_mesuree: C3–C6
f0_min_hz: 130.8
f0_max_hz: 1046.5

classe: B
pertinence_formant: non pertinente
rho_hauteur: 0.84

formants: []

candidats_rejetes:
  - { hz: 2635, residu_db: 1.04, verdict: rejete, raison: "résidus nuls par rang harmonique et échec de robustesse" }

pente_db_octave: -12.5
pente_sd: 4.7
harmoniques_par_note: 29
harmoniques_min_max: [10, 83]
balance_impair_pair_db: -4.0
enveloppe_vraie_hz: 420
enveloppe_vraie_sd: 776

corpus_suffisant: false
version_mesure: 2026-10-23-wav-v1
---

# Flûte basse

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **VSL / ConTimbre / perso (Yan_Adds)**.
>
> Aucun formant retenu. Si le résultat négatif est répliqué, c'est un acquis — voir `_CORPUS-CROISES.md`.

## Aucun formant retenu — corpus insuffisant et classe B

> **Méthode 8 appliquée — aucun formant retenu.** Voir `00-METHODE.md`.

Deux causes se cumulent, et il importe de ne pas les confondre.

**Cause structurelle.** La flûte basse est de classe B, ρ = 0,84 : son spectre suit la
hauteur. L'écart-type de l'enveloppe vraie, σ = 776 Hz, est le deuxième plus élevé du
corpus des bois après le piccolo. Comme pour toute la famille, l'invariant pertinent est la
coupure, non une résonance de corps.

**Cause contingente.** Le corpus ne compte que **39 échantillons**. C'est trop peu pour
peupler les groupes de rang harmonique du test d'indépendance : le seul candidat détecté
n'atteint 4 notes distinctes dans aucun groupe, contre 5 requises. La conclusion « aucun
formant » est donc ici moins solide que pour le piccolo, où elle reposait sur 111
échantillons et sur un argument de principe. **Un corpus élargi pourrait changer ce
résultat.**

## Candidat rejeté

| Zone | Résidu | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse | Verdict |
|---|---|---|---|---|---|---|
| 2 635 Hz | +1,0 dB | +0,1 | 0,0 | +0,9 | non | **rejeté** |

Les résidus sont quasi nuls sur tous les groupes de rang : ce n'est pas un problème de
couverture mais une absence de signal. Le rejet est franc.

## Descripteurs valides

| Descripteur | Valeur | Lecture |
|---|---|---|
| **Pente spectrale médiane** | **−12,5 dB/octave** (σ 4,7) | près de deux fois plus plate que le piccolo |
| Harmoniques par note | 29 [10–83] | spectre riche |
| Balance impair − pair | −4,0 dB | tuyau ouvert |
| Enveloppe vraie | 420 Hz, σ 776 | σ considérable : pas de formant fixe |

La pente est le résultat le plus parlant. À −12,5 dB/octave contre −21,2 pour le piccolo et
−18,3 pour la grande flûte, elle quantifie précisément ce que l'on décrit comme un « son
chaud et soufflé, plus riche en harmoniques » : la flûte basse conserve beaucoup plus
d'énergie dans l'aigu relatif. La dispersion σ = 4,7 traduit en outre une forte variation
selon la dynamique et le souffle.

C'est un bon exemple de ce que l'abandon du formant fait gagner : la pente décrit
l'instrument mieux que ne le faisait un F1 de 301 Hz, qui était la zone de la fondamentale.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Lecture |
|---|---|---|---|
| Yan_Adds (ancien) | F1 = 301 ± 92 Hz, /u/ | — | zone de la fondamentale, pas un formant |
| Ancien Fp | 1 338 Hz | — | artefact de fenêtre |
| Backus, Giesler, Meyer | aucune donnée | — | instrument non étudié dans les sources classiques |

Aucune source académique classique ne traite la flûte basse. Il n'y a donc rien à confirmer
ni à contredire, et les valeurs antérieures ne reposaient que sur le pipeline `specenv`.

## Registres

Non défini dans `Data/registres.md`. À établir.

<!-- DEBUT-TRAVAUX-27-10 -->
## Ce que les travaux du 27/10/2026 disent de cet instrument

### Dans la table des mélanges

18 paires calculées. **Fusion moyenne : 15 %.** Voir `../_MELANGES.md`.

| Couleur nouvelle avec | | Se fond avec | |
|---|---|---|---|
| Trombone basse | 42 % | Clarinette en Sib | 61 % |
| Hautbois | 30 % | Clarinette basse en Sib | 61 % |
| Trombone ténor | 30 % | Cor en Fa | 53 % |

<!-- FIN-TRAVAUX-27-10 -->

## Observations pour l'orchestration

À compléter. Pistes ouvertes :

- L'ancienne étude annonçait une « excellente » convergence flûte basse–clarinette basse à
  Δ = 22 Hz (301 contre 323 Hz). Les deux valeurs étant des zones de fondamentale et non
  des formants, cette convergence est à reprendre entièrement. Elle n'est ni confirmée ni
  infirmée : elle n'a simplement jamais été mesurée.
- La pente très plate suggère un rôle de soutien harmonique plutôt que de projection
  mélodique dans les mélanges, hypothèse à tester par la méthode 12.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `flute_basse_composite.png` | enveloppe composite, avec la couverture par bande montrant l'insuffisance du corpus |
| `flutes_pentes.png` | comparaison des quatre pentes de la famille (piccolo → contrebasse) — figure de famille |

La seconde figure est à partager avec les trois autres flûtes : c'est la pente, et non le
formant, qui différencie la famille.

## Journal

- **2026-10-23** — Aucun formant retenu. À distinguer du piccolo : ici la cause principale
  est l'insuffisance du corpus (39 échantillons), non un argument de principe. À revoir si
  le corpus est élargi.
