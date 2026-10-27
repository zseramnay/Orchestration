---
id: flute-contrebasse
nom: Flûte contrebasse
nom_en: Contrabass flute
famille: bois
sous_famille: flûtes
ordre: 4
source_echantillons: Yan_Adds-Divers/Winds/Contrabass-Flute/ordinario
provenance: VSL
technique: ordinario
n_echantillons: 38
tessiture_mesuree: G2–G5
f0_min_hz: 97.4
f0_max_hz: 784.0

classe: B
pertinence_formant: partielle  # revise 23/10
rho_hauteur: 0.76

formants:
  - { hz: 1759, statut: "1 corpus", residu_db: 2.30, etendue_residu: [1567, 5915], verdict: retenu-cr, notes: 5,
      ic_bas: [1.6, 1.0], commentaire: "retenu par critere assoupli 23/10 malgre 38 echantillons" }

candidats_rejetes:
  - { hz: 2792, residu_db: 1.26, verdict: rejete, raison: "résidus nuls et échec de robustesse" }

pente_db_octave: -11.5
pente_sd: 2.4
harmoniques_par_note: 39
harmoniques_min_max: [14, 90]
balance_impair_pair_db: -0.9
enveloppe_vraie_hz: 382
enveloppe_vraie_sd: 262

corpus_suffisant: false
version_mesure: 2026-10-23-wav-v1
---

# Flûte contrebasse

> ## Statut des formants — révision du 26/10/2026
>
> Confrontation des **cinq corpus**. Provenance de ce corpus : **VSL / ConTimbre / perso (Yan_Adds)**.
>
> | Formant | Statut |
> |---|---|
> | **1759 Hz** | présent dans le seul corpus SOL |
>
> `etendue_residu` dans l'en-tête **n'est pas une largeur de bande** : c'est l'étendue du résidu de la méthode 8, souvent une butée de plage. Les largeurs mesurées sont dans `_LARGEURS-DE-BANDE.md`.

## Un candidat fragile — corpus insuffisant

> **Méthode 8 appliquée — un candidat fragile, aucun retenu.** Voir `00-METHODE.md`.

Classe B, ρ = 0,76 : c'est la valeur la plus basse des quatre flûtes, ce qui est logique,
sa tessiture grave la rapprochant du domaine où la séparation source/filtre redevient
possible. Avec **39 harmoniques par note** en médiane, chaque note échantillonne densément
la réponse du tuyau.

Mais le corpus ne compte que **38 échantillons**, et c'est ce qui bloque. Le candidat à
1 759 Hz passe le test de robustesse de fenêtre et présente des résidus positifs cohérents,
mais il n'atteint 5 notes distinctes que dans un seul groupe de rang. Le verdict « fragile »
traduit un manque de données, non une contradiction.

Il faut donc lire ce fichier comme provisoire. C'est l'instrument du corpus des bois dont
le résultat changerait le plus probablement avec un corpus élargi.

## Candidats

| Zone | Résidu | k = 4–7 | k = 8–15 | k ≥ 16 | Robustesse | Verdict |
|---|---|---|---|---|---|---|
| **1 759 Hz** | +2,3 dB | +2,2 (3 notes) | +3,2 (5 notes) | +1,0 (1 note) | oui | fragile |
| 2 792 Hz | +1,3 dB | +0,7 | 0,0 | +2,7 | non | rejeté |

## Descripteurs valides

| Descripteur | Valeur | Lecture |
|---|---|---|
| **Pente spectrale médiane** | **−11,5 dB/octave** (σ 2,4) | la plus plate des quatre flûtes |
| Harmoniques par note | 39 [14–90] | spectre très riche |
| **Balance impair − pair** | **−0,9 dB** | quasi équilibrée, cas unique chez les flûtes |
| Enveloppe vraie | 382 Hz, σ 262 | σ la plus faible des flûtes |

Deux chiffres méritent attention.

La **pente de −11,5 dB/octave** complète la progression régulière de la famille : −21,2 pour
le piccolo, −18,3 pour la flûte, −12,5 pour la basse, −11,5 pour la contrebasse. Plus
l'instrument descend, plus son spectre s'aplatit. C'est le descripteur de famille, et il
remplace avantageusement les « formants » F1 antérieurs, qui ne faisaient que suivre les
tessitures respectives.

La **balance impair − pair de −0,9 dB** est singulière : toutes les autres flûtes sont
nettement négatives (−7,0 à −4,0 dB). Le spectre de la flûte contrebasse est donc à peu près
équilibré entre partiels pairs et impairs. À vérifier sur un corpus plus large avant d'en
tirer quoi que ce soit, mais si le fait se confirme, il distinguerait acoustiquement cet
instrument du reste de sa famille.

## Comparaison avec la littérature

| Source | Valeur annoncée | Notre mesure | Lecture |
|---|---|---|---|
| Yan_Adds (ancien) | F1 = 334 ± 95 Hz, /u/ | — | zone de la fondamentale |
| Ancien Fp | 1 092 Hz | — | artefact de fenêtre |
| Sources classiques | aucune donnée | — | instrument non étudié |

## Registres

Non défini dans `Data/registres.md`. À établir.

<!-- DEBUT-TRAVAUX-27-10 -->
## Ce que les travaux du 27/10/2026 disent de cet instrument

### Corroboration par le LPC, note par note

Rang du formant dans son propre spectre : 50 % signifie que les deux méthodes ne s'accordent pas mieux là qu'ailleurs. Voir `../_APPARIEMENT-LPC.md`.

| Formant | Statut | Rang |
|---|---|---|
| 1759 Hz | 1 corpus | 85,0 % |

### Effet de la nuance sur la localisation

Part de pics **non localisables** — l'enveloppe n'y redescend pas de 3 dB des deux côtés, donc le formant n'a pas de largeur mesurable. Voir `../_LARGEUR-PAR-REGISTRE.md`.

| | doux | moyen | fort |
|---|---|---|---|
| **tous registres** | 67 % | 50 % | 9 % |

**Effet doux → fort : +58 points.** La localisation s'améliore en jeu fort.

### Dans la table des mélanges

19 paires calculées. **Fusion moyenne : 18 %.** Voir `../_MELANGES.md`.

| Couleur nouvelle avec | | Se fond avec | |
|---|---|---|---|
| Hautbois | 35 % | Tuba basse | 79 % |
| Alto | 34 % | Tuba contrebasse | 68 % |
| Trompette en Ut | 31 % | Cor en Fa | 57 % |

<!-- FIN-TRAVAUX-27-10 -->

## Observations pour l'orchestration

À compléter. L'ancienne étude proposait des doublures avec la contrebasse (Δ = 162 Hz) et le
tuba basse (Δ = 108 Hz) sur la base de « zones graves /u/ partagées ». Ces écarts portaient
sur des zones de fondamentale, et ces doublures sont donc à remesurer intégralement.

Instrument rare, principalement employé pour ses qualités texturales en musique
contemporaine : la pente très plate et la balance équilibrée sont cohérentes avec cet usage,
puisqu'elles décrivent un son dense plutôt que focalisé.

## Graphiques

Aucun à ce stade. Figures prévues dans `../Versions-html-and-docx/media/` :

| Fichier prévu | Contenu |
|---|---|
| `flute_cb_composite.png` | enveloppe composite avec couverture par bande |
| `flutes_pentes.png` | figure de famille partagée — les quatre pentes |
| `flutes_balance.png` | balance impair/pair des quatre flûtes, mettant en évidence le cas de la contrebasse |

## Révision du 23/10/2026 — critère assoupli

La zone à **1 759 Hz** passe de *fragile* à **retenu-cr**, et c'est le gain le plus net de
l'assouplissement : **38 échantillons suffisent**. Le bootstrap donne +2,4 dB [IC +1,6] en
dessous du rang 10 et +1,7 dB [IC +1,0] au-dessus, sur 25 et 12 points. Les deux bornes
basses d'intervalle de confiance sont positives.

Les résidus restent modestes, autour de +2 dB, et le formant doit être lu comme tel : réel
mais faible. La pertinence formantique passe de « non pertinente » à « partielle ».

Le second candidat, à 2 792 Hz, reste **rejeté** : +0,2 dB [IC +0,0] sous le rang 17. Le
critère assoupli ne le sauve pas, ce qui est le comportement attendu.

## Journal

- **2026-10-23** — Un seul candidat, fragile, à 1 759 Hz. Résultat le plus provisoire de la
  famille : 38 échantillons seulement.
- **2026-10-23** — Balance impair/pair atypique (−0,9 dB) relevée comme piste, à confirmer.
- **2026-10-23 (soir)** — Critère assoupli appliqué (voir `00-METHODE.md`).
