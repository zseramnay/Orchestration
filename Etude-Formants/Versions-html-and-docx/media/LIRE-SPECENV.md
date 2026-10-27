# Les 47 figures `specenv_*` sont périmées — ne pas les réutiliser

**Date : 26 octobre 2026.**

Ces figures appartiennent à l'ancienne étude (`index.html`, version de mai 2026). Elles sont
conservées ici parce qu'elles font partie de cette étude telle qu'elle a été publiée, et que la
page `Site-v2/syntheses/heritage.html` en rend compte. **Elles ne doivent pas servir à la
version 2.**

## Pourquoi

Elles montrent des pics qui ont été diagnostiqués comme des artefacts : les « formants » qu'elles
dessinent tombent exactement sur des centres de bin de FFT (2 099 valeurs F1–F4 sur un pas de
10,7666 Hz, sans une exception), et les « convergences » qu'on y lisait étaient des bins voisins,
jusqu'à Δ = 11 Hz — c'est-à-dire un seul bin d'écart. Ce n'était pas une mesure de l'instrument.

## Ce qui les remplace

`Scripts/v7-site/figures.py` produit `Etude-Formants/Site-v2/figures/` :

| Figure | Ce qu'elle montre |
|---|---|
| `composite-<id>.png` | le résidu local médian, **une courbe par corpus** — un formant retenu est un maximum que deux corpus de provenances distinctes placent au même point de grille |
| `vocalique-<id>.png` | les cellules registre × dynamique dans le plan (F1, F2) en Bark, contre les voyelles du français (Calliope, voix d'homme) |
| `synthese-*.png` | convergences, ρ par corpus, pentes, balance impair − pair, largeurs confrontées |

La différence tient en une phrase : l'ancienne figure montrait **un** spectre et ses pics ; la
nouvelle en superpose **plusieurs, de provenances différentes**, et ne retient que ce qui survit
à la superposition.
