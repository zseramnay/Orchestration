# Amorce de session — reprendre l'étude formantique dans une conversation vierge

Ce fichier existe pour une raison précise : permettre de relancer le travail **sans rien
savoir** de la session du 23–24 octobre 2026 où la méthode a été refondue.

## Comment l'utiliser

Coller le bloc ci-dessous au début de la nouvelle conversation, après avoir donné accès au
dépôt `Orchestration` et au dossier d'échantillons.

---

> Je reprends une étude de mesure formantique des instruments de l'orchestre. Toute la
> méthode et son historique sont dans le dépôt. **Lis d'abord, dans cet ordre :**
>
> 1. `Etude-Formants/Instruments/00-METHODE.md` — la méthode, les critères de rejet, la
>    typologie à trois classes, et surtout la liste des méthodes écartées et pourquoi.
> 2. `Scripts/v7-formants-2026/README.md` — la chaîne de mesure et les huit pièges à ne pas
>    refaire.
> 3. `Etude-Formants/Instruments/synthese-formants-2026-10-23.csv` — l'état des mesures,
>    54 configurations.
> 4. Une ou deux fiches d'instrument pour le gabarit : `07-basson.md` (cas le mieux établi)
>    et `01-petite-flute.md` (cas où la notion de formant n'a pas de sens).
>
> Ne réutilise pas les scripts v1 à v6 ni `extract_formants_all_techniques_*.py` : ils
> reposent sur les `specenv` d'Orchidea et sur le Fp par centroïde, tous deux invalides pour
> cet usage, ce qui est démontré dans `00-METHODE.md`. Tout passe par
> `Scripts/v7-formants-2026/mesurer.py`.
>
> Règles de travail : français pour tout le contenu du projet ; aucun `git push` sans mon
> accord exprès ; prototype avant génération en série.

---

## L'essentiel en dix lignes, si vous devez l'expliquer de mémoire

1. On mesure un formant par **enveloppe composite** (méthode 8) : les amplitudes harmoniques
   de toutes les notes de la tessiture, agrégées sur un axe fréquentiel commun. Des centaines
   de sources différentes éclairent le même filtre.
2. Les amplitudes sont **redressées localement** (±3/4 d'octave) pour ne laisser passer que
   les résonances plus étroites que ~1,5 octave.
3. Les spectres moyens sont **recalculés depuis les WAV, attaque exclue**. Les `specenv`
   d'Orchidea ne conviennent pas : lissage d'ordre fixe 160, valide seulement sous 276 Hz.
4. Un formant n'est publié que s'il passe **deux tests** : indépendance au rang harmonique
   (découpage au rang médian, bootstrap par note, IC à 90 % positif des deux côtés) et
   robustesse au changement de fenêtre de redressement.
5. Trois classes : **A** résonance fixe dominante · **B** spectre suiveur sans résonance
   étroite · **C** résonances multiples dont aucune ne domine. ρ seul ne distingue pas B de C.
6. Pour certains instruments — les flûtes — **la notion de formant n'a pas de sens**, et il
   faut le dire. Leur descripteur est la pente spectrale et la coupure de Benade.
7. Les écarts sous 10,77 Hz, soit un bin de FFT, ne signifient rien.
8. Une mesure stable n'est pas nécessairement juste : le Fp était stable et faux, le LPC était
   dispersé et visait juste.
9. Chaque instrument a une fiche `.md` à en-tête YAML, qui sert à la fois de notes de cours et
   de source de données pour le build HTML/docx.
10. Le suffixe `-cr` d'un verdict signale le critère assoupli pour corpus réduit. Le conserver.

## État au 24 octobre 2026

**Fait — 23 fiches rédigées** : les 12 bois (piccolo → clarinette contrebasse), les 6 cuivres
principaux, les 4 cordes solistes, plus `00-METHODE.md`.

**Mesuré mais pas encore rédigé — 28 configurations.** Les données sont dans
`synthese-formants-2026-10-23.csv` et dans les JSON `variantes-*.json` :

| Groupe | Nombre | Fichiers à créer |
|---|---|---|
| Sourdines de cordes | 7 | 23 à 29 |
| Ensembles de cordes et ensembles avec sourdine | 7 | 30 à 36 |
| Sourdines de cor, tuba et trompette | 7 | 37 à 43 |
| Sourdines de trombone (5) et de bois (2) | 7 | 44 à 50 |

Format décidé : **compact**. Un en-tête YAML, le tableau des formants, les tests de rejet,
puis trois paragraphes sur ce que la sourdine ou le pupitre change par rapport à l'instrument
nu. Une variante n'a pas besoin d'une revue de littérature : elle renvoie à la fiche de
l'instrument nu.

**Questions ouvertes, par ordre de priorité :**

1. **Mesure par registre** — tranche la zone médiane du hautbois (1 100–1 700 Hz, annoncée
   par Meyer et Giesler, détectée seulement sur la moitié grave) et celle du cor anglais
   (1 174 Hz à +10 dB sur trois sous-ensembles mais absente du global).
2. **Analyse par dynamique** — c'est le test décisif pour les cuivres, non une amélioration.
   Hypothèse : chez eux la résonance existe mais sa position dépend de la force de jeu, donc
   la bande agrégée sur toutes les dynamiques est étalée et échoue au test de robustesse.
   Cela expliquerait les cas du cor à 466 Hz et de la trompette à 1 479 Hz, tous deux
   fragiles alors qu'ils correspondent à la littérature.
3. **Coupure de Benade pour les flûtes** — le descripteur qui remplacerait le formant absent.
4. **Méthode 12 : motifs d'excitation auditive sur mélanges réels** — pour les doublures, avec
   deux fonctions à distinguer, fusion et brillance/renforcement. Non implémentée.
5. **Rapprochement avec les formants du français** — données Calliope dans
   `Etude-Formants/Formants-Francais/`. Attention : F1 code l'aperture et presque rien
   d'autre (308 Hz pour /i/, 315 Hz pour /u/, 684 Hz pour /a/ chez l'homme). Un instrument ne
   peut pas être rangé sur un axe vocalique depuis une fréquence unique : il faut le couple
   (F1, F2) et le seuil de 3,5 Bark sur F2 − F1.
6. **Reconstruction du build HTML/docx** depuis les en-têtes YAML des fiches, au lieu des
   valeurs en dur des versions v1 à v6.

**Résultats à ne pas perdre de vue :**

- Le basson à **554 Hz** est le seul chiffre de l'ancienne étude qui ressort intact, confirmé
  par quatre voies indépendantes.
- Le violoncelle a un formant à **554 Hz également** — coïncidence exacte entre deux familles,
  chacune validée séparément. C'est la convergence la plus précise de la reprise.
- La clarinette en Sib a **deux** formants, 1 567 et 3 517 Hz, ce qui réconcilie Backus/Meyer
  et Giesler : ils ne se contredisaient pas, ils mesuraient deux formants distincts.
- Le violon à **2 348 Hz** correspond au *bridge hill* de la littérature acoustique.
- Une convergence inter-familles autour de **4 200 Hz** réunit violon, cor, hautbois, cor
  anglais et trompette, chacun validé séparément. À vérifier : effet réel ou artefact de la
  chaîne de mesure ?
- **Les sourdines changent la pente spectrale, pas seulement le niveau** : trompette harmon
  −3,5 dB/octave contre −16,8 nue ; violon piombo −5,8 contre −13,5.
- **Les sourdines rendent les cuivres plus « classe A »** : trombone wah fermé ρ = −0,09,
  minimum absolu du corpus.
- L'hypothèse selon laquelle les **pupitres seraient de meilleurs témoins formantiques que les
  solistes n'est pas soutenue**, et elle était de toute façon non testable : les solistes
  viennent de SOL2020 (IRCAM) et les ensembles de Yan_Adds (VSL, ConTimbre, personnels), donc
  toute différence est confondue avec la chaîne de prise de son.

## Convergences mesurées entre instruments

À ne pas confondre avec les « convergences » de l'ancienne étude, qui comparaient des
médianes de tessiture tombant dans des bins voisins. Toutes celles-ci portent sur des formants
validés indépendamment.

| Zone | Instruments |
|---|---|
| 554 Hz | basson · violoncelle |
| 1 108–1 174 Hz | basson · contrebasson · violoncelle |
| 1 396–1 479 Hz | contrebasse · tuba contrebasse · tuba basse |
| 1 567 Hz | clarinette en Sib · trombone ténor |
| 2 091–2 216 Hz | cor · cor anglais · trombone ténor · violoncelle · alto |
| 3 517 Hz | clarinette en Sib · alto |
| 3 948 Hz | violoncelle · contrebasse · hautbois+sourdine |
| 4 183–4 695 Hz | cor · hautbois · cor anglais · trompette · violon |
| 466 Hz | contrebasse · tuba basse |
| 879 Hz | trombone basse · tuba contrebasse |

## Convergences de l'ancienne étude qui sont écartées

Ne pas les reprendre sans nouvelle mesure : flûte–hautbois (Δ = 0 Hz), violon–basson
(Δ = 11 Hz), cor anglais–clarinette Sib (Δ = 11 Hz), cor–alto (Δ = 11 Hz),
trompette–hautbois (Δ = 43 Hz), et l'« unisson parfait » à 226 Hz entre contrebasson, tuba
basse et tuba contrebasse. Toutes reposaient sur des hauteurs médianes de tessiture dans des
bins de FFT voisins.
