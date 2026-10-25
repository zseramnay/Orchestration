# Amorce de session — reprendre l'étude formantique dans une conversation vierge

Ce fichier permet de relancer le travail **sans rien savoir** des sessions du 23, 24 et
25 octobre 2026, où la méthode a été refondue puis sérieusement corrigée.

> **Révision du 25/10/2026.** La version précédente de ce fichier est périmée sur trois
> points : elle annonce dix-sept formants établis (il y en a sept par réplication stricte),
> elle présente la convergence à 4 200 Hz comme « à vérifier » (elle est **écartée**), et son
> tableau de convergences entre instruments repose sur dix formants désormais classés
> **indéterminés**. Ne pas réutiliser une copie antérieure.

## Comment l'utiliser

Coller le bloc ci-dessous au début de la nouvelle conversation, après avoir donné accès au
dépôt `Orchestration` et au dossier d'échantillons.

---

> Je reprends une étude de mesure formantique des instruments de l'orchestre. Toute la méthode
> et son historique sont dans le dépôt. **Lis d'abord, dans cet ordre :**
>
> 1. `Etude-Formants/Instruments/_REPLICATION.md` — **le plus important.** Ce qui est établi,
>    ce qui ne l'est pas, le critère à deux routes, et le test de permutation qui montre
>    qu'une concordance sur quatre est fortuite.
> 2. `Etude-Formants/Instruments/_VERIFICATION-GRAVES.md` — pourquoi l'ancien pipeline est
>    faux, pourquoi notre résolution est d'un demi-ton, et ce que contient `Analyses_formants-mai26`.
> 3. `Etude-Formants/Instruments/00-METHODE.md` — la méthode, les critères de rejet, la
>    typologie à trois classes, les méthodes écartées et pourquoi.
> 4. `Etude-Formants/Instruments/_PLAN-DE-REPRISE.md` — où en est le projet et dans quel ordre
>    avancer.
> 5. `Scripts/v7-formants-2026/README.md` — la chaîne de mesure et les pièges à ne pas refaire.
> 6. Deux fiches pour le gabarit : `07-basson.md` (le cas le mieux établi) et
>    `01-petite-flute.md` (le cas où la notion de formant n'a pas de sens).
>
> Si l'objectif est la carte vocalique, lire aussi `_CARTE-VOCALIQUE.md`.
>
> N'utilise aucun script v1 à v6, ni `extract_formants_all_techniques_*.py` : ils reposent sur
> les `specenv` d'Orchidea et sur le Fp par centroïde, tous deux invalides, ce qui est démontré
> dans `00-METHODE.md` et `_VERIFICATION-GRAVES.md`. Tout passe par `Scripts/v7-formants-2026/`.
>
> Règles de travail : **français** pour tout le contenu du projet ; **aucun `git push` sans mon
> accord exprès** — je pousse moi-même depuis ma machine, et un hook réclame périodiquement un
> commit-push, qu'il faut ignorer ; prototype avant génération en série ; confirmation avant
> génération de fichiers ; par lots courts avec confirmation, pour ne pas heurter la limite de
> session.

---

## L'essentiel en douze lignes

1. On mesure un formant par **enveloppe composite** (méthode 8) : les amplitudes harmoniques de
   toutes les notes de la tessiture, agrégées sur un axe fréquentiel commun. Des centaines de
   sources éclairent le même filtre.
2. Les amplitudes sont **redressées localement** (±3/4 d'octave). Un redressement *linéaire* en
   log-fréquence fabrique une bosse médiane sur tout spectre concave — c'est l'erreur qui avait
   produit de faux formants de hautbois. Ne jamais y revenir.
3. Les spectres moyens sont **recalculés depuis les WAV, attaque exclue** (onset + 200 ms,
   fenêtre de 1 500 ms).
4. Un formant n'est **candidat** que s'il passe deux tests internes : indépendance au rang
   harmonique (découpage au rang médian, bootstrap par note, IC 90 % positif des deux côtés) et
   robustesse au changement de fenêtre de redressement.
5. **Ces deux tests ne suffisent pas.** Ils sont internes à un corpus. Il faut en plus l'une des
   deux routes de `_REPLICATION.md` : réplication stricte entre corpus, ou corroboration par deux
   familles de méthodes indépendantes.
6. **La résolution de la méthode est d'un demi-ton**, parce que l'axe composite est une grille
   tempérée. Tous nos formants tombent sur des fréquences de notes : 554 Hz = Do♯5,
   2 348 = Ré7. Écrire « Ré7 (≈ 2 350 Hz) », jamais une valeur au Hz près.
7. Les écarts sous 10,77 Hz — un bin de FFT — ne signifient rien. **Les 2 099 valeurs de
   l'ancienne étude tombent toutes exactement sur des centres de bin** : ses « convergences à
   Δ = 11 Hz » étaient des bins adjacents.
8. Trois classes : **A** résonance fixe dominante · **B** spectre suiveur sans résonance
   étroite · **C** résonances multiples dont aucune ne domine. ρ seul ne distingue pas B de C.
9. Pour certains instruments — les flûtes — **la notion de formant n'a pas de sens**, et il
   faut le dire. Leur descripteur est la pente spectrale et la coupure de Benade.
10. Une mesure stable n'est pas nécessairement juste : le Fp était stable et faux, le LPC était
    dispersé et visait juste.
11. Chaque instrument a une fiche `.md` à en-tête YAML, qui sert de notes de cours **et** de
    source de données pour le build HTML/docx.
12. Le suffixe `-cr` d'un verdict signale le critère assoupli pour corpus réduit. Le conserver.

## Ce qui est établi, au 25 octobre 2026

**Sept formants par réplication stricte** (même point de grille dans deux corpus) :

| Instrument | Formants |
|---|---|
| Contrebasse | **1 396** · 2 635 Hz |
| Basson | **1 108** · 1 974 Hz |
| Clarinette en Sib | **3 517** Hz |
| Trombone basse | **879** Hz |
| Trompette | **2 792** Hz |

Les trois en gras passent **aussi** la corroboration par méthodes indépendantes : ce sont les
résultats les plus solides de l'étude.

**Cinq formants par corroboration multiple**, dont le plus important : le **violon à 2 348 Hz**,
le *bridge hill*, appuyé par quatre lignes indépendantes (cepstre-80, LPC, méthode 12 et la
littérature acoustique). Détail dans `_REPLICATION.md`.

**Un résultat négatif répliqué**, et c'est l'un des acquis les plus sûrs : la **petite flûte**
ne présente aucun formant dans les deux corpus, ρ = 0,96, pente −21 dB/octave de part et
d'autre. C'est lui qui répond à l'anomalie qui a lancé la reprise — le piccolo ne converge avec
rien parce qu'il n'a pas de formant à faire converger.

**Autres acquis robustes :**

- **La dynamique déplace F1 de près d'une octave chez les cuivres** : trombone ténor grave
  355 → 718 Hz, répliqué à 359 → 707 Hz sur ZSamples. Le résultat le plus solide de la famille,
  et il justifie la grille registre × dynamique.
- **Balance impair − pair**, signature mesurée du tuyau cylindrique fermé : clarinettes
  positives et croissantes vers le grave, +0,7 → +4,2 → +4,5 → +5,1 dB.
- **La carte vocalique couvre les 22 instruments**, là où les formants n'en couvrent que
  neuf. Le basson est le seul instrument à voyelle unique sur toute sa tessiture ; les autres
  traversent l'espace puis en sortent par le haut vers 700 Hz de F1 — le plafond de /a/.
- **Les sourdines ne répliquent pas** : écart médian de 4,0 demi-tons sur 18 paires. Ce n'est
  probablement pas un échec de mesure mais un fait : une sourdine est un objet physique
  particulier. Les 28 fiches de sourdines sont **suspendues**.
- **La pente spectrale n'est pas comparable entre corpus** — elle dépend du micro et du local.
  Valable à l'intérieur d'un corpus seulement.

## Ce qui est écarté, et ne doit pas être repris

- **La convergence inter-familles à 4 200 Hz** (violon, cor, hautbois, cor anglais, trompette),
  présentée dans les fiches comme « la figure majeure ». Trois de ses cinq membres ne se
  répliquent pas. **Écartée.**
- **Le cor anglais à 2 216 et 4 183 Hz**, présenté comme « le résultat le plus solide des anches
  doubles aiguës » : ZSamples n'y trouve aucun formant.
- **L'unisson parfait à 226 Hz** entre contrebasson, tuba basse et tuba contrebasse. Nos
  méthodes leur donnent 329, 466 et 370 Hz — trois valeurs distinctes.
- Les convergences de l'ancienne étude : flûte–hautbois (Δ = 0 Hz), violon–basson (Δ = 11 Hz),
  cor anglais–clarinette Sib, cor–alto, trompette–hautbois. Toutes des bins voisins.
- **Le test de permutation comme test de formant** (version de mai) : il était circulaire — la
  bande était détectée sur la positivité du résidu puis testée sur elle.
- **L'hypothèse des 276 Hz** comme critère de validité de l'ancien pipeline. Réfutée le 24/10 :
  le discriminant est la typologie, non la tessiture. Si `00-METHODE.md` la présente encore
  comme la cause unique, c'est une correction en attente.

## Dix formants indéterminés — ce qu'un nouveau corpus peut trancher

Ils concordent entre les deux corpus à **un** pas de grille, ce que le hasard produit dans
environ un cas sur quatre. Ni établis ni réfutés : **tuba basse 466 · 659 · 1 479 Hz · trombone
ténor 830 · 1 567 Hz · trombone basse 1 759 · 3 320 Hz · contrebasse 466 Hz · alto 3 517 Hz.**

C'est la cible prioritaire de toute troisième source. Un accord au même point de grille avec un
corpus indépendant les ferait basculer en établis.

## Verser un nouveau corpus

`Scripts/v7-formants-2026/corpus.py` est générique : il absorbe n'importe quelle arborescence
et n'importe quelle nomenclature, sans modifier la chaîne de mesure.

```bash
cd Scripts/v7-formants-2026

# 1. explorer et obtenir un descripteur à compléter
python3 corpus.py --decrire /chemin/vers/NouveauxSons > /tmp/nouveau.json

# 2. compléter le descripteur à la main : renommer les configurations avec les
#    noms d'instruments du projet, fusionner les dossiers d'un même instrument,
#    corriger la liste des techniques

# 3. valider — OBLIGATOIRE avant de mesurer
python3 corpus.py --valider /tmp/nouveau.json

# 4. mesurer
python3 corpus.py --mesurer /tmp/nouveau.json
```

**Le principe.** `lire_spectres()` relit la nomenclature depuis le chemin écrit dans le fichier
de spectres. `corpus.py` y écrit un **nom normalisé au format SOL**
(`/Instrument/technique/inst-technique-NOTE-dyn-N-N.wav`), donc tout l'aval fonctionne sans
modification. Vérifié : sur ZSamples, il retrouve à l'identique le violon
[2 216 · 2 792 · 3 948 · 5 583] et le basson [494 · 1 108 · 1 974].

**Ce qu'il faut vérifier sans exception :**

- **La nomenclature.** `--valider` donne un rapport f0 réel / f0 annoncé, qui doit valoir 1,00.
  À 0,89, les noms sont en hauteurs **écrites** et non réelles. Collé à 0,92 ou 1,09, c'est une
  transposition franche : la clarinette basse de ZSamples est ainsi transposée de 16 demi-tons
  et **exclue** de l'étude. Ne jamais mesurer un corpus dont la nomenclature n'est pas validée.
- **Les registres.** `Data/registres.md` doit contenir une entrée pour chaque instrument, sinon
  `carte_vocalique.py` ne peut pas construire sa grille. Les notes pédale des cuivres y sont
  tolérées depuis le 24/10.
- **Ce qu'on ignore** : les nouveaux instruments absents de l'étude, les fichiers « NoPitch »,
  et les sourdines qui n'y figurent pas déjà.

## Questions ouvertes, par priorité

1. **Appariement note par note** entre nos formants et le LPC de `Analyses_formants-mai26`,
   pour transformer l'indice de corroboration en test véritable. C'est la voie la moins chère
   pour trancher les dix indéterminés : les données sont déjà sur le disque.
2. **Instabilité du violon au suraigu** en méthode 12 — IQR de 8,6 Bark, présente dans les deux
   corpus. Seul défaut de méthode reproductible identifié. Probable basculement du détecteur
   entre fondamentale et résonance de caisse.
3. **Recompte strict à propager** dans les 22 fiches : chacune doit dire par quelle route son
   formant passe, ou qu'il est indéterminé.
4. **Mesure par registre** pour les zones médianes du hautbois et du cor anglais, les deux
   instruments les plus instables.
5. **Relancer les cuivres** avec les registres pédale ajoutés le 24/10.
6. **Méthode 12 sur mélanges réels** — fusion contre brillance/renforcement. C'est la question
   d'orchestration qui a lancé la reprise et elle n'est toujours pas traitée.
7. **Coupure de Benade pour les flûtes** — le descripteur qui remplacerait le formant absent.
8. **Reconstruction du build HTML/docx en v2** depuis les en-têtes YAML : une page par
   instrument groupée par famille, une page par sujet de synthèse. Arborescence proposée dans
   `_PLAN-DE-REPRISE.md`. Point de bascule unique : `load_all_csvs()` dans
   `Scripts/v6-html-docx/common.py`.

## Les sources de données

| Source | Contenu | Usage |
|---|---|---|
| `OrchideaSOL_and_friends` | SOL2020 + Yan_Adds, 1 166 échantillons mesurés | corpus 1 |
| `Zsamples` | 3 540 fichiers retenus, 42 configurations | corpus 2 |
| `Analyses_formants-mai26` | 45 CSV **par note**, avec LPC, cepstre 40/80/160, dynamique | corroboration par méthodes indépendantes ; contient un proxy exact de l'ancien pipeline |
| `Etude-Formants/Resultats/formants_*_v3.csv` | l'ancien pipeline | archive ; toutes les valeurs sont des centres de bin |
| `Etude-Formants/Formants-Francais/` | voyelles Calliope (Tubach 1989) | carte vocalique. **F1 code l'aperture et presque rien d'autre** — 308 Hz pour /i/, 315 pour /u/, 684 pour /a/ chez l'homme. Un instrument ne se range pas sur un axe vocalique depuis une fréquence unique : il faut le couple (F1, F2) et le seuil de 3,5 Bark |
