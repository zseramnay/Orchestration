#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mesurer.py — chaîne complète pour UNE configuration instrumentale.
Dépend uniquement de formants_lib.py (même dossier), numpy et soundfile.

    python3 mesurer.py --racine <base> --dossier <Famille/Instrument> \
                       --technique ordinario --nom "Violon" --sortie resultats/

Produit dans --sortie :
    <nom>.spectres.txt   spectres moyens, attaque exclue (format Orchidea)
    <nom>.json           mesures complètes
    <nom>.md             fiche pré-remplie (avec --fiche)

Pour ne refaire que l'analyse sans relire les WAV : --depuis-spectres <fichier>

Exemples
--------
# un instrument nu
python3 mesurer.py --racine ~/kDrive/Samples/Orchidea-dbs/OrchideaSOL_and_friends/FullSOL2020 \
  --dossier Strings/Violin --technique ordinario --nom Violon --sortie res/ --fiche

# une sourdine (deux techniques fusionnées, ex. wah)
python3 mesurer.py --racine .../FullSOL2020 --dossier "Brass/Trumpet_C+sordina_wah" \
  --technique ordinario_open --nom "Trompette+wah ouvert" --sortie res/

# toute une base d'un coup
python3 mesurer.py --racine .../Yan_Adds-Divers --balayer Winds --technique ordinario --sortie res/
"""
import os, sys, json, argparse, glob
import formants_lib as FL


def fiche_markdown(nom, res, dossier, technique):
    """Squelette de fiche conforme au gabarit du projet."""
    f = res.get('formants', [])
    retenus = [x for x in f if x['verdict'] == 'retenu-cr']
    fragiles = [x for x in f if x['verdict'] == 'fragile-cr']
    rejetes = [x for x in f if x['verdict'] in ('rejeté', 'non mesurable')]
    rho = res.get('rho')
    # classe proposée, à VALIDER à la main (voir 00-METHODE.md)
    if rho is None:
        classe = "?"
    elif rho < 0.55:
        classe = "A"
    elif retenus:
        classe = "C"          # rho élevé MAIS des formants : résonances multiples
    else:
        classe = "B"
    pert = "pertinente" if len(retenus) >= 2 else ("partielle" if retenus or fragiles
                                                   else "non pertinente")
    L = []
    L.append("---")
    L.append(f"id: {nom.lower().replace(' ', '-').replace('+', '-')}")
    L.append(f"nom: {nom}")
    L.append(f"source_echantillons: {dossier}/{technique}")
    L.append(f"technique: {technique}")
    L.append(f"n_echantillons: {res['n']}")
    L.append(f"n_notes_distinctes: {res['n_notes']}")
    L.append(f"f0_min_hz: {res['f0_min']}")
    L.append(f"f0_max_hz: {res['f0_max']}")
    L.append(f"controle_nomenclature: {res['nomenclature']}   # doit valoir ~1,00")
    L.append(f"classe: {classe}        # A/B/C — À VALIDER À LA MAIN")
    L.append(f"pertinence_formant: {pert}   # À VALIDER À LA MAIN")
    L.append(f"rho_hauteur: {rho}")
    L.append("formants:")
    for x in f:
        g = x.get('groupes', [])
        ic = [y.get('ic_bas') for y in g if y.get('testable')]
        L.append(f"  - {{ hz: {x['hz']}, residu_db: {x['residu']}, "
                 f"bande_db3: [{x['bande_bas']}, {x['bande_haut']}], "
                 f"verdict: {x['verdict']}, notes: {x['n_notes']}, "
                 f"rang_median: {x['rang_median']}, ic_bas: {ic} }}")
    if not f:
        L[-1] = "formants: []"
    L.append(f"pente_db_octave: {res['pente']}")
    L.append(f"pente_sd: {res['pente_sd']}")
    L.append(f"harmoniques_par_note: {res['harmoniques_median']}")
    L.append(f"harmoniques_min_max: [{res['harmoniques_min']}, {res['harmoniques_max']}]")
    L.append(f"balance_impair_pair_db: {res['balance_impair_pair']}")
    L.append(f"enveloppe_vraie_hz: {res['env_vraie']}")
    L.append(f"enveloppe_vraie_sd: {res['env_vraie_sd']}")
    L.append(f"corpus_suffisant: {str(res['n'] >= 50).lower()}")
    L.append("version_mesure: 2026-10-24-wav-v1")
    L.append("---")
    L.append("")
    L.append(f"# {nom}")
    L.append("")
    L.append("## Méthode retenue, et pourquoi")
    L.append("")
    L.append(f"> **Méthode 8 — enveloppe composite.** {len(retenus)} formant(s) retenu(s), "
             f"{len(fragiles)} fragile(s), {len(rejetes)} rejeté(s). Voir `00-METHODE.md`.")
    L.append("")
    L.append(f"À RÉDIGER. Éléments de fait : ρ = {rho} · {res['harmoniques_median']} harmoniques "
             f"par note en médiane · {res['n']} échantillons sur {res['n_notes']} notes "
             f"distinctes · pente σ = {res['pente_sd']} dB/octave.")
    L.append("")
    L.append("## Formants")
    L.append("")
    L.append("| Zone | Résidu | Largeur −3 dB | Notes | Verdict |")
    L.append("|---|---|---|---|---|")
    for x in f:
        L.append(f"| {x['hz']:.0f} Hz | {x['residu']:+.1f} dB | "
                 f"{x['bande_bas']:.0f}–{x['bande_haut']:.0f} Hz | {x['n_notes']} | "
                 f"{x['verdict']} |")
    L.append("")
    L.append("## Tests de rejet")
    L.append("")
    L.append("| Zone | rang ≤ médian | rang > médian | Robustesse |")
    L.append("|---|---|---|---|")
    for x in f:
        g = x.get('groupes', [])
        def cell(i):
            if i >= len(g) or not g[i].get('testable'):
                return f"non testable ({g[i]['n_notes']} notes)" if i < len(g) else "—"
            return f"{g[i]['mediane']:+.1f} dB [IC {g[i]['ic_bas']:+.1f}], {g[i]['n_notes']} notes"
        L.append(f"| {x['hz']:.0f} Hz (k méd. {x['rang_median']}) | {cell(0)} | {cell(1)} | "
                 f"{'oui' if x.get('robustesse') else 'non'} |")
    L.append("")
    L.append("## Descripteurs globaux")
    L.append("")
    L.append("| Descripteur | Valeur |")
    L.append("|---|---|")
    L.append(f"| Pente spectrale médiane | {res['pente']} dB/octave (σ {res['pente_sd']}) |")
    L.append(f"| Harmoniques par note | {res['harmoniques_median']} "
             f"[{res['harmoniques_min']}–{res['harmoniques_max']}] |")
    L.append(f"| Balance impair − pair | {res['balance_impair_pair']} dB |")
    L.append(f"| Enveloppe vraie (méthode 6) | {res['env_vraie']:.0f} Hz, "
             f"σ {res['env_vraie_sd']:.0f} ; ordres {res['ordre_min']}–{res['ordre_max']} |")
    L.append(f"| Pics tombant sur un harmonique | {res['pics_sur_harmonique']:.1%} |")
    L.append("")
    L.append("## Comparaison avec la littérature")
    L.append("")
    L.append("À RÉDIGER. Vérifier aux sources primaires : Backus (1969), Giesler (1985), "
             "Meyer (2009), et comparer aux valeurs de l'ancien pipeline.")
    L.append("")
    L.append("## Observations pour l'orchestration")
    L.append("")
    L.append("À RÉDIGER.")
    L.append("")
    L.append("## Journal")
    L.append("")
    L.append("- **2026-10-24** — Première mesure (`mesurer.py`, méthode 8 + critère assoupli).")
    L.append("")
    return "\n".join(L)


def traiter(racine, dossier, techniques, nom, sortie, fiche, depuis=None):
    os.makedirs(sortie, exist_ok=True)
    sp = depuis or os.path.join(sortie, f"{nom}.spectres.txt")
    if not depuis:
        ok, ko, tot = FL.ecrire_spectres(racine, dossier, techniques, sp)
        print(f"  spectres : {ok} écrits, {ko} échecs, {tot} fichiers")
        if ok == 0:
            print("  !! aucun spectre — vérifiez --dossier et --technique")
            return None
    res = FL.mesurer(sp)
    if 'erreur' in res:
        print("  !!", res['erreur']); return None
    if res['nomenclature'] is None or abs(res['nomenclature'] - 1.0) > 0.03:
        print(f"  !! CONTRÔLE DE NOMENCLATURE : {res['nomenclature']} (attendu ~1,00). "
              "Les noms de fichiers ne sont peut-être pas en hauteurs réelles.")
    res['_config'] = dict(racine=racine, dossier=dossier, techniques=techniques, nom=nom)
    with open(os.path.join(sortie, f"{nom}.json"), "w") as fh:
        json.dump(res, fh, ensure_ascii=False)
    v = [f"{x['hz']:.0f}:{x['verdict']}" for x in res['formants']]
    print(f"  ρ={res['rho']} pente={res['pente']} odd={res['balance_impair_pair']} | "
          + ("  ".join(v) if v else "aucun candidat"))
    if fiche:
        p = os.path.join(sortie, f"{nom}.md")
        with open(p, "w") as fh:
            fh.write(fiche_markdown(nom, res, dossier, techniques[0]))
        print(f"  fiche : {p}")
    return res


def main():
    ap = argparse.ArgumentParser(description="Mesure formantique d'une configuration.")
    ap.add_argument("--racine", help="base de samples (…/FullSOL2020 ou …/Yan_Adds-Divers)")
    ap.add_argument("--dossier", help="Famille/Instrument, ex. Strings/Violin")
    ap.add_argument("--balayer", help="au lieu de --dossier : balaie tous les instruments "
                                      "d'une famille, ex. Winds")
    ap.add_argument("--technique", default="ordinario",
                    help="une ou plusieurs techniques séparées par des virgules")
    ap.add_argument("--nom", help="nom d'affichage (défaut : nom du dossier)")
    ap.add_argument("--sortie", default="resultats", help="dossier de sortie")
    ap.add_argument("--fiche", action="store_true", help="écrire aussi la fiche .md")
    ap.add_argument("--depuis-spectres", help="réanalyser un fichier de spectres existant")
    a = ap.parse_args()
    techs = [t.strip() for t in a.technique.split(",")]

    if a.depuis_spectres:
        nom = a.nom or os.path.basename(a.depuis_spectres).split('.')[0]
        print(nom); traiter(None, None, techs, nom, a.sortie, a.fiche, a.depuis_spectres); return

    if a.balayer:
        base = os.path.join(a.racine, a.balayer)
        for d in sorted(glob.glob(os.path.join(base, "*"))):
            if not os.path.isdir(d):
                continue
            inst = os.path.basename(d)
            dispo = [t for t in techs if os.path.isdir(os.path.join(d, t))]
            if not dispo:
                print(f"{inst} : aucune des techniques {techs} — ignoré"); continue
            print(inst)
            traiter(a.racine, f"{a.balayer}/{inst}", dispo, inst, a.sortie, a.fiche)
        return

    if not (a.racine and a.dossier):
        ap.error("--racine et --dossier sont requis (ou --balayer, ou --depuis-spectres)")
    nom = a.nom or os.path.basename(a.dossier)
    print(nom)
    traiter(a.racine, a.dossier, techs, nom, a.sortie, a.fiche)


if __name__ == "__main__":
    main()
