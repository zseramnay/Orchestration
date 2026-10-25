#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zsamples.py — mesure formantique du corpus ZSamples.

ZSamples diffère de SOL2020 sur deux points de structure :
  · la TECHNIQUE est dans le NOM DE FICHIER, pas dans un sous-dossier ;
  · la profondeur varie (flutes/flute/*.wav mais violin/*.wav).
Le format audio est en revanche identique : 44 100 Hz, mono, PCM 16.

Les cordes sont de plus organisées PAR CORDE (string_I … string_V), ce que SOL ne
permettait pas. D'où l'option --test-cordes : une même hauteur jouée sur deux cordes
différentes, c'est deux sources pour un même corps. Si un formant est une résonance de
caisse, il doit être identique d'une corde à l'autre. C'est un test d'indépendance plus
fort que celui du rang harmonique.

IMPORTANT — ne jamais fusionner ZSamples et SOL2020 dans une même mesure. Chaînes de prise
de son différentes : toute différence serait confondue avec le matériel d'enregistrement.
Mesurés séparément, les deux corpus forment une RÉPLICATION INDÉPENDANTE, ce qui vaut mieux
qu'un corpus agrandi.

Usage
-----
    python3 zsamples.py --racine ~/Documents/Instrumental_formants/Zsamples \\
                        --sortie resultats_zsamples/ --fiche
    python3 zsamples.py --racine ... --sortie ... --seulement Violon,Basson
    python3 zsamples.py --racine ... --sortie ... --test-cordes
    python3 zsamples.py --racine ... --lister        (inventaire sans mesurer)
"""
import os, sys, json, argparse, collections
import formants_lib as FL
from mesurer import fiche_markdown

# ─────────────────────────────────────────────────────────────────────────────
# CORRESPONDANCE ZSamples -> étude  (arbitrage de Yan Maresz, 24/10/2026)
#
# Instruments : seuls ceux DÉJÀ dans l'étude. Sont donc ignorés : flûte alto,
# hautbois baryton, trompette piccolo, trombone contrebasse, tous les saxophones
# (y compris l'alto, écarté explicitement), tubax.
# Sourdines : seules celles déjà dans l'étude. Sont donc ignorées : bucket, whisper,
# silent_brass, stop_mute (bouché), insert_straight_mute, cup_closed, cup_half_open,
# wahwah_stem_out, wahwah half_open, half_closed, con_sordino_wood.
# Les fichiers NoPitch sont écartés automatiquement (aucune note dans le nom).
# ─────────────────────────────────────────────────────────────────────────────
ORD = {"ordinario": ["ordinario"]}

CONFIG = {
    # dossier ZSamples                    nom affiché            catégorie -> techniques fusionnées
    "flutes/piccolo_flute":              ("Petite flûte",        ORD),
    "flutes/flute":                      ("Flûte",               ORD),
    "flutes/bass_flute":                 ("Flûte basse",         ORD),
    "flutes/contrabass_flute":           ("Flûte contrebasse",   ORD),
    "oboes/oboe":                        ("Hautbois",            ORD),
    "oboes/english_horn":                ("Cor anglais",         ORD),
    "Bassons/Basson":                    ("Basson",              ORD),
    "Bassons/Contrebasson":              ("Contrebasson",        ORD),
    "Clarinets/eb_clarinet_boehm_system":   ("Clarinette en Mib",      ORD),
    "Clarinets/bb_clarinet_boehm_system":   ("Clarinette en Sib",      ORD),
    "Clarinets/bass_clarinet_boehm_system": ("Clarinette basse",       ORD),
    "Clarinets/contrabass_clarinet":        ("Clarinette contrebasse", ORD),
    "horns/horn_in_f": ("Cor", {
        "ordinario": ["ordinario"],
        "straight":  ["straight_mute"],              # la sourdine du cor dans l'étude
    }),
    "trumpets/trumpet_in_c": ("Trompette", {
        "ordinario":  ["ordinario"],
        "straight":   ["straight", "straight_metal", "straight_wood"],
        "cup":        ["cup", "cup_metal", "cup_open"],
        "harmon":     ["harmon"],
        "wah ouvert": ["wahwah_stem_in_open"],
        "wah fermé":  ["wahwah_stem_in_closed"],
    }),
    "trombones/tenor_trombone": ("Trombone ténor", {
        "ordinario":  ["ordinario"],
        "straight":   ["straight", "straight_metal", "straight_wood"],
        "cup":        ["cup", "cup_metal", "cup_open"],
        "harmon":     ["harmon"],
        "wah ouvert": ["wahwah_stem_in_open"],
        "wah fermé":  ["wahwah_stem_in_closed"],
    }),
    "trombones/bass_trombone": ("Trombone basse", ORD),
    "tubas/bass_tuba": ("Tuba basse", {
        "ordinario": ["ordinario"],
        "straight":  ["straight", "straight_metal", "straight_wood"],
    }),
    "violin": ("Violon", {
        "ordinario": ["ordinario"],
        "sourdine":  ["con_sordino"],
        "piombo":    ["practice_mute"],              # = sordina_piombo de l'étude
    }),
    "viola": ("Alto", {
        "ordinario": ["ordinario"],
        "sourdine":  ["con_sordino"],
        "piombo":    ["practice_mute"],
    }),
    "violoncello": ("Violoncelle", {
        "ordinario": ["ordinario"],
        "sourdine":  ["con_sordino_rubber"],         # la sourdine « normale » au violoncelle
        "piombo":    ["practice_mute_metal"],
    }),
    "double_bass": ("Contrebasse", {
        "ordinario": ["ordinario"],
        "sourdine":  ["con_sordino"],
        "piombo":    ["practice_mute"],
    }),
}

SUFFIXE = {"ordinario": "", "straight": "+sourdine sèche", "cup": "+cup",
           "harmon": "+harmon", "wah ouvert": "+wah ouvert", "wah fermé": "+wah fermé",
           "sourdine": "+sourdine", "piombo": "+piombo"}


def inventaire(racine):
    lignes = []
    for dossier, (nom, cats) in CONFIG.items():
        d = os.path.join(racine, dossier)
        if not os.path.isdir(d):
            lignes.append((nom, dossier, "DOSSIER ABSENT", 0, {})); continue
        for cat, techs in cats.items():
            fich = FL.lister_par_nom(d, set(techs))
            cordes = collections.Counter(c for _, _, c in fich if c)
            lignes.append((nom + SUFFIXE[cat], dossier, cat, len(fich), dict(cordes)))
    return lignes


def mesurer_config(racine, dossier, nom, cat, techs, sortie, fiche, par_corde=False):
    d = os.path.join(racine, dossier)
    fich = FL.lister_par_nom(d, set(techs))
    if len(fich) < 5:
        print(f"  {nom}{SUFFIXE[cat]} : {len(fich)} fichier(s) — ignoré"); return None
    etiquette = (nom + SUFFIXE[cat]).replace("/", "-")
    os.makedirs(sortie, exist_ok=True)
    groupes = {"": [f for f, t, c in fich]}
    if par_corde:
        cs = sorted({c for _, _, c in fich if c})
        if len(cs) >= 2:
            groupes = {f" corde {c}": [f for f, t, cc in fich if cc == c] for c in cs}
            groupes[""] = [f for f, t, c in fich]
    sorties = {}
    for suff, liste in groupes.items():
        if len(liste) < 5:
            continue
        et = etiquette + suff
        sp = os.path.join(sortie, f"{et}.spectres.txt")
        ok, ko, tot = FL.ecrire_spectres_liste(liste, sp, racine)
        res = FL.mesurer(sp)
        if 'erreur' in res:
            print(f"  {et} : {res['erreur']}"); continue
        nom_ctrl = res['nomenclature']
        alerte = ""
        if nom_ctrl is None or abs(nom_ctrl - 1.0) > 0.03:
            alerte = f"  !! NOMENCLATURE {nom_ctrl} (attendu ~1,00)"
        res['_config'] = dict(corpus="ZSamples", dossier=dossier, nom=nom,
                              categorie=cat, techniques=techs, corde=suff.strip() or None)
        with open(os.path.join(sortie, f"{et}.json"), "w") as fh:
            json.dump(res, fh, ensure_ascii=False)
        v = "  ".join(f"{x['hz']:.0f}:{x['verdict']}" for x in res['formants']) or "aucun candidat"
        print(f"  {et:34s} n={res['n']:4d} ρ={res['rho']} pente={res['pente']} | {v}{alerte}")
        if fiche and not suff:
            with open(os.path.join(sortie, f"{et}.md"), "w") as fh:
                fh.write(fiche_markdown(et, res, f"ZSamples/{dossier}", "+".join(techs)))
        sorties[suff.strip() or "global"] = res
    return sorties


def test_cordes(sorties, nom):
    """Les formants coïncident-ils d'une corde à l'autre ? Tolérance 1/6 d'octave."""
    import math
    cordes = {k: v for k, v in sorties.items() if k.startswith("corde")}
    if len(cordes) < 2:
        return
    print(f"    test inter-cordes ({nom}) :")
    for k, v in cordes.items():
        hz = [x['hz'] for x in v['formants'] if x['verdict'] == 'retenu-cr']
        print(f"      {k:10s} n={v['n']:3d}  retenus : {[round(h) for h in hz] or '—'}")
    g = [x['hz'] for x in sorties.get('global', {}).get('formants', [])
         if x['verdict'] == 'retenu-cr']
    for h in g:
        pres = [k for k, v in cordes.items()
                if any(abs(math.log2(x['hz'] / h)) <= 1/6 for x in v['formants']
                       if x['verdict'] in ('retenu-cr', 'fragile-cr'))]
        print(f"      {h:7.0f} Hz retrouvé sur {len(pres)}/{len(cordes)} cordes "
              f"{'— COHÉRENT' if len(pres) >= max(2, len(cordes)-1) else '— à examiner'}")


def main():
    ap = argparse.ArgumentParser(description="Mesure formantique du corpus ZSamples.")
    ap.add_argument("--racine", required=True)
    ap.add_argument("--sortie", default="resultats_zsamples")
    ap.add_argument("--fiche", action="store_true")
    ap.add_argument("--lister", action="store_true", help="inventaire seul, sans mesure")
    ap.add_argument("--seulement", help="noms d'instruments, séparés par des virgules")
    ap.add_argument("--test-cordes", action="store_true",
                    help="mesurer chaque corde séparément et comparer")
    a = ap.parse_args()

    if a.lister:
        tot = 0
        print(f"{'configuration':40s} {'cat.':12s}  n   cordes")
        for nom, dossier, cat, n, cordes in inventaire(a.racine):
            print(f"{nom:40s} {cat:12s} {n:4d}  {cordes if cordes else ''}")
            tot += n
        print(f"\n{tot} fichiers retenus par la correspondance.")
        return

    filtre = {s.strip() for s in a.seulement.split(",")} if a.seulement else None
    for dossier, (nom, cats) in CONFIG.items():
        if filtre and nom not in filtre:
            continue
        if not os.path.isdir(os.path.join(a.racine, dossier)):
            print(f"{nom} : dossier absent — ignoré"); continue
        print(nom)
        for cat, techs in cats.items():
            s = mesurer_config(a.racine, dossier, nom, cat, techs, a.sortie, a.fiche,
                               par_corde=a.test_cordes)
            if s and a.test_cordes:
                test_cordes(s, f"{nom}{SUFFIXE[cat]}")


if __name__ == "__main__":
    main()
