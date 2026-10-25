#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
carte_vocalique.py — méthode 12 : carte vocalique par registre et par dynamique.

Répond à une question DIFFÉRENTE de celle des formants (méthode 8). Un formant est une
propriété du filtre, indépendante de la note. Une voyelle, non : quand on dit qu'un basson
sonne /o/, on constate que son spectre RAYONNÉ a son énergie là où /o/ a la sienne. Les
voyelles de Calliope sont elles aussi mesurées sur le son, pas sur le conduit vocal.

D'où une méthode qui ne sépare jamais source et filtre, donc robuste pour TOUT instrument,
y compris ceux dont la méthode 8 ne peut rien tirer (flûtes).

Chaîne
------
1. spectre moyen (attaque exclue) -> puissance
2. motif d'excitation auditive sur échelle de Bark, fonction d'étalement de Schroeder
3. maxima locaux, puis FUSION des maxima distants de moins de 3,5 Bark en leur centre de
   gravité énergétique — c'est l'effet de centre de gravité de Chistovich, le même seuil que
   celui des graphiques Formants-Francais
4. F1 = maximum survivant le plus bas, F2 = le suivant
5. agrégation par registre × dynamique : médiane et dispersion interquartile
6. voyelle française la plus proche dans le plan (F1, F2) en Bark

Garde-fous
----------
La carte est marquée, cellule par cellule :
  · « 1 max »   un seul maximum survit : il n'y a pas de couple (F1, F2) à lire
  · « F1=f0 »   F1 coïncide avec la fondamentale : la « voyelle » n'est que la hauteur
  · « dispersé » l'ellipse dépasse 2 Bark sur un axe : le registre n'a pas de couleur stable
Une cellule marquée ne doit pas être publiée comme une voyelle.
"""
import os, sys, re, math, json
import numpy as np
import formants_lib as FL

TIRET = "\u2014"

# ───────────────────────── échelle de Bark ──────────────────────────────────
def hz_vers_bark(f):
    f = np.asarray(f, dtype=float)
    return 13.0 * np.arctan(0.00076 * f) + 3.5 * np.arctan((f / 7500.0) ** 2)

_ZG = np.arange(0.0, 24.01, 0.05)
_FG = np.logspace(math.log10(20), math.log10(16000), 4000)
_ZF = hz_vers_bark(_FG)

def bark_vers_hz(z):
    return float(np.interp(z, _ZF, _FG))

def etalement_schroeder(dz):
    """Fonction d'étalement de Schroeder (1979), en dB. Standard en codage perceptif,
    indépendante du niveau — choix assumé pour rester reproductible."""
    return 15.81 + 7.5 * (dz + 0.474) - 17.5 * np.sqrt(1.0 + (dz + 0.474) ** 2)

# pré-calcul de la matrice d'étalement bins -> grille de Bark
_FBINS = np.arange(FL.NBINS) * FL.BW
_FBINS[0] = 1.0
_ZBINS = hz_vers_bark(_FBINS)
_MASQUE = (_FBINS >= 50.0) & (_FBINS <= 10500.0)
_SPREAD = 10.0 ** (etalement_schroeder(_ZG[:, None] - _ZBINS[None, :]) / 10.0)
_SPREAD[:, ~_MASQUE] = 0.0

def motif_excitation(spectre):
    """Motif d'excitation auditive, en dB, sur la grille de Bark."""
    p = np.asarray(spectre, dtype=float) ** 2
    e = _SPREAD @ p
    return 10.0 * np.log10(np.maximum(e, 1e-20))

FUSION_BARK = 3.5
MIN_N_CELLULE = 5        # une cellule sous ce seuil n'est pas publiée
SEUIL_ASSIGNABLE = 2.0   # distance en Bark : voyelle assignable
SEUIL_HORS = 3.5         # au-delà (seuil de Chistovich) : hors espace vocalique

def statut_distance(d):
    if d is None:            return "—"
    if d <= SEUIL_ASSIGNABLE: return "assignable"
    if d <= SEUIL_HORS:       return "approximative"
    return "HORS ESPACE"

def maxima_fusionnes(exc, seuil_rel=25.0):
    """Maxima locaux du motif, puis fusion de ceux distants de < 3,5 Bark en leur
    centre de gravité énergétique (effet Chistovich). Rend [(z, niveau_dB), ...]."""
    pics = [i for i in range(1, len(exc) - 1)
            if exc[i] >= exc[i - 1] and exc[i] > exc[i + 1]]
    if not pics:
        return []
    plancher = exc.max() - seuil_rel
    pics = [i for i in pics if exc[i] >= plancher]
    if not pics:
        return []
    groupes, courant = [], [pics[0]]
    for i in pics[1:]:
        if _ZG[i] - _ZG[courant[-1]] < FUSION_BARK:
            courant.append(i)
        else:
            groupes.append(courant); courant = [i]
    groupes.append(courant)
    out = []
    for g in groupes:
        lin = 10.0 ** (exc[np.array(g)] / 10.0)
        z = float(np.sum(_ZG[np.array(g)] * lin) / np.sum(lin))   # centre de gravité
        out.append((z, float(max(exc[i] for i in g))))
    return out

def f1_f2(spectre, f0):
    """Rend (z1, z2, hz1, hz2, marques)."""
    exc = motif_excitation(spectre)
    m = maxima_fusionnes(exc)
    marques = []
    if not m:
        return (None,) * 4 + ([("aucun max")],)
    z1 = m[0][0]; h1 = bark_vers_hz(z1)
    if len(m) < 2:
        return z1, None, h1, None, ["1 max"]
    z2 = m[1][0]; h2 = bark_vers_hz(z2)
    if abs(12 * math.log2(max(h1, 1e-6) / f0)) < 1.0:
        marques.append("F1=f0")
    return z1, z2, h1, h2, marques

# ───────────────────── registres : lecture de registres.md ──────────────────
_NOTE_RE = re.compile(r'\b([A-G]#?)(-?\d{1,2})\b')
_SEMI = {'C': 0, 'D': 2, 'E': 4, 'F': 5, 'G': 7, 'A': 9, 'B': 11}

def _midi(txt):
    m = _NOTE_RE.search(txt)
    if not m:
        return None
    return 12 * (int(m.group(2)) + 1) + _SEMI[m.group(1)[0]] + (1 if '#' in m.group(1) else 0)

def lire_registres(chemin):
    """Parse Data/registres.md. Rend {instrument: [(label, midi_bas, midi_haut), ...]}.
    Tolère « à partir de X », « au-dessus », « en dessous de X » et les parenthèses."""
    out = {}
    for ligne in open(chemin, encoding='utf-8'):
        if not ligne.startswith('* ') or ' : ' not in ligne:
            continue
        nom, reste = ligne[2:].split(' : ', 1)
        nom = nom.strip().lower()
        segs = [s.strip() for s in reste.split(',')]
        reg, dernier_haut = [], None
        for s in segs:
            s_net = re.sub(r'\([^)]*\)', '', s).strip()
            label = re.sub(r'^registre\s+', '', s_net)
            label = re.split(r'\s+[A-G]#?-?\d|\s+à partir|\s+au-?dessus|\s+en dessous', label)[0].strip()
            notes = _NOTE_RE.findall(s_net)
            if 'à partir' in s_net or 'au-dessus' in s_net or 'au dessus' in s_net:
                bas = _midi(s_net) if notes else (dernier_haut if dernier_haut else 0)
                reg.append((label or 'suraigu', bas, 127)); continue
            if 'en dessous' in s_net:
                h = _midi(s_net) or 127
                reg.append((label or 'extrême grave', 0, h)); continue
            if len(notes) >= 2:
                a = 12*(int(notes[0][1])+1)+_SEMI[notes[0][0][0]]+(1 if '#' in notes[0][0] else 0)
                b = 12*(int(notes[1][1])+1)+_SEMI[notes[1][0][0]]+(1 if '#' in notes[1][0] else 0)
                reg.append((label, min(a, b), max(a, b))); dernier_haut = max(a, b)
        if reg:
            out[nom] = reg
    return out

def registre_de(regs, midi):
    for label, lo, hi in regs:
        if lo <= midi <= hi:
            return label
    return "hors registre"

DYN = {'ppp': 'doux', 'pp': 'doux', 'p': 'doux', 'mp': 'doux',
       'mf': 'moyen',
       'f': 'fort', 'ff': 'fort', 'fff': 'fort'}

# ───────────────────── voyelles du français (Calliope, homme) ───────────────
VOYELLES = [("i", 308, 2064), ("y", 300, 1750), ("u", 315, 764),
            ("e", 365, 1961), ("ø", 381, 1417), ("o", 383, 793),
            ("ɛ", 530, 1718), ("œ", 517, 1391), ("ɔ", 531, 998),
            ("a", 684, 1256)]
VOY_BARK = [(v, float(hz_vers_bark(f1)), float(hz_vers_bark(f2))) for v, f1, f2 in VOYELLES]

def voyelle_proche(z1, z2):
    if z1 is None or z2 is None:
        return None, None
    d = [(math.hypot(z1 - a, z2 - b), v) for v, a, b in VOY_BARK]
    d.sort()
    return d[0][1], round(d[0][0], 2)

# ───────────────────────────── agrégation ──────────────────────────────────
def carte(chemin_spectres, nom_registres, chemin_registres, forcer_grille=None):
    """forcer_grille : None = choix automatique, "rd" = registre x dynamique,
    "r" = registre seul. Le repli est automatique si n/12 < MIN_N_CELLULE : en dessous
    d'environ 80 echantillons, les 12 cellules ne se remplissent pas. On ne baisse PAS
    le seuil de 5 par cellule — on reduit la finesse de la grille."""
    regs = lire_registres(chemin_registres)
    cle = nom_registres if nom_registres in regs else None
    if cle is None:    # tolerance : « violon » trouve « violon et ensemble de violons »
        cands = [k for k in regs if k == nom_registres or k.startswith(nom_registres + " ")
                 or nom_registres in k.split(" et ")]
        if len(cands) == 1:
            cle = cands[0]
        else:
            raise SystemExit(f"registre inconnu ou ambigu : {nom_registres} "
                             f"(candidats {cands}; connus {sorted(regs)})")
    R = regs[cle]
    ech = FL.lire_spectres(chemin_spectres)
    n_reg = len({registre_de(R, m) for f0, m, d, sp in ech})
    grille = forcer_grille or ("rd" if len(ech) / max(1, n_reg * 3) >= MIN_N_CELLULE else "r")
    cells = {}
    for f0, midi, dyn, sp in ech:
        z1, z2, h1, h2, marq = f1_f2(sp, f0)
        if z1 is None:
            continue
        dy = DYN.get((dyn or '').lower(), 'autre') if grille == "rd" else "toutes"
        cells.setdefault((registre_de(R, midi), dy), []).append((z1, z2, h1, h2, marq, midi))
    ordre = [l for l, a, b in R]
    res = []
    for (reg, dy), v in sorted(cells.items(),
                               key=lambda kv: (ordre.index(kv[0][0]) if kv[0][0] in ordre else 99,
                                               ['doux', 'moyen', 'fort', 'toutes', 'autre'].index(kv[0][1]))):
        if len(v) < MIN_N_CELLULE:
            res.append(dict(registre=reg, dynamique=dy, n=len(v), F1_bark=None, F2_bark=None,
                            F1_hz=None, F2_hz=None, IQR1=None, IQR2=None, voyelle=None,
                            distance_bark=None, statut="non publiable",
                            marques=[f"n<{MIN_N_CELLULE}"]))
            continue
        z1 = np.array([x[0] for x in v])
        ok2 = [x for x in v if x[1] is not None]
        n1max = len(v) - len(ok2)
        nf0 = sum(1 for x in v if "F1=f0" in x[4])
        if len(ok2) < len(v) / 2:      # la majorite n'a qu'un maximum : pas de F2 lisible
            ok2 = []
        if ok2:
            z2 = np.array([x[1] for x in ok2])
            m1, m2 = float(np.median(z1)), float(np.median(z2))
            i1 = float(np.percentile(z1, 75) - np.percentile(z1, 25))
            i2 = float(np.percentile(z2, 75) - np.percentile(z2, 25))
            voy, dist = voyelle_proche(m1, m2)
        else:
            m1 = float(np.median(z1)); m2 = i2 = None
            i1 = float(np.percentile(z1, 75) - np.percentile(z1, 25))
            voy = dist = None
        marques = []
        if n1max > len(v) / 2: marques.append("1 max")
        if nf0 > len(v) / 2:   marques.append("F1=f0")
        if i1 > 2.0 or (i2 is not None and i2 > 2.0): marques.append("dispersé")
        res.append(dict(registre=reg, dynamique=dy, n=len(v), statut=statut_distance(dist),
                        F1_bark=round(m1, 2), F2_bark=None if m2 is None else round(m2, 2),
                        F1_hz=round(bark_vers_hz(m1)), F2_hz=None if m2 is None else round(bark_vers_hz(m2)),
                        IQR1=round(i1, 2), IQR2=None if i2 is None else round(i2, 2),
                        voyelle=voy, distance_bark=dist, marques=marques))
    return dict(grille=grille, cellules=res)

INSTRUMENTS = {                 # nom affiché : (fichier de spectres, clé dans registres.md)
    "Petite flûte":           ("Piccolo",                 "piccolo"),
    "Flûte":                  ("Flute",                   "flûte"),
    "Flûte basse":            ("Bass-Flute",              "flûte basse"),
    "Flûte contrebasse":      ("Contrabass-Flute",        "flûte contrebasse"),
    "Hautbois":               ("Oboe",                    "hautbois"),
    # Cor anglais : ABSENT de registres.md — registres à définir avant mesure
    "Basson":                 ("Bassoon",                 "basson"),
    "Contrebasson":           ("Contrebasson",            "contrebasson"),
    "Clarinette en Mib":      ("Clarinet-Eb",             "clarinette en mib"),
    "Clarinette en Sib":      ("Clarinet_Bb",             "clarinette"),
    "Clarinette basse":       ("Bass-Clarinet-Bb",        "clarinette basse"),
    "Clarinette contrebasse": ("Contrabass-Clarinet-Bb",  "clarinette contrebasse"),
    "Cor":                    ("Horn",                    "cor"),
    "Trompette":              ("Trumpet_C",               "trompette"),
    "Trombone ténor":         ("Trombone",                "trombones"),
    "Trombone basse":         ("Bass-Trombone",           "trombones"),
    "Tuba basse":             ("Bass_Tuba",               "tubas"),
    "Tuba contrebasse":       ("Contrabass-Tuba",         "tubas"),
    "Violon":                 ("Violin",                  "violon"),
    "Alto":                   ("Viola",                   "alto"),
    "Violoncelle":            ("Violoncello",             "violoncelle"),
    "Contrebasse":            ("Contrabass",              "contrebasse"),
}

if __name__ == "__main__":
    base = sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser("~/work/out")
    chreg = sys.argv[2] if len(sys.argv) > 2 else "../../Data/registres.md"
    lot = None
    for a in sys.argv[3:]:
        if a.startswith("--seulement="):
            lot = {x.strip() for x in a.split("=",1)[1].split(",")}
    sortie = "carte_vocalique.json"
    motif = "{fich}.txt"      # nomenclature SOL ; ZSamples : "{nom}.spectres.txt"
    for a in sys.argv[3:]:
        if a.startswith("--sortie="): sortie = a.split("=",1)[1]
        if a.startswith("--motif="):  motif  = a.split("=",1)[1]
    tout = {}
    for nom, (fich, cle) in INSTRUMENTS.items():
        if lot and nom not in lot: continue
        p = os.path.join(base, motif.format(fich=fich, nom=nom))
        if not os.path.exists(p):
            print(f"{nom} : {p} absent"); continue
        out = carte(p, cle, chreg)
        r = out["cellules"]; tout[nom] = out
        gl = "registre x dynamique" if out["grille"] == "rd" else "REGISTRE SEUL (repli, corpus restreint)"
        print("\n=== {}   grille : {}".format(nom, gl))
        print("{:14s} {:6s} {:>4s}  {:>13s} {:>13s}  {:>9s}  {:7s} {:>5s}  {:14s} {}".format(
            "registre", "dyn", "n", "F1", "F2", "IQR", "voyelle", "dist", "statut", "marques"))
        for c in r:
            marq = ",".join(c["marques"])
            if c["F1_bark"] is None:
                print("{:14s} {:6s} {:4d}  non publiable ({})".format(
                    c["registre"], c["dynamique"], c["n"], marq))
                continue
            f2 = ("{:5.2f}B {:5d}Hz".format(c["F2_bark"], c["F2_hz"])
                  if c["F2_bark"] is not None else "      " + TIRET + "      ")
            iq = ("{:.1f}/{:.1f}".format(c["IQR1"], c["IQR2"]) if c["IQR2"] is not None
                  else "{:.1f}/".format(c["IQR1"]) + TIRET)
            d  = "" if c["distance_bark"] is None else "{:.2f}".format(c["distance_bark"])
            print("{:14s} {:6s} {:4d}  {:5.2f}B {:5d}Hz {}  {:>9s}  {:7s} {:>5s}  {:14s} {}".format(
                c["registre"], c["dynamique"], c["n"], c["F1_bark"], c["F1_hz"], f2, iq,
                c["voyelle"] or TIRET, d, c["statut"], marq))

    if os.path.exists(sortie):
        try: tout = {**json.load(open(sortie)), **tout}
        except Exception: pass
    json.dump(tout, open(sortie, "w"), ensure_ascii=False, indent=1)
    print(f"\n{sortie} écrit ({len(tout)} instrument(s) cumulés)")
