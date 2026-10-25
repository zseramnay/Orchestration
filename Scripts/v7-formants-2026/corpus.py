#!/usr/bin/env python3
"""Pilote générique pour verser un NOUVEAU CORPUS dans l'étude formantique.

Problème résolu
---------------
Chaque banque arrive avec sa propre arborescence et sa propre nomenclature.
Constaté sur cinq banques :

  SOL2020    Instrument/technique/inst-technique-NOTE-dyn-N-N.wav
  ZSamples   instrument/inst-technique_string_N-NOTE-dyn-n.wav
  Xsamples   instrument/DYNAMIQUE/code_note.aif        note en MINUSCULES
  VO         instrument/sourdine/ordinario/String-N/code-NOTE-dyn.AIF
  VN         Instrument/TAG_dyn/TAG_dyn_NOTE.wav       dynamique AVANT la note

Principe
--------
On ne touche PAS à la chaîne de mesure. `formants_lib.lire_spectres()` relit la
nomenclature depuis le chemin écrit dans le fichier de spectres. Il suffit donc
d'y écrire un **nom normalisé au format SOL** :

    /Instrument/technique/instrument-technique-NOTE-dyn-N-N.wav

et tout l'aval — mesurer(), les tests de rejet, carte_vocalique.py — fonctionne
sans modification. Adapter une banque = écrire un descripteur JSON, pas du code.
Vérifié : sur ZSamples, retrouve à l'identique le violon
[2 216 · 2 792 · 3 948 · 5 583] et le basson [494 · 1 108 · 1 974].

Usage
-----
    python3 corpus.py --decrire "/chemin/vers/Banque" > /tmp/d.json
    python3 corpus.py --valider /tmp/d.json          # obligatoire
    python3 corpus.py --mesurer /tmp/d.json [--seulement Violon,Alto]

Descripteur
-----------
    {
      "corpus": "VO",
      "racine": "/Users/.../VO",
      "sortie": "Etude-Formants/Instruments/VO",
      "extensions": [".AIF"],
      "separateurs": ["-"],
      "dyn_alias": {"forte": "f", "piano": "p", "mezzo": "mf"},
      "exclure": ["NoPitch"],
      "configurations": [
        {"nom": "Violon", "technique": "ordinario",
         "dossiers": ["violin/senza-sordino/ordinario/String-I", "…II", "…III", "…IV"],
         "dyn_depuis": "nom",
         "corde_depuis": "dossier"}
      ]
    }

Champs d'une configuration
    nom           nom d'instrument du projet (sert de nom de fichier de sortie)
    technique     technique imposée ; à défaut, lue dans le nom de fichier
    dossiers      sous-dossiers à FUSIONNER (chemins relatifs à la racine)
    dyn_depuis    "nom" (défaut) | "dossier" — nom du dossier immédiat
    corde_depuis  null (défaut) | "dossier" | "nom" — active le test inter-cordes
"""

import argparse, json, os, re, sys, glob, collections

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import formants_lib as FL

DYN_CONNUES = ('pppp', 'ppp', 'pp', 'p', 'mp', 'mf', 'ff', 'fff', 'ffff', 'f')
_RX_CORDE = re.compile(r'\bstring[-_ ]?([IVX]+)\b', re.I)


# ───────────────────────────── nomenclature ─────────────────────────────────

def note_souple(jeton):
    """Comme note_vers_f0, mais tolère trois écarts de notation observés :

    - **minuscules** — Xsamples nomme ses fichiers `fagf_a#1.aif` ; sans cela,
      aucune note n'est lue et la banque entière est illisible ;
    - **H allemand** = si naturel — le cor et le trombone de Xsamples écrivent
      `Corf_h2`, les autres instruments de la même banque écrivent `b2`. Vérifié
      mutuellement exclusif par dossier : aucune ambiguïté, `b` reste l'anglais ;
    - **bémols**, réécrits en dièse de la note inférieure.
    """
    if not jeton:
        return None
    j = jeton.strip()
    m = re.match(r'^([A-Ha-h])([#b]?)(-?\d{1,2})$', j)
    if not m:
        return None
    lettre, alt, oct_ = m.group(1).upper(), m.group(2), m.group(3)
    if lettre == 'H':                    # notation allemande : H = si naturel
        lettre = 'B'
    if alt == 'b':                       # réécrit le bémol en dièse de la note inférieure
        voisin = {'C': 'B', 'D': 'C', 'E': 'D', 'F': 'E', 'G': 'F', 'A': 'G', 'B': 'A'}
        bas = voisin[lettre]
        oct_ = str(int(oct_) - 1) if lettre == 'C' else oct_
        jeton_norm = bas + ('#' if bas not in ('E', 'B') else '') + oct_
    else:
        jeton_norm = lettre + alt + oct_
    r = FL.note_vers_f0(jeton_norm)
    return (r[0], r[1], jeton_norm) if r else None


def _jetons(base, separateurs):
    return re.split('[' + re.escape(''.join(separateurs)) + ']', base)


def parser_nom(base, separateurs, dyn_alias):
    """Balaye les jetons. La note est le premier jeton qui en est une ; la
    dynamique est cherchée APRÈS puis AVANT — VN écrit `..._mp_G#4`, SOL écrit
    `...-G#4-mp-N-N`. Sans la recherche en arrière, toute la banque VN perd ses
    dynamiques et la grille registre × dynamique s'effondre en registre seul."""
    toks = _jetons(base, separateurs)
    for i, t in enumerate(toks):
        n = note_souple(t)
        if not n:
            continue
        dyn = None
        for u in toks[i + 1:]:            # après (SOL, VO, ZSamples)
            v = dyn_alias.get(u.lower(), u.lower())
            if v in DYN_CONNUES:
                dyn = v
                break
        if dyn is None:                   # avant (VN)
            for u in reversed(toks[:i]):
                v = dyn_alias.get(u.lower(), u.lower())
                if v in DYN_CONNUES:
                    dyn = v
                    break
        return dict(note=n[2], f0=n[0], dyn=dyn,
                    technique='-'.join(toks[1:i]) if i > 1 else '')
    return None


def dyn_de_dossier(chemin, dyn_alias):
    """Xsamples range par dynamique : bassoon/forte/, bassoon/mf/, bassoon/piano/."""
    d = os.path.basename(os.path.dirname(chemin)).lower()
    v = dyn_alias.get(d, d)
    return v if v in DYN_CONNUES else None


def corde_de(chemin, info, depuis):
    if depuis == 'dossier':
        m = _RX_CORDE.search(chemin)
    elif depuis == 'nom':
        m = _RX_CORDE.search(os.path.basename(chemin))
    else:
        return None
    return m.group(1).upper() if m else None


def nom_normalise(instrument, technique, note, dyn, corde=None):
    """Reconstruit un nom au format SOL, relu sans erreur par analyser_nom().

    Garde-fous : ni '-' ni espace dans l'instrument et la technique (le tiret
    sert de séparateur), et aucun jeton ne doit ressembler à une note, sinon
    analyser_nom() s'arrêterait dessus. La corde est écrite en suffixe
    `_string_N` pour que separer_corde() la retrouve."""
    def propre(s, defaut):
        s = re.sub(r'[^0-9A-Za-z_]+', '_', (s or '').strip()).strip('_') or defaut
        return 'x' + s if note_souple(s) else s
    i = propre(instrument, 'inst')
    t = propre(technique, 'ordinario')
    if corde:
        t = f'{t}_string_{corde}'
    return f"/{i}/{t}/{i}-{t}-{note}-{dyn or 'mf'}-N-N.wav"


# ───────────────────────────── collecte ─────────────────────────────────────

def fichiers_sous(racine, sous, extensions):
    base = os.path.join(racine, sous) if sous else racine
    out = []
    for e in extensions:
        out += glob.glob(os.path.join(base, '**', '*' + e), recursive=True)
    return sorted(set(out))


def collecter(desc, cfg):
    """Rend (paires, rejets) où paires = [(chemin_reel, nom_normalise), …]."""
    exts = desc.get('extensions', ['.wav'])
    seps = desc.get('separateurs', ['-'])
    alias = {k.lower(): v for k, v in (desc.get('dyn_alias') or {}).items()}
    exclure = [e.lower() for e in desc.get('exclure', [])]
    tech_impose = cfg.get('technique')
    dyn_depuis = cfg.get('dyn_depuis', 'nom')
    corde_depuis = cfg.get('corde_depuis')

    paires, rejets, exemples = [], collections.Counter(), []
    for d in cfg.get('dossiers') or ['']:
        trouve = fichiers_sous(desc['racine'], d, exts)
        if not trouve:
            rejets[f'dossier vide: {d}'] += 1
        for f in trouve:
            if any(e in f.lower() for e in exclure):
                rejets['exclu'] += 1
                continue
            base = os.path.splitext(os.path.basename(f))[0]
            info = parser_nom(base, seps, alias)
            if info is None:
                rejets['note illisible'] += 1
                if len(exemples) < 3:
                    exemples.append(base)
                continue
            dyn = (dyn_de_dossier(f, alias) if dyn_depuis == 'dossier' else info['dyn'])
            if dyn is None:
                dyn = info['dyn'] or dyn_de_dossier(f, alias)
            tech = tech_impose or info['technique'] or 'ordinario'
            corde = corde_de(f, info, corde_depuis)
            paires.append((f, nom_normalise(cfg['nom'], tech, info['note'], dyn, corde)))
    if exemples:
        rejets[f'exemples illisibles: {exemples}'] = 1
    return paires, rejets


# ───────────────────────────── mode --decrire ───────────────────────────────

def decrire(racine):
    exts = ['.wav', '.aif', '.aiff', '.AIF', '.AIFF', '.flac']
    tous = []
    for e in exts:
        tous += glob.glob(os.path.join(racine, '**', '*' + e), recursive=True)
    tous = sorted(set(tous))
    E = sys.stderr
    print(f"# {len(tous)} fichiers audio sous {racine}", file=E)
    if not tous:
        print("# AUCUN fichier audio. Vérifier le chemin — attention aux espaces "
              "en fin de nom de dossier.", file=E)
        return
    print(f"# extensions : {dict(collections.Counter(os.path.splitext(f)[1] for f in tous))}", file=E)

    par_dossier = collections.Counter(os.path.relpath(os.path.dirname(f), racine) for f in tous)
    print("\n# Dossiers contenant de l'audio :", file=E)
    prem = {}
    for f in tous:
        prem.setdefault(os.path.relpath(os.path.dirname(f), racine), os.path.basename(f))
    for d, n in sorted(par_dossier.items()):
        print(f"#   {n:6d}  {d[:58]:58s} ex: {prem[d][:32]}", file=E)

    print("\n# Lisibilité de la nomenclature :", file=E)
    ech = tous[::max(1, len(tous) // 400)][:400]
    best = None
    for seps in (['-'], ['_'], ['-', '_'], ['-', '_', '.', ' ']):
        ok = dy = 0
        for f in ech:
            i = parser_nom(os.path.splitext(os.path.basename(f))[0], seps, {})
            if i:
                ok += 1
                dy += bool(i['dyn'])
        print(f"#   separateurs={seps!r:26s} note {ok:4d}/{len(ech)}   dynamique {dy:4d}/{len(ech)}", file=E)
        if best is None or ok > best[1]:
            best = (seps, ok)
    seps = best[0]
    print(f"\n# Exemples lus avec separateurs={seps!r} :", file=E)
    for f in ech[:10]:
        b = os.path.splitext(os.path.basename(f))[0]
        print(f"#   {b[:58]:58s} -> {parser_nom(b, seps, {}) or 'ILLISIBLE'}", file=E)

    dyns = collections.Counter()
    for f in ech:
        i = parser_nom(os.path.splitext(os.path.basename(f))[0], seps, {})
        if i and i['dyn']:
            dyns[i['dyn']] += 1
    print(f"\n# Dynamiques lues dans les noms : {dict(dyns)}", file=E)
    dd = collections.Counter(os.path.basename(d).lower() for d in par_dossier)
    print(f"# Noms de dossiers terminaux : {sorted(dd)[:20]}", file=E)
    print("# Si ce sont des dynamiques (forte/piano/mf), mettre dyn_depuis=\"dossier\""
          "\n# et compléter dyn_alias.", file=E)

    feuilles = sorted(d for d in par_dossier if par_dossier[d] >= 5)
    desc = dict(corpus=os.path.basename(os.path.normpath(racine)).strip(),
                racine=os.path.abspath(racine),
                sortie="Etude-Formants/Instruments/"
                       + os.path.basename(os.path.normpath(racine)).strip(),
                extensions=sorted({os.path.splitext(f)[1] for f in tous}),
                separateurs=seps,
                dyn_alias={"forte": "f", "piano": "p"},
                exclure=["NoPitch"],
                configurations=[dict(nom="A_RENOMMER_" + re.sub(r'[^0-9A-Za-z]+', '_', d),
                                     technique="ordinario", dossiers=[d],
                                     dyn_depuis="nom")
                                for d in feuilles])
    print("\n# Descripteur proposé — RENOMMER les configurations avec les noms"
          "\n# d'instruments du projet et FUSIONNER les dossiers d'un même instrument.", file=E)
    json.dump(desc, sys.stdout, ensure_ascii=False, indent=2)
    print()


# ───────────────────────────── mode --valider ───────────────────────────────

def valider(desc, seulement=None, n_controle=25):
    """Contrôle le descripteur sans mesurer. Vérifie surtout la NOMENCLATURE :
    un corpus en hauteurs écrites produirait des formants transposés."""
    souci = 0
    for cfg in desc['configurations']:
        if seulement and cfg['nom'] not in seulement:
            continue
        paires, rejets = collecter(desc, cfg)
        if not paires:
            print(f"✗ {cfg['nom']:26s} AUCUN fichier   {dict(rejets)}")
            souci += 1
            continue
        notes = [note_souple(p[1].split('-')[-4])[0] for p in paires]
        dyns = collections.Counter(p[1].split('-')[-3] for p in paires)
        cordes = collections.Counter(FL.separer_corde(p[1].split('/')[2])[1] for p in paires)
        print(f"  {cfg['nom']:26s} {len(paires):5d} fich. {len({round(f,1) for f in notes}):3d} notes  "
              f"f0 {min(notes):7.1f}–{max(notes):7.1f}  dyn={dict(dyns)}"
              + (f"  cordes={dict(cordes)}" if list(cordes) != [None] else ""))
        if rejets:
            print(f"      rejets : {dict(rejets)}")
        if len(paires) < 15:
            print("      ⚠ corpus réduit : verdicts en critère assoupli (suffixe -cr)")
        ech = paires[::max(1, len(paires) // n_controle)][:n_controle]
        liste = []
        for chemin, norm in ech:
            a, _ = FL.spectre_moyen(chemin)
            if a is not None:
                f0, midi, dyn = FL.analyser_nom(norm)
                if f0:
                    liste.append((f0, midi, dyn, a[:FL.NBINS]))
        if not liste:
            print("      ✗ aucun spectre lisible")
            souci += 1
            continue
        r = FL.controle_nomenclature(liste)
        if r and 0.97 <= r <= 1.03:
            print(f"      nomenclature = {r}")
        else:
            souci += 1
            print(f"      nomenclature = {r}   ⚠⚠ À VÉRIFIER À LA MAIN")
            print("      0,89 ≈ hauteurs écrites et non réelles ; valeur collée à "
                  "0,92/1,09 ≈ transposition franche (cf. clarinette basse ZSamples)")
    print()
    print("VALIDATION : rien à signaler" if not souci
          else f"VALIDATION : {souci} point(s) à régler AVANT de mesurer")
    return souci


# ───────────────────────────── mode --mesurer ───────────────────────────────

def _hz(v):
    """numpy.float64 -> entier. La résolution de la méthode est d'un demi-ton :
    publier des décimales serait mentir."""
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return v


def _recaler(a, f0_nomme):
    """Recale le spectre en fréquence pour que la hauteur RÉELLE tombe sur la
    hauteur nommée. Rend (spectre recalé, rapport r).

    Pourquoi c'est indispensable : harmoniques() cherche chaque harmonique dans
    une fenêtre de ±2 bins (±21,5 Hz) autour de k·f0 **calculé depuis le nom**.
    Si la banque est accordée 20 cents haut, l'écart dépasse la fenêtre dès que
    k·f0 > 1 265 Hz, et tous les harmoniques au-dessus sont relevés dans un
    creux. Mesuré sur la banque VO : +15 à +48 cents selon l'instrument, soit un
    seul harmonique correctement relevé pour la flûte. Sans ce recalage, les
    enveloppes composites de VO sont construites sur des amplitudes fausses.

    La correction est appliquée PAR ÉCHANTILLON, donc la dispersion de l'accord
    (σ ≈ 1 % sur VO) n'intervient pas. Les formants sortent dans le référentiel
    des hauteurs nommées : la fiche les remultiplie par le rapport médian."""
    r = FL.f0_par_somme_harmonique(a, f0_nomme) / f0_nomme
    if not (0.9 < r < 1.1):
        return a, 1.0
    import numpy as np
    b = np.arange(FL.NBINS)
    return np.interp(b * r, b, a), r


def ecrire_spectres_normalises(paires, sortie, corriger=False):
    import numpy as np
    ok = ko = 0
    motifs = collections.Counter()
    rapports = []
    with open(sortie, 'w') as fh:
        fh.write("spectrum 4096 2048 1024\n")
        for chemin, norm in paires:
            a, err = FL.spectre_moyen(chemin)
            if a is None:
                ko += 1
                motifs[str(err)[:40]] += 1
                continue
            if corriger:
                f0, _m, _d = FL.analyser_nom(norm)
                if f0:
                    a, r = _recaler(a, f0)
                    rapports.append(r)
            fh.write(norm + ';' + ';'.join(f"{v:.7g}" for v in a) + '\n')
            ok += 1
    r_med = float(np.median(rapports)) if rapports else 1.0
    return ok, ko, motifs, r_med


def fiche_markdown(nom, corpus, res, n):
    f = res.get('formants', [])
    ret = [x for x in f if str(x.get('verdict', '')).startswith('retenu')]
    L = ['---',
         f'id: {re.sub(r"[^0-9a-z]+", "-", nom.lower()).strip("-")}',
         f'nom: {nom}', f'corpus: {corpus}', f'n_echantillons: {n}',
         f'rho_hauteur: {res.get("rho", "-")}',
         f'pente_db_octave: {res.get("pente", "-")}',
         f'harmoniques_par_note: {res.get("harmoniques_median", "-")}',
         f'balance_impair_pair_db: {res.get("balance_impair_pair", "-")}',
         f'enveloppe_vraie_hz: {res.get("env_vraie", "-")}',
         f'enveloppe_vraie_sd: {res.get("env_vraie_sd", "-")}',
         f'nomenclature: {res.get("nomenclature", "-")}',
         f'accord_corrige: {str(res.get("accord_corrige", False)).lower()}',
         f'accord_rapport: {res.get("accord_rapport", 1.0)}',
         'classe: A_VALIDER_A_LA_MAIN', 'formants:']
    for x in f:
        L.append(f'  - {{ hz: {_hz(x.get("hz"))}, residu_db: {x.get("residu")}, '
                 f'verdict: {x.get("verdict")}, notes: {x.get("n_notes")} }}')
    L += [f'version_mesure: {corpus}-2026-10-25', '---', '',
          f'# {nom} — corpus {corpus}', '',
          '> **Méthode 8** (enveloppe composite). Méthode et critères dans '
          '`../00-METHODE.md`. Statut des formants dans `../_REPLICATION.md`.', '',
          '## Formants', '', '| Zone | Résidu | Notes | Verdict |', '|---|---|---|---|']
    for x in f:
        L.append(f'| {_hz(x.get("hz"))} Hz | {x.get("residu")} dB | '
                 f'{x.get("n_notes")} | {x.get("verdict")} |')
    if not f:
        L.append('| — | — | — | aucun candidat |')
    L += ['', '## Descripteurs', '',
          '| Descripteur | Valeur |', '|---|---|',
          f'| ρ (pic / hauteur) | {res.get("rho", "-")} |',
          f'| Pente spectrale | {res.get("pente", "-")} dB/octave '
          f'(σ {res.get("pente_sd", "-")}) |',
          f'| Harmoniques par note | {res.get("harmoniques_median", "-")} |',
          f'| Balance impair − pair | {res.get("balance_impair_pair", "-")} dB |',
          f'| Enveloppe vraie | {res.get("env_vraie", "-")} Hz '
          f'(σ {res.get("env_vraie_sd", "-")}) |',
          '',
          '> La **pente spectrale n\'est pas comparable entre corpus** : elle dépend du '
          'micro et du local. À lire à l\'intérieur de ce corpus seulement.', '']
    r = res.get('accord_rapport', 1.0)
    if res.get('accord_corrige') and r and abs(r - 1) > 0.002:
        cents = round(1200 * (r - 1) / 0.0005776)
        L += ['> **Correction d\'accord appliquée.** Cette banque est accordée '
              f'{cents:+d} cents par rapport aux hauteurs de ses noms de fichiers '
              f'(rapport {r:.4f}). Sans correction, harmoniques() relève les pics '
              'dans un creux au-delà de ~1 300 Hz et l\'enveloppe composite est '
              'fausse. Le recalage est fait par échantillon et les fréquences '
              'ci-dessus sont remises dans le référentiel réel.', '']
    L += [
          '## Confrontation aux corpus existants — À RÉDIGER', '',
          'Comparer chaque formant ci-dessus à la fiche de l\'instrument nu '
          '(`../NN-*.md`) et à ZSamples (`../../Instruments-ZSamples/`). **Ce corpus est '
          'une troisième source : il peut faire basculer en établis les dix formants '
          'classés « indéterminés » dans `../_REPLICATION.md`.** Route A = même point de '
          'grille ; route B = deux familles de méthodes indépendantes. La grille est au '
          'demi-ton : écrire « Ré7 (≈ 2 350 Hz) », pas une valeur au Hz près.', '',
          '## Observations — À RÉDIGER', '', '## Journal', '',
          f'- **2026-10-25** — Mesure initiale, corpus {corpus}, {n} échantillons. '
          f'{len(ret)} formant(s) retenu(s) sur {len(f)} candidat(s). '
          f'Nomenclature {res.get("nomenclature", "-")}.']
    return '\n'.join(L) + '\n'


def mesurer(desc, seulement=None):
    sortie = desc.get('sortie') or ('Instruments-' + desc['corpus'])
    os.makedirs(sortie, exist_ok=True)
    for cfg in desc['configurations']:
        if seulement and cfg['nom'] not in seulement:
            continue
        paires, rejets = collecter(desc, cfg)
        if len(paires) < 5:
            print(f"✗ {cfg['nom']}: {len(paires)} fichier(s), insuffisant")
            continue
        sp = os.path.join(sortie, cfg['nom'] + '.spectres.txt')
        corr = desc.get('corriger_accord', False)
        ok, ko, motifs, r_med = ecrire_spectres_normalises(paires, sp, corr)
        res = FL.mesurer(sp)
        if 'erreur' in res:
            print(f"✗ {cfg['nom']}: {res['erreur']}")
            continue
        res['accord_corrige'] = corr
        res['accord_rapport'] = round(r_med, 4)
        res['accord_cents'] = round(1200 * (r_med - 1) / 0.0005776, 0) if r_med else 0
        if corr and r_med != 1.0:            # remise dans les fréquences réelles
            for x in res.get('formants', []):
                if x.get('hz'):
                    x['hz'] = float(x['hz']) * r_med
        with open(os.path.join(sortie, cfg['nom'] + '.json'), 'w') as fh:
            json.dump(res, fh, ensure_ascii=False, indent=1, default=float)
        with open(os.path.join(sortie, cfg['nom'] + '.md'), 'w') as fh:
            fh.write(fiche_markdown(cfg['nom'], desc['corpus'], res, ok))
        ret = sorted(_hz(x['hz']) for x in res.get('formants', [])
                     if str(x.get('verdict', '')).startswith('retenu'))
        acc = f"  accord {r_med:.4f}" if corr else ""
        print(f"  {cfg['nom']:26s} {ok:5d} spectres ({ko} échecs"
              + (f" {dict(motifs)}" if motifs else "") + ")  "
              f"nom. {res.get('nomenclature')}{acc}  ρ {res.get('rho')}  retenus: {ret}")
    print(f"\nÉcrit dans {sortie}/")


# ───────────────────────────────── CLI ──────────────────────────────────────

def main():
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--decrire', metavar='RACINE')
    p.add_argument('--valider', metavar='D.json')
    p.add_argument('--mesurer', metavar='D.json')
    p.add_argument('--seulement')
    p.add_argument('--sans-validation', action='store_true',
                   help="passer outre (déconseillé : la nomenclature ne sera pas contrôlée)")
    a = p.parse_args()
    s = set(a.seulement.split(',')) if a.seulement else None
    if a.decrire:
        decrire(a.decrire)
    elif a.valider:
        valider(json.load(open(a.valider, encoding='utf-8')), s)
    elif a.mesurer:
        d = json.load(open(a.mesurer, encoding='utf-8'))
        if not a.sans_validation and valider(d, s):
            print("\nMesure interrompue. Régler les points ci-dessus, ou relancer avec "
                  "--seulement sur les configurations saines.")
            sys.exit(1)
        mesurer(d, s)
    else:
        p.print_help()


if __name__ == '__main__':
    main()
