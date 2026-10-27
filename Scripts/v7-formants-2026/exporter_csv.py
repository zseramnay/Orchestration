#!/usr/bin/env python3
"""Exporte les fiches d'instrument en CSV, pour réalimenter la chaîne HTML/docx.

Pourquoi un nouveau schéma
--------------------------
L'ancien `formants_all_techniques.csv` a six colonnes fixes `F1_hz`…`F6_hz`,
plus quartiles, écarts-types et largeurs. Ce modèle suppose que tout instrument
possède six formants ordonnés. Nous avons établi le contraire :

* le nombre de formants retenus va de **zéro à quatre** selon l'instrument, et
  zéro est un **résultat** (petite flûte sur trois corpus, hautbois sur quatre),
  pas une absence de mesure. Six colonnes vides ne sauraient le dire ;
* chaque formant porte une **route** et une **liste de corpus** — la preuve est
  là, et l'ancien schéma n'a pas de place pour elle ;
* les quartiles et écarts-types de l'ancien pipeline décrivaient la dispersion
  d'un F1 note à note, notion que nous avons abandonnée ;
* `bandwidths_3db.csv`, que `get_f()` lit pour remplir `bw`, a **toutes ses
  valeurs multiples entiers du bin de FFT** : c'est l'ancien pipeline et ses
  défauts.

D'où **deux fichiers en format long**, le nombre de formants étant variable :

`formants-v7.csv`    une ligne par formant, avec sa route, ses corpus, son écart
`instruments-v7.csv` une ligne par instrument, les descripteurs globaux

Usage
-----
    python3 exporter_csv.py --racine ETUDE --sortie ETUDE/Resultats-v7
"""

import argparse, csv, glob, itertools, json, math, os, re, unicodedata

NOTES = ['Do', 'Do♯', 'Ré', 'Ré♯', 'Mi', 'Fa', 'Fa♯', 'Sol', 'Sol♯', 'La', 'La♯', 'Si']


def note_tempérée(hz):
    """La résolution de la méthode est d'un demi-ton : publier une fréquence au
    Hz près serait mentir. On donne donc la note, et le Hz en appoint."""
    if not hz or hz <= 0:
        return ''
    m = 69 + 12 * math.log2(hz / 440.0)
    i = int(round(m))
    return '{}{}'.format(NOTES[i % 12], i // 12 - 1)


def slug(s):
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn').lower()
    s = re.sub(r'\b(en)\s+(sib|mib|fa|ut)\b', '', s)
    return re.sub(r'[^a-z]+', '', s)


ALIAS = {'corenfa': 'cor', 'trompetteenut': 'trompette', 'trombone': 'trombonetenor'}


def cle(s):
    k = slug(s)
    return ALIAS.get(k, k)


def champ(y, nom, defaut=''):
    m = re.search(r'^%s:\s*(.+?)\s*(?:#.*)?$' % re.escape(nom), y, re.M)
    return m.group(1).strip() if m else defaut


def formants_yaml(y):
    """[(hz, residu, verdict, statut, notes), …] dans l'ordre du fichier."""
    out = []
    bloc = re.search(r'^formants:\s*\n((?:\s*-\s*\{.*\n)+)', y, re.M)
    if not bloc:
        return out
    for ln in bloc.group(1).strip().split('\n'):
        h = re.search(r'hz:\s*(\d+)', ln)
        if not h:
            continue
        out.append(dict(
            hz=int(h.group(1)),
            residu=champ_inline(ln, 'residu_db'),
            verdict=champ_inline(ln, 'verdict'),
            statut=(re.search(r'statut:\s*"([^"]*)"', ln) or [None, ''])[1]
            if re.search(r'statut:\s*"', ln) else '',
            notes=champ_inline(ln, 'notes')))
    return out


def champ_inline(ln, nom):
    m = re.search(r'%s:\s*"?([^",}\]]+)"?' % re.escape(nom), ln)
    return m.group(1).strip() if m else ''


def corpus_voyants(k, f0, CORPUS):
    """Corpus dont un formant retenu tombe à moins d'un demi-ton."""
    out = []
    for c, d in CORPUS.items():
        for f in d.get(k, []):
            if abs(12 * math.log2(f / f0)) <= 1.05:
                out.append(c)
                break
    return out


def charger_corpus(racine):
    C = {}
    for c in ('ZSamples', 'Xsamples', 'VO', 'VN'):
        d = {}
        for p in sorted(glob.glob(os.path.join(racine, 'Mesures', c, '*.json'))):
            b = os.path.basename(p)[:-5]
            if any(x in b for x in (' corde', '+sourdine', '+piombo', '+sord', 'vibrato')):
                continue
            try:
                j = json.load(open(p))
            except Exception:
                continue
            d[cle(b)] = sorted(int(round(float(x['hz']))) for x in j.get('formants', [])
                               if str(x.get('verdict', '')).startswith('retenu'))
        C[c] = d
    return C


def charger_csv(p, cle_inst='instrument'):
    out = {}
    if os.path.exists(p):
        for r in csv.DictReader(open(p)):
            out.setdefault(cle(r[cle_inst]), []).append(r)
    return out


def main():
    a = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    a.add_argument('--racine', required=True)
    a.add_argument('--sortie', required=True)
    o = a.parse_args()
    os.makedirs(o.sortie, exist_ok=True)

    CORPUS = charger_corpus(o.racine)
    RHO = charger_csv(os.path.join(o.racine, 'Largeurs', 'Comparaison', 'rho_par_corpus.csv'))
    LOC = charger_csv(os.path.join(o.racine, 'Largeurs', 'Comparaison',
                                   'confrontation_lpc_enveloppe.csv'))

    lignes_f, lignes_i = [], []
    for p in sorted(glob.glob(os.path.join(o.racine, 'Instruments', '*',
                                           '[0-9][0-9]-*.md'))):
        t = open(p, encoding='utf-8').read()
        m = re.match(r'^---\n(.*?)\n---\n', t, re.S)
        if not m:
            continue
        y = m.group(1)
        nom = champ(y, 'nom')
        if not nom:
            continue
        k = cle(nom)
        fs = formants_yaml(y)
        retenus = [f for f in fs if f['verdict'].startswith('retenu')]
        rho = RHO.get(k, [{}])[0]
        loc = {int(r['formant_hz']): r for r in LOC.get(k, []) if r.get('formant_hz')}

        par_route = {}
        for f in retenus:
            par_route[f['statut']] = par_route.get(f['statut'], 0) + 1

        lignes_i.append(dict(
            id=champ(y, 'id'), instrument=nom, nom_en=champ(y, 'nom_en'),
            famille=champ(y, 'famille'), sous_famille=champ(y, 'sous_famille'),
            ordre=champ(y, 'ordre'), provenance=champ(y, 'provenance'),
            source=champ(y, 'source_echantillons'),
            n_echantillons=champ(y, 'n_echantillons'),
            tessiture=champ(y, 'tessiture_mesuree'),
            f0_min_hz=champ(y, 'f0_min_hz'), f0_max_hz=champ(y, 'f0_max_hz'),
            classe=champ(y, 'classe'), pertinence=champ(y, 'pertinence_formant'),
            rho_SOL=champ(y, 'rho_hauteur'),
            rho_ZSamples=rho.get('rho_ZSamples', ''), rho_Xsamples=rho.get('rho_Xsamples', ''),
            rho_VO=rho.get('rho_VO', ''), rho_VN=rho.get('rho_VN', ''),
            rho_etendue=rho.get('etendue', ''),
            n_formants_retenus=len(retenus),
            n_route_A=par_route.get('route A', 0),
            n_indetermine=par_route.get('indéterminé', 0),
            n_un_corpus=par_route.get('1 corpus', 0),
            pente_db_octave=champ(y, 'pente_db_octave'), pente_sd=champ(y, 'pente_sd'),
            harmoniques_par_note=champ(y, 'harmoniques_par_note'),
            balance_impair_pair_db=champ(y, 'balance_impair_pair_db'),
            enveloppe_vraie_hz=champ(y, 'enveloppe_vraie_hz'),
            enveloppe_vraie_sd=champ(y, 'enveloppe_vraie_sd'),
            version_mesure=champ(y, 'version_mesure')))

        for rang, f in enumerate(fs, 1):
            vus = ['SOL'] + corpus_voyants(k, f['hz'], CORPUS)
            l = loc.get(f['hz'], {})
            q = l.get('Q_env_v7', '')
            lignes_f.append(dict(
                instrument=nom, id=champ(y, 'id'), famille=champ(y, 'famille'),
                ordre=champ(y, 'ordre'), provenance=champ(y, 'provenance'),
                rang=rang, hz=f['hz'], note=note_tempérée(f['hz']),
                verdict=f['verdict'], statut=f['statut'] or '',
                corpus=' '.join(vus), n_corpus=len(vus),
                residu_db=f['residu'], n_notes=f['notes'],
                largeur_lpc_hz=l.get('lpc_v7', ''), largeur_env_hz=l.get('env_v7', ''),
                Q_enveloppe=q,
                localisation=('non localisé' if q and float(q) < 2 else
                              ('localisé' if q else ''))))

    # ---- zones établies dont aucune fiche ne porte la valeur ------------
    # Appariement SANS regroupement préalable : un regroupement séquentiel
    # chaîne depuis le premier élément d'une zone et peut séparer deux valeurs
    # qui s'apparient. C'est ce qui avait fait manquer trois formants le 25/10,
    # dont celui de l'alto.
    prov = {cle(l['instrument']): l['provenance'] for l in lignes_i}
    deja = {}
    for l in lignes_f:
        deja.setdefault(cle(l['instrument']), []).append(l['hz'])
    info_de = {cle(l['instrument']): l for l in lignes_i}
    TOUS = dict(CORPUS)
    TOUS['SOL'] = {k: [h for h in deja.get(k, [])] for k in deja}
    for k, info in info_de.items():
        pts = [(f, c) for c, d in TOUS.items() for f in d.get(k, [])]
        etab = []
        for (a, ca), (b, cb) in itertools.combinations(pts, 2):
            if ca == cb or abs(12 * math.log2(a / b)) >= 0.5:
                continue
            if {ca, cb} == {'SOL', 'VN'} and prov.get(k) == 'VSL':
                continue
            etab.append((math.sqrt(a * b), ca, cb))
        etab.sort()
        fus = []
        for h, ca, cb in etab:
            if fus and abs(12 * math.log2(h / fus[-1][0])) < 1.0:
                fus[-1][1] |= {ca, cb}
            else:
                fus.append([h, {ca, cb}])
        for h, cs in fus:
            centre = round(h)
            if any(abs(12 * math.log2(centre / x)) <= 1.05 for x in deja.get(k, [])):
                continue          # déjà porté par une ligne de fiche
            lignes_f.append(dict(
                instrument=info['instrument'], id=info['id'], famille=info['famille'],
                ordre=info['ordre'], provenance=info['provenance'],
                rang='', hz=centre, note=note_tempérée(centre),
                verdict='retenu (hors SOL)', statut='route A',
                corpus=' '.join(sorted(cs)), n_corpus=len(cs),
                residu_db='', n_notes='', largeur_lpc_hz='', largeur_env_hz='',
                Q_enveloppe='', localisation=''))
    lignes_f.sort(key=lambda l: (int(l['ordre'] or 0), int(l['hz'])))

    f1 = os.path.join(o.sortie, 'formants-v7.csv')
    with open(f1, 'w', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=list(lignes_f[0]))
        w.writeheader()
        w.writerows(lignes_f)
    f2 = os.path.join(o.sortie, 'instruments-v7.csv')
    with open(f2, 'w', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=list(lignes_i[0]))
        w.writeheader()
        w.writerows(lignes_i)
    print('{} : {} formants ({} instruments)'.format(os.path.basename(f1), len(lignes_f),
                                                     len(lignes_i)))
    print('{} : {} instruments'.format(os.path.basename(f2), len(lignes_i)))
    ra = sum(1 for x in lignes_f if x['statut'] == 'route A')
    print('   dont route A : {}   indéterminé : {}   1 corpus : {}'.format(
        ra, sum(1 for x in lignes_f if x['statut'] == 'indéterminé'),
        sum(1 for x in lignes_f if x['statut'] == '1 corpus')))
    print('   sans aucun formant retenu : {}'.format(
        sum(1 for x in lignes_i if x['n_formants_retenus'] == 0)))


if __name__ == '__main__':
    main()
