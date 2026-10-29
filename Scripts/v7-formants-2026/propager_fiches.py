# -*- coding: utf-8 -*-
"""
propager_fiches.py — verse dans les 22 fiches ce que les travaux du 27/10 disent de
chaque instrument : corroboration LPC, effet de la nuance sur la localisation, et place
dans la table des mélanges.

Chaque fiche reçoit une section unique, insérée avant « Observations pour l'orchestration ».
Les trois sources sont des CSV produits par les tâches C4, C5 et C9 ; rien n'est retranscrit
à la main, et relancer le script réécrit la section au lieu de l'empiler.

La section est volontairement courte : le détail vit dans les documents de synthèse, et une
fiche qui répète le chapitre devient un deuxième endroit à corriger.
"""
import os, re, csv, io, sys, json, statistics as st
from collections import defaultdict

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(os.path.dirname(ICI))
ETUDE = os.path.join(RACINE, 'Etude-Formants')
INST = os.path.join(ETUDE, 'Instruments')
MES = os.path.join(ETUDE, 'Mesures')

DEBUT = '<!-- DEBUT-TRAVAUX-27-10 -->'
FIN = '<!-- FIN-TRAVAUX-27-10 -->'
ANCHE_DOUBLE = {'Hautbois', 'Cor anglais', 'Basson', 'Contrebasson'}

def lire_csv(p):
    return list(csv.DictReader(open(p, encoding='utf-8'))) if os.path.exists(p) else []

LPC = lire_csv(os.path.join(MES, 'Appariement-LPC', 'appariement-lpc.csv'))
DYN = lire_csv(os.path.join(MES, 'Largeur-par-registre', 'largeur-par-registre.csv'))
MEL = lire_csv(os.path.join(MES, 'Melanges', 'melanges.csv'))
PED = lire_csv(os.path.join(MES, 'Registres-Pedale', 'par-nuance.csv'))
# `_CORPUS-CROISES.md`, 25/10 : Yan_Adds-Divers et VN sont la même bibliothèque (VSL).
# Pour ces instruments la paire (SOL, VN) est écartée — ce qui n'est PAS un désaccord.
PROV_VSL = {'Trombone basse', 'Tuba contrebasse', 'Contrebasson', 'Cor anglais'}
ACC = lire_csv(os.path.join(MES, 'Registres-Pedale', 'accords-stricts.csv'))
COI = lire_csv(os.path.join(MES, 'Coincidences-F2', 'coincidences-detail.csv'))
COIR = lire_csv(os.path.join(MES, 'Coincidences-F2', 'coincidences-par-instrument.csv'))
RESUME_DYN = json.load(open(os.path.join(MES, 'Largeur-par-registre',
                                         'largeur-par-registre.json'),
                            encoding='utf-8'))['resume']

def bloc_lpc(nom):
    L = [r for r in LPC if r['instrument'] == nom and r['rang_percentile']]
    if not L:
        return None
    lignes = ['### Corroboration par le LPC, note par note',
              '',
              'Rang du formant dans son propre spectre : 50 % signifie que les deux méthodes '
              'ne s\'accordent pas mieux là qu\'ailleurs. Voir `../_APPARIEMENT-LPC.md`.',
              '',
              '| Formant | Statut | Rang |', '|---|---|---|']
    for r in sorted(L, key=lambda r: float(r['formant_hz'])):
        lignes.append('| {} Hz | {} | {} % |'.format(
            r['formant_hz'], r['statut'], r['rang_percentile'].replace('.', ',')))
    if nom in ANCHE_DOUBLE:
        lignes += ['', '> **Cet instrument est une anche double, et le LPC y échoue par '
                   'construction** : il suppose une source filtrée par une cavité séparable, '
                   'alors que l\'anche et la perce sont couplées et que la résonance de '
                   'l\'anche fait partie de la source. Les six formants des anches doubles '
                   'sont tous sous la médiane. **Ses formants restent établis par la route A '
                   '— le critère le plus exigeant — mais ne peuvent pas revendiquer de '
                   'corroboration LPC.**']
    return '\n'.join(lignes)

def bloc_dyn(nom):
    r = next((x for x in RESUME_DYN if x['instrument'] == nom), None)
    if not r or r.get('effet') is None:
        return None
    par = defaultdict(lambda: {'n': 0, 'o': 0})
    for x in DYN:
        if x['instrument'] != nom:
            continue
        k = (x['registre'], x['dynamique'])
        par[k]['n'] += int(x['n'])
        par[k]['o'] += int(x['n']) * int(x['pct_ouvert']) / 100
    lignes = ['### Effet de la nuance sur la localisation',
              '',
              'Part de pics **non localisables** — l\'enveloppe n\'y redescend pas de 3 dB des '
              'deux côtés, donc le formant n\'a pas de largeur mesurable. '
              'Voir `../_LARGEUR-PAR-REGISTRE.md`.',
              '',
              '| | doux | moyen | fort |', '|---|---|---|---|',
              '| **tous registres** | {} % | {} % | {} % |'.format(
                  r['ouvert_doux'], r['ouvert_moyen'], r['ouvert_fort'])]
    ordre = ['sons pédale', 'extrême grave', 'grave', 'chalumeau', 'bas médium', 'gorge',
             'médium', 'clarine', 'haut médium', 'aigu', 'suraigu', 'tous']
    regs = sorted({k[0] for k in par},
                  key=lambda x: ordre.index(x) if x in ordre else 99)
    for reg in regs:
        c = []
        for dy in ('doux', 'moyen', 'fort'):
            v = par.get((reg, dy))
            c.append('—' if not v or v['n'] < 10 else '{:.0f} %'.format(100 * v['o'] / v['n']))
        if c.count('—') < 3 and reg != 'tous':
            lignes.append('| {} | {} | {} | {} |'.format(reg, *c))
    e = r['effet']
    lignes += ['', '**Effet doux → fort : {:+d} point{}.** {}'.format(
        e, '' if abs(e) <= 1 else 's',
        'La localisation s\'améliore en jeu fort.' if e >= 10 else
           ('Elle se dégrade en jeu fort — comportement inverse de celui des trombones et de '
            'la trompette.' if e <= -5 else 'Effet négligeable.'))]
    return '\n'.join(lignes)

def bloc_pedale(nom):
    """Tâche C7, cuivres seulement : la couleur (F1) par nuance, et ce qui réplique à
    nuance égale. Le détail vit dans `_REGISTRES-PEDALE.md`."""
    L = [r for r in PED if r['instrument'] == nom]
    if not L:
        return None
    par = {(r['corpus'], r['nuance']): r for r in L}
    corpus = sorted({r['corpus'] for r in L},
                    key=lambda c: ['SOL', 'ZSamples', 'Xsamples', 'VO', 'VN'].index(c)
                    if c in ('SOL', 'ZSamples', 'Xsamples', 'VO', 'VN') else 99)
    lignes = ['### La couleur dépend de la nuance, et ne se lit que jusqu\'au médium',
              '',
              'F1 du motif d\'excitation, notes **pédale + grave + médium** réunies : au-dessus '
              'du médium la mesure donne la fondamentale et non une couleur (F1/f0 = 1,0, pente '
              'contre f0 de 1,03 au suraigu). Voir `../_REGISTRES-PEDALE.md`.',
              '',
              '| Corpus | doux | moyen | fort | doux → fort |', '|---|---|---|---|---|']
    for c in corpus:
        cel, bas, haut = [], None, None
        for dy in ('doux', 'moyen', 'fort'):
            r = par.get((c, dy))
            cel.append('—' if not r else '{} Hz'.format(int(round(float(r['F1'])))))
            if r and dy == 'doux':
                bas = float(r['F1'])
            if r and dy == 'fort':
                haut = float(r['F1'])
        import math as _m
        ec = (('{:+.1f} dt'.format(12 * _m.log2(haut / bas))).replace('.', ',')
              if bas and haut else '—')
        lignes.append('| {} | {} | {} | {} | **{}** |'.format(c, *cel, ec))
    A = [r for r in ACC if r['instrument'] == nom]
    if A:
        lignes += ['',
                   '**Ce qui réplique, à nuance égale** — paires de corpus de provenances '
                   'distinctes accordées à moins d\'un demi-ton :', '']
        for r in sorted(A, key=lambda r: float(r['ecart_demi_tons'])):
            lignes.append(('- **{}** : {} {} Hz / {} {} Hz — **{} demi-ton**, pentes contre '
                           'f0 {:+.2f} et {:+.2f}').format(
                              r['nuance'], r['corpus_a'], r['F1_a'], r['corpus_b'], r['F1_b'],
                              r['ecart_demi_tons'], float(r['pente_a']),
                              float(r['pente_b'])).replace('.', ','))
    elif nom in PROV_VSL and set(corpus) <= {'SOL', 'VN'}:
        lignes += ['',
                   '**Aucune paire utilisable.** Les deux seuls corpus de cet instrument, SOL '
                   '(`Yan_Adds-Divers`) et VN, sont **la même bibliothèque** — même nombre '
                   'd\'échantillons au sample près, corrélation 0,996. Ils s\'accordent de fait '
                   'à 0,06 demi-ton sur la tessiture basse, et **cela ne compte pas** : c\'est '
                   'la même source comparée à elle-même. Voir la règle de provenance dans '
                   '`../_CORPUS-CROISES.md`. La couleur de cet instrument reste **non '
                   'répliquée**, faute de second corpus.']
    else:
        lignes += ['',
                   '**Aucune paire de corpus ne s\'accorde à moins d\'un demi-ton**, à aucune '
                   'nuance. La couleur de cet instrument n\'est pas répliquée.']
    return '\n'.join(lignes)

def bloc_coincidence(nom):
    """Tâche C8 : la carte vocalique tombe-t-elle sur les formants établis du même
    instrument ? Non en général — le bloc le dit aussi quand c'est non, pour qu'on ne
    croie pas la question non posée."""
    R = [r for r in COIR if r['instrument'] == nom and r['seuil'] == '0.5']
    if not R:
        return None
    L = [r for r in COI if r['instrument'] == nom and float(r['ecart_demi_tons']) <= 1.0]
    titre = '### La carte vocalique corrobore-t-elle les formants ?'
    ref = ('Témoin : les propres formants de l\'instrument transposés de ±2 à ±6 demi-tons, '
           'ce qui conserve leur nombre et leur espacement. Voir `../_COINCIDENCES-F2.md`.')
    def fr(x, d=1):
        return ('{:.%df}' % d).format(float(x)).replace('.', ',')
    def mil(x):
        n = int(float(x))
        return '{:,}'.format(n).replace(',', '\u202f')
    fort = [r for r in R if r['rapport'] and float(r['rapport']) >= 2.0]
    if not fort or not L:
        pires = ' · '.join('{} {} % contre {} % au témoin'.format(
            r['maximum'], fr(r['taux']), fr(r['taux_temoin'])) for r in R)
        return '\n'.join([
            titre, '',
            '**Non.** {} — soit le hasard, ou moins. Les deux méthodes ne mesurent pas la '
            'même région du spectre : F1 se situe en médiane à 15,5 demi-tons du formant '
            'établi le plus proche, parce que la méthode 8 ne mesure rien sous 3,2 à 5,7 fois '
            'f0_min.'.format(pires), '', ref])
    lignes = [titre, '',
              '**Oui, et c\'est le seul cas de l\'étude.** Rapport mesure / témoin :', '',
              '| Maximum | Mesure | Témoin | Rapport |', '|---|---|---|---|']
    for r in R:
        lignes.append('| {} | {} % | {} % | **×{}** |'.format(
            r['maximum'], fr(r['taux']), fr(r['taux_temoin']), fr(r['rapport'])))
    par = defaultdict(lambda: [0, 0, set()])
    for r in L:
        k = r['formant_hz']
        par[k][0] += 1
        par[k][1] += 1 if float(r['ecart_demi_tons']) <= 0.5 else 0
        par[k][2].add(r['corpus'])
    lignes += ['', 'Les cellules concernées, formant par formant :', '',
               '| Formant | Cellules < 1,0 dt | dont < 0,5 dt | Corpus |', '|---|---|---|---|']
    for k in sorted(par, key=lambda x: float(x)):
        n, s, cs = par[k]
        lignes.append('| **{} Hz** | {} | {} | {} |'.format(
            mil(k), n, s, ' · '.join(sorted(cs))))
    lignes += ['', ref]
    return '\n'.join(lignes)

def bloc_mel(nom):
    L = [r for r in MEL if nom in (r['A'], r['B'])]
    if not L:
        return None
    def autre(r):
        return r['B'] if r['A'] == nom else r['A']
    f = sorted(L, key=lambda r: -int(r['pct_fusion'] or 0))[:3]
    h = sorted(L, key=lambda r: -int(r['pct_homogene'] or 0))[:3]
    moy = st.mean([int(r['pct_fusion'] or 0) for r in L])
    lignes = ['### Dans la table des mélanges',
              '',
              '{} paires calculées. **Fusion moyenne : {:.0f} %.** '
              'Voir `../_MELANGES.md`.'.format(len(L), moy),
              '',
              '| Couleur nouvelle avec | | Se fond avec | |', '|---|---|---|---|']
    for a, b in zip(f, h):
        lignes.append('| {} | {} % | {} | {} % |'.format(
            autre(a), a['pct_fusion'], autre(b), b['pct_homogene']))
    return '\n'.join(lignes)

def main():
    fiches = sorted(f for d in ('FullSOL2020', 'Yan_Adds')
                    for f in (os.path.join(INST, d, x) for x in os.listdir(os.path.join(INST, d)))
                    if re.match(r'^\d\d-.*\.md$', os.path.basename(f)))
    n = 0
    for p in fiches:
        t = io.open(p, encoding='utf-8').read()
        m = re.match(r'^---\n(.*?)\n---\n', t, re.S)
        nom = re.search(r'^nom:\s*(.+)$', m.group(1), re.M).group(1).strip() if m else None
        if not nom:
            print('  !! sans nom :', os.path.basename(p)); continue
        blocs = [b for b in (bloc_lpc(nom), bloc_dyn(nom), bloc_pedale(nom),
                             bloc_coincidence(nom), bloc_mel(nom)) if b]
        if not blocs:
            print('  — {:30s} aucune donnée'.format(nom)); continue
        section = (DEBUT + '\n## Ce que les travaux des 27 et 28/10/2026 disent de cet '
                   'instrument\n\n'
                   + '\n\n'.join(blocs) + '\n\n' + FIN + '\n\n')
        if DEBUT in t:
            t = re.sub(re.escape(DEBUT) + r'.*?' + re.escape(FIN) + r'\n\n',
                       section, t, flags=re.S)
        else:
            a = "## Observations pour l'orchestration"
            if t.count(a) != 1:
                print('  !! {} : ancre en {} exemplaires'.format(nom, t.count(a))); continue
            t = t.replace(a, section + a, 1)
        io.open(p, 'w', encoding='utf-8').write(t)
        n += 1
        print('  + {:30s} {}'.format(nom, ' · '.join(
            b.split('\n')[0].replace('### ', '') for b in blocs)))
    print('\n%d fiches sur %d mises à jour' % (n, len(fiches)))

if __name__ == '__main__':
    main()
