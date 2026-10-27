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
        blocs = [b for b in (bloc_lpc(nom), bloc_dyn(nom), bloc_mel(nom)) if b]
        if not blocs:
            print('  — {:30s} aucune donnée'.format(nom)); continue
        section = (DEBUT + '\n## Ce que les travaux du 27/10/2026 disent de cet instrument\n\n'
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
