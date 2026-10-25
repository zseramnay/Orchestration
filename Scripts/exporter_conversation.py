#!/usr/bin/env python3
"""Convertit une transcription Claude Code (.jsonl) en Markdown lisible.

Le fichier .jsonl est le journal brut d'une session : une entrée JSON par ligne,
mêlant messages, appels d'outils, résultats, pièces jointes internes et rappels
système. Ce script en extrait le dialogue et le met en forme.

Usage
-----
    python3 exporter_conversation.py TRANSCRIPTION.jsonl -o sortie.md
    python3 exporter_conversation.py T.jsonl -o s.md --pensees --max-resultat=4000
    python3 exporter_conversation.py T.jsonl -o s.md --complet   # rien n'est tronqué

Options
    --pensees          inclut les blocs de raisonnement interne, en sections repliables
    --max-resultat=N   longueur maximale d'un résultat d'outil (défaut 1200 ; 0 = omettre)
    --complet          aucune troncature, et inclut les pensées
    --sessions=a,b     ne garde que ces identifiants de session
    --titre="..."      titre du document

Ce qui est systématiquement retiré
    les blocs <system-reminder>, qui sont du contexte injecté par l'application et
    non des paroles de l'utilisateur ; les pièces jointes internes (instantanés de
    mémoire, listes d'outils, rappels de quota) ; les résumés de hooks.
"""

import argparse, json, os, re, sys, collections, datetime

RX_REMINDER = re.compile(r'<system-reminder>.*?</system-reminder>', re.S)
RX_VIDE = re.compile(r'\n{3,}')


def propre(txt):
    """Retire les blocs injectés par l'application et normalise les blancs."""
    if not txt:
        return ''
    t = RX_REMINDER.sub('', txt)
    t = re.sub(r'<(user_memory_snapshot|memory_listing|profile|preferences)>.*?</\1>', '', t, flags=re.S)
    return RX_VIDE.sub('\n\n', t).strip()


def apercu_args(nom, args, limite=400):
    """Résume les arguments d'un appel d'outil sur une ligne ou deux."""
    if not isinstance(args, dict):
        return ''
    cles = ('command', 'file_path', 'path', 'pattern', 'query', 'url', 'description',
            'prompt', 'old_string', 'new_string', 'content', 'files', 'devicePath')
    bouts = []
    for k in cles:
        if k not in args:
            continue
        v = args[k]
        v = json.dumps(v, ensure_ascii=False) if not isinstance(v, str) else v
        v = re.sub(r'\s+', ' ', v).strip()
        if len(v) > limite:
            v = v[:limite] + ' […]'
        bouts.append(f'`{k}` : {v}')
        if len(bouts) >= 3:
            break
    if not bouts:
        autres = {k: v for k, v in args.items() if k not in cles}
        s = re.sub(r'\s+', ' ', json.dumps(autres, ensure_ascii=False))
        bouts = [s[:limite] + (' […]' if len(s) > limite else '')] if autres else []
    return '  \n'.join(bouts)


def texte_resultat(bloc):
    c = bloc.get('content')
    if isinstance(c, str):
        return c
    if isinstance(c, list):
        return '\n'.join(x.get('text', '') for x in c if isinstance(x, dict))
    return ''


def charger(chemin, sessions=None):
    entrees = []
    with open(chemin, errors='ignore') as fh:
        for L in fh:
            L = L.strip()
            if not L:
                continue
            try:
                o = json.loads(L)
            except Exception:
                continue
            if sessions and o.get('sessionId') not in sessions:
                continue
            entrees.append(o)
    entrees.sort(key=lambda o: (o.get('timestamp') or '', 0))
    return entrees


def convertir(entrees, pensees=False, max_res=1200, titre=None):
    L = []
    sess = [o.get('sessionId') for o in entrees if o.get('sessionId')]
    bornes = {}
    for o in entrees:
        s, t = o.get('sessionId'), o.get('timestamp')
        if s and t:
            a, b = bornes.get(s, (t, t))
            bornes[s] = (min(a, t), max(b, t))
    ordre = sorted(bornes, key=lambda s: bornes[s][0])

    def jour(ts):
        return (ts or '')[:10]

    L.append('---')
    L.append(f'titre: {titre or "Transcription de conversation"}')
    L.append(f'export: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}')
    L.append(f'sessions: {len(ordre)}')
    L.append(f'entrees: {len(entrees)}')
    if ordre:
        L.append(f'periode: {jour(bornes[ordre[0]][0])} → {jour(bornes[ordre[-1]][1])}')
    L.append('---')
    L.append('')
    L.append(f'# {titre or "Transcription de conversation"}')
    L.append('')
    L.append('> Export automatique d\'une transcription Claude Code. Les blocs '
             '`<system-reminder>`, les instantanés de mémoire et les rappels de quota — '
             'injectés par l\'application et non écrits par l\'utilisateur — ont été retirés.')
    L.append('')
    if len(ordre) > 1:
        L.append('## Sessions')
        L.append('')
        L.append('| Session | Entrées | Début | Fin |')
        L.append('|---|---|---|---|')
        for s in ordre:
            n = sum(1 for o in entrees if o.get('sessionId') == s)
            L.append(f'| `{s[:8]}…` | {n} | {bornes[s][0][:16].replace("T", " ")} '
                     f'| {bornes[s][1][:16].replace("T", " ")} |')
        L.append('')

    sess_courante = None
    jour_courant = None
    n_user = n_claude = n_outils = 0

    for o in entrees:
        t = o.get('type')
        s = o.get('sessionId')
        ts = o.get('timestamp') or ''

        if t == 'system' and o.get('subtype') == 'compact_boundary':
            L.append('')
            L.append('---')
            L.append('')
            L.append('> ### ⟨ limite de contexte atteinte — la conversation reprend sur un résumé ⟩')
            L.append('')
            continue

        if t not in ('user', 'assistant'):
            continue

        msg = o.get('message') or {}
        contenu = msg.get('content')
        if isinstance(contenu, str):
            contenu = [{'type': 'text', 'text': contenu}]
        if not isinstance(contenu, list):
            continue

        morceaux = []
        for b in contenu:
            if not isinstance(b, dict):
                continue
            tb = b.get('type')
            if tb == 'text':
                p = propre(b.get('text'))
                if p:
                    morceaux.append(('texte', p))
            elif tb == 'thinking' and pensees:
                p = propre(b.get('thinking') or b.get('text'))
                if p:
                    morceaux.append(('pensee', p))
            elif tb == 'tool_use':
                morceaux.append(('outil', (b.get('name', '?'), b.get('input') or {})))
            elif tb == 'tool_result':
                if max_res != 0:
                    morceaux.append(('resultat', texte_resultat(b)))
        if not morceaux:
            continue

        if s != sess_courante:
            sess_courante = s
            jour_courant = None
            if len(ordre) > 1:
                L.append('')
                L.append(f'## Session `{(s or "?")[:8]}…`')
                L.append('')
        j = jour(ts)
        if j and j != jour_courant:
            jour_courant = j
            L.append('')
            L.append(f'### {j}')
            L.append('')

        for genre, val in morceaux:
            if genre == 'texte':
                if t == 'user':
                    n_user += 1
                    L.append('')
                    L.append('#### 🗣 Yan')
                    L.append('')
                    L.append(val)
                else:
                    n_claude += 1
                    L.append('')
                    L.append('#### 💬 Claude')
                    L.append('')
                    L.append(val)
            elif genre == 'pensee':
                L.append('')
                L.append('<details><summary>raisonnement interne</summary>')
                L.append('')
                L.append(val)
                L.append('')
                L.append('</details>')
            elif genre == 'outil':
                n_outils += 1
                nom, args = val
                L.append('')
                L.append(f'<details><summary>🔧 <code>{nom}</code></summary>')
                L.append('')
                ap = apercu_args(nom, args)
                if ap:
                    L.append(ap)
                L.append('')
                L.append('</details>')
            elif genre == 'resultat':
                r = (val or '').strip()
                if not r:
                    continue
                tronque = False
                if max_res and len(r) > max_res:
                    r = r[:max_res]
                    tronque = True
                L.append('')
                L.append('<details><summary>↳ résultat</summary>')
                L.append('')
                L.append('```')
                L.append(r)
                if tronque:
                    L.append(f'[… tronqué, {len(val) - max_res} caractères de plus]')
                L.append('```')
                L.append('')
                L.append('</details>')

    L.append('')
    L.append('---')
    L.append('')
    L.append(f'*{n_user} messages de l\'utilisateur · {n_claude} réponses · '
             f'{n_outils} appels d\'outils.*')
    return '\n'.join(L) + '\n'


def main():
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('transcription')
    p.add_argument('-o', '--sortie', required=True)
    p.add_argument('--pensees', action='store_true')
    p.add_argument('--complet', action='store_true')
    p.add_argument('--max-resultat', type=int, default=1200)
    p.add_argument('--sessions')
    p.add_argument('--titre')
    a = p.parse_args()
    sessions = {x.strip() for x in a.sessions.split(',')} if a.sessions else None
    entrees = charger(a.transcription, sessions)
    if not entrees:
        sys.exit('aucune entrée lue — vérifier le chemin ou --sessions')
    md = convertir(entrees,
                   pensees=a.pensees or a.complet,
                   max_res=0 if False else (10 ** 9 if a.complet else a.max_resultat),
                   titre=a.titre)
    with open(a.sortie, 'w') as fh:
        fh.write(md)
    print(f'{a.sortie} — {len(md):,} caractères, {md.count(chr(10)):,} lignes, '
          f'{len(entrees):,} entrées lues')


if __name__ == '__main__':
    main()
