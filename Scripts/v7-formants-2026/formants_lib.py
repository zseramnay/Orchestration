#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
formants_lib.py — bibliothèque autonome de la mesure formantique (version 2026-10).

AUCUNE dépendance à un autre fichier du projet. Requiert numpy et soundfile.

Chaîne complète :
  1. WAV  -> spectre moyen, attaque exclue          : spectre_moyen()
  2. spectre -> amplitudes harmoniques               : harmoniques()
  3. amplitudes -> résidus locaux                    : residu_local()
  4. résidus -> enveloppe composite (méthode 8)      : composite()
  5. enveloppe -> candidats                          : trouver_formants()
  6. candidats -> verdicts (tests de rejet)          : verdict()
  7. lecture note par note (méthode 6)               : enveloppe_vraie()

Voir Etude-Formants/Instruments/00-METHODE.md pour la justification de chaque choix.
"""
import os, re, math, glob
import numpy as np

# ─────────────────────────── constantes de mesure ───────────────────────────
FS, NFFT, HOP, NBINS = 44100.0, 4096, 2048, 1024
BW    = FS / NFFT              # 10,7666 Hz par bin
FMAX  = (NBINS - 1) * BW       # ~11 015 Hz
QCOMB = 2 * (NBINS - 1)        # longueur effective du cepstre demi-spectre

SKIP_MS, DUREE_MS, MIN_MS = 200, 1500, 250   # segment analysé après l'onset
HALFWIN   = 0.75               # ±3/4 d'octave : fenêtre du redressement local.
                               # NE PAS PORTER À 1,00 — essayé et annulé le 27/10/2026.
                               # À 1,00, 96 % des formants route A se retrouvent À UN TIERS
                               # D'OCTAVE PRÈS, ce qui semblait sans coût. Mais la route A
                               # exige un accord à un DEMI-TON entre corpus : une dérive
                               # d'un demi-ton par corpus, invisible à la tolérance du test
                               # de robustesse, lui est fatale. Résultat mesuré : 23 -> 19
                               # formants route A, et les TROIS unissons formantiques
                               # (1 108, 1 244, 2 216 Hz) disparaissent.
                               # Voir `_FENETRE-DE-REDRESSEMENT.md`.
DELTA_ROB = 0.25               # écart des fenêtres de comparaison du test de robustesse
RESIDU_MIN = 1.0               # seuil de résidu, en dB — NE PAS ABAISSER
NBOOT      = 2000              # tirages du bootstrap par note
MIN_NOTES_GROUPE = 3           # plancher de testabilité d'un groupe de rang
MIN_PTS_GROUPE   = 5

F_LO, F_HI = 80.0, 10000.0     # grille d'analyse, 1/12 d'octave
NB      = int(round(12 * math.log2(F_HI / F_LO)))
BORDS   = F_LO * 2 ** (np.arange(NB + 1) / 12.0)
CENTRES = np.sqrt(BORDS[:-1] * BORDS[1:])

_NOTE = re.compile(r'^([A-G])(#?)(-?\d{1,2})$')
_SEMI = {'C': 0, 'D': 2, 'E': 4, 'F': 5, 'G': 7, 'A': 9, 'B': 11}
DYN_DOUCES = {'p', 'pp', 'ppp', 'mp'}
DYN_FORTES = {'f', 'ff', 'fff', 'mf'}

# ─────────────────────────── nom de fichier -> hauteur ──────────────────────
def note_vers_f0(jeton):
    """'A#3' -> (233.08, 58). Rend None si le jeton n'est pas une note."""
    m = _NOTE.match(jeton)
    if not m:
        return None
    midi = 12 * (int(m.group(3)) + 1) + _SEMI[m.group(1)] + (1 if m.group(2) else 0)
    if not 12 <= midi <= 120:
        return None
    return 440.0 * 2 ** ((midi - 69) / 12.0), midi

def analyser_nom(chemin):
    """Extrait (f0, midi, dynamique) du nom de fichier SOL/Yan_Adds.
    Les noms sont en HAUTEURS RÉELLES : vérifié, rapport pic/f0 = 1,00 sur 22
    instruments. Si vous ajoutez une base, refaites ce contrôle (voir
    controle_nomenclature())."""
    base = os.path.basename(chemin).replace('.wav', '')
    jetons = base.split('-')
    for i, j in enumerate(jetons):
        r = note_vers_f0(j)
        if r:
            dyn = jetons[i + 1].lower() if i + 1 < len(jetons) else None
            return r[0], r[1], dyn
    return None, None, None

# ─────────────────── 1. WAV -> spectre moyen, attaque exclue ────────────────
_WIN = np.hanning(NFFT)

def spectre_moyen(chemin_wav):
    """Spectre moyen d'amplitude, transitoire d'attaque EXCLU.
    Rend (vecteur de 1024 bins normalisé, None) ou (None, motif d'échec).

    Pourquoi exclure l'attaque : le compute_features d'Orchidea démarre à
    l'échantillon 0 et pondère chaque trame par son énergie ; l'attaque, la plus
    énergétique, pèse donc lourd dans sa moyenne. Les deux mesures concordent à
    un ou deux douzièmes d'octave, mais la présente est la bonne."""
    import soundfile as sf
    x, sr = sf.read(chemin_wav, always_2d=True)
    x = x.mean(axis=1).astype(np.float64)
    if sr != FS:
        return None, f"fréquence d'échantillonnage {sr}"
    k = max(1, int(sr * 0.01))
    rms = np.convolve(np.abs(x), np.ones(k) / k, mode='same')
    idx = np.nonzero(rms > 0.2 * rms.max())[0]
    onset = idx[0] if idx.size else 0
    d = onset + int(sr * SKIP_MS / 1000)
    f = min(len(x), d + int(sr * DUREE_MS / 1000))
    if f - d < int(sr * MIN_MS / 1000):
        d, f = onset, len(x)
        if f - d < NFFT:
            return None, "trop court"
    seg = x[d:f]
    trames = [np.abs(np.fft.rfft(seg[i:i + NFFT] * _WIN, n=NFFT))
              for i in range(0, len(seg) - NFFT + 1, HOP)]
    if not trames:
        pad = np.zeros(NFFT)
        pad[:len(seg)] = seg[:NFFT] * _WIN[:len(seg)]
        trames = [np.abs(np.fft.rfft(pad, n=NFFT))]
    moy = np.mean(trames, axis=0)[:NBINS]
    s = moy.sum()
    return (moy / s if s > 0 else moy), None

_CORDE = re.compile(r'^(.*)_string_([IVX]+)$')

def technique_de_nom(chemin):
    """Technique lue dans le NOM DE FICHIER (corpus ZSamples et assimilés).
    'violin-con_sordino_string_I-B5-fff-1.wav' -> 'con_sordino_string_I'.
    Rend None si aucune note n'est trouvée (fichiers NoPitch)."""
    toks = os.path.basename(chemin)[:-4].split('-')
    for i, t in enumerate(toks):
        if note_vers_f0(t):
            return "-".join(toks[1:i]) if i > 1 else ""
    return None

def separer_corde(technique):
    """'ordinario_string_III' -> ('ordinario', 'III'). Sinon (technique, None).
    Permet le test d'indépendance inter-cordes : une même hauteur jouée sur deux
    cordes différentes, c'est deux sources pour un même corps. Si un formant est
    une résonance de caisse, il doit être identique d'une corde à l'autre."""
    m = _CORDE.match(technique or "")
    return (m.group(1), m.group(2)) if m else (technique, None)

def corde_de_nom(chemin):
    return separer_corde(technique_de_nom(chemin))[1]

def lister_par_nom(dossier, techniques_voulues, recursif=True):
    """Fichiers d'un dossier dont la technique, lue dans le NOM, appartient à
    `techniques_voulues` (le suffixe _string_N est ignoré pour la comparaison).
    Rend [(chemin, technique_normalisee, corde), ...]."""
    motif = os.path.join(dossier, "**", "*.wav") if recursif else os.path.join(dossier, "*.wav")
    out = []
    for f in sorted(glob.glob(motif, recursive=recursif)):
        t = technique_de_nom(f)
        if t is None:
            continue                      # NoPitch : écarté automatiquement
        base, corde = separer_corde(t)
        if base in techniques_voulues:
            out.append((f, base, corde))
    return out

def ecrire_spectres_liste(fichiers, sortie, racine=None):
    """Comme ecrire_spectres, à partir d'une liste explicite de chemins."""
    ok = ko = 0
    with open(sortie, "w") as fh:
        fh.write("spectrum 4096 2048 1024\n")
        for f in fichiers:
            a, err = spectre_moyen(f)
            if a is None:
                ko += 1
                continue
            rel = "/" + (os.path.relpath(f, racine) if racine else os.path.basename(f))
            fh.write(rel + ";" + ";".join(f"{v:.7g}" for v in a) + "\n")
            ok += 1
    return ok, ko, len(fichiers)

def ecrire_spectres(racine, dossier, techniques, sortie):
    """Écrit un fichier au format Orchidea `spectrum` : chemin;v1;…;v1024.
    Rend (n_ok, n_echecs, n_total)."""
    motifs = [os.path.join(racine, dossier, t, "*.wav") for t in techniques]
    fichiers = sorted(f for m in motifs for f in glob.glob(m))
    ok = ko = 0
    with open(sortie, "w") as fh:
        fh.write("spectrum 4096 2048 1024\n")
        for f in fichiers:
            a, err = spectre_moyen(f)
            if a is None:
                ko += 1
                continue
            rel = "/" + os.path.relpath(f, racine)
            fh.write(rel + ";" + ";".join(f"{v:.7g}" for v in a) + "\n")
            ok += 1
    return ok, ko, len(fichiers)

def lire_spectres(chemin):
    """Relit un fichier de spectres. Rend [(f0, midi, dyn, spectre), …]."""
    out = []
    with open(chemin, errors='ignore') as fh:
        fh.readline()
        for L in fh:
            if not L.startswith('/'):
                continue
            i = L.index(';')
            f0, midi, dyn = analyser_nom(L[:i])
            if f0 is None:
                continue
            v = np.fromstring(L[i + 1:], sep=';')
            if v.size >= NBINS:
                out.append((f0, midi, dyn, v[:NBINS]))
    return out

# ─────────────────── 2-3. harmoniques et résidus locaux ─────────────────────
def pic_proche(spectre, f, tol):
    b = int(round(f / BW))
    lo, hi = max(1, b - tol), min(NBINS, b + tol + 1)
    if lo >= hi:
        return None
    j = lo + int(np.argmax(spectre[lo:hi]))
    return j * BW, spectre[j]

# ───────────────────── résolution : les notes très graves ───────────────────
RESOL_MIN = 7.0          # bins d'écart minimal entre deux harmoniques (tâche C7)
#   `pic_proche` cherche le maximum dans ±2 bins autour de k·f0, soit une fenêtre de
#   5 bins. Quand f0/BW tombe sous cette largeur, les fenêtres de deux rangs voisins se
#   recouvrent et RENDENT LE MÊME PIC : la liste contient des fréquences dupliquées.
#   Mesuré sur les 6 920 notes de SOL : 0 % de notes touchées au-dessus de 7 bins,
#   28 % à 6, 51 % à 5, 74 % à 4, 100 % à 3 et moins. Soit un seuil à f0 >= 75 Hz.
#   Concerne 4,8 % du corpus, et jusqu'à 40 % des notes d'un instrument (clarinette
#   contrebasse, tuba contrebasse, contrebasson, contrebasses, sons pédale des cuivres).
#   CONSÉQUENCE MESURÉE : aucune. Dédoublonner laisse les formants en place (contrôle de
#   la tâche C7, sur neuf instruments ; voir `_REGISTRES-PEDALE.md`), parce que
#   `residu_local` retranche une médiane locale et qu'un point compté deux fois la
#   repondère sans la déplacer. Ne pas confondre avec le fait de RETIRER ces notes, qui
#   supprime le domaine grave et fait disparaître les formants bas : ce sont bien ces
#   notes-là qui les portent.
DEDOUBLONNER = os.environ.get('FORMANTS_DEDOUBLONNER') == '1'

def harmoniques(spectre, f0, kmax=90):
    """Amplitudes en dB des harmoniques, pic local autour de k·f0.

    Sous RESOL_MIN bins d'écart (f0 < 75 Hz) la liste peut contenir des fréquences
    dupliquées — voir ci-dessus. DEDOUBLONNER=1 les écarte ; par défaut non, pour que
    les résultats publiés restent reproductibles à l'identique."""
    fr, am = [], []
    for k in range(1, kmax + 1):
        f = k * f0
        if f > FMAX - 3 * BW:
            break
        r = pic_proche(spectre, f, 2 if k < 20 else 3)
        if r is None:
            break
        fr.append(r[0]); am.append(20 * math.log10(r[1] + 1e-12))
    fr, am = np.array(fr), np.array(am)
    if DEDOUBLONNER and fr.size:
        _, i = np.unique(fr, return_index=True)
        i = np.sort(i)
        fr, am = fr[i], am[i]
    return fr, am

def residu_local(fr, am, demi_fenetre=HALFWIN):
    """Retire à chaque harmonique la médiane locale des amplitudes de la MÊME
    note dans une fenêtre de ±demi_fenetre octave.

    Pourquoi local et non linéaire : un redressement linéaire en log-fréquence
    FABRIQUE une bosse médiane sur tout spectre concave. Il avait produit de
    faux formants à 1 318 et 1 863 Hz pour le hautbois, apparemment en accord
    avec la littérature. Ne pas y revenir.

    DEUX limites, à ne pas confondre — l'une a longtemps servi à expliquer des cas
    qui relevaient de l'autre (corrigé le 27/10/2026, tâche C2) :

    1. Largeur. Aveugle aux résonances plus larges que ~1,5 octave : une telle
       résonance est absorbée par la médiane locale. C'est bien le cas des flûtes,
       dont la zone d'énergie est une coupure de rayonnement et non une résonance.

    2. Plancher du domaine. Un harmonique n'entre dans le composite que si sa
       fenêtre contient au moins TROIS harmoniques de la même note (`q >= 3` dans
       `composite()`). Le nombre d'harmoniques dans [k/2^w, k·2^w] croît avec le
       rang k, donc les rangs bas sont écartés : à ±3/4 d'octave le rang 3 est le
       plus bas admis, et le plancher mesuré se situe entre 3,2 et 5,7 fois f0_min
       selon le corpus. **Rien n'est mesuré en dessous.** C'est cela — et non la
       largeur — qui fait manquer la caisse de l'alto (384 Hz, f0_min 131) et du
       violoncelle (219 Hz, f0_min 65). Voir `_FENETRE-DE-REDRESSEMENT.md`."""
    lf = np.log2(fr)
    res = np.empty_like(am)
    npts = np.empty(am.size, dtype=int)
    for i in range(am.size):
        m = np.abs(lf - lf[i]) <= demi_fenetre
        npts[i] = m.sum()
        res[i] = am[i] - np.median(am[m])
    return res, npts

def par_note(echantillons, demi_fenetre=HALFWIN):
    """[(freqs, residus, rangs, midi), …] — unité de base du bootstrap."""
    out = []
    for f0, midi, dyn, sp in echantillons:
        fr, am = harmoniques(sp, f0)
        if fr.size == 0:
            continue
        r, nl = residu_local(fr, am, demi_fenetre)
        g = nl >= 3
        if g.sum():
            out.append((fr[g], r[g], np.arange(1, fr.size + 1)[g], midi))
    return out

# ─────────────────── 4-5. enveloppe composite et candidats ──────────────────
def composite(echantillons, demi_fenetre=HALFWIN):
    """Méthode 8. Rend (mediane_par_bande, n_notes, n_points) sur la grille
    de 1/12 d'octave. Un formant = zone de résidu médian positif."""
    acc  = [[] for _ in range(NB)]
    vues = [set() for _ in range(NB)]
    for f0, midi, dyn, sp in echantillons:
        fr, am = harmoniques(sp, f0)
        if fr.size == 0:
            continue
        r, nl = residu_local(fr, am, demi_fenetre)
        idx = np.digitize(fr, BORDS) - 1
        for i, v, q in zip(idx, r, nl):
            if 0 <= i < NB and q >= 3:
                acc[i].append(v); vues[i].add(midi)
    med = np.array([np.median(a) if len(a) >= 5 else np.nan for a in acc])
    return med, np.array([len(s) for s in vues]), np.array([len(a) for a in acc])

def _lisser(y, w=2):
    o = np.full_like(y, np.nan)
    for i in range(len(y)):
        s = y[max(0, i - w):min(len(y), i + w + 1)]
        s = s[~np.isnan(s)]
        if s.size:
            o[i] = s.mean()
    return o

def trouver_formants(med, n_notes, n_pts, min_notes=5, min_pts=10,
                     min_residu=RESIDU_MIN, max_n=4, f_haut=6000.0):
    """Maxima locaux de l'enveloppe composite. Rend
    [(hz, residu, bande_bas, bande_haut, n_notes), …] triés par résidu."""
    s = _lisser(med)
    ok = (~np.isnan(s)) & (n_notes >= min_notes) & (n_pts >= min_pts) & (CENTRES <= f_haut)
    P = []
    for i in range(1, NB - 1):
        if not (ok[i] and ok[i - 1] and ok[i + 1]):
            continue
        if s[i] >= s[i - 1] and s[i] > s[i + 1] and s[i] >= min_residu:
            lo = hi = i
            while lo > 0 and ok[lo - 1] and s[lo - 1] > s[i] - 3:
                lo -= 1
            while hi < NB - 1 and ok[hi + 1] and s[hi + 1] > s[i] - 3:
                hi += 1
            P.append((CENTRES[i], s[i], CENTRES[lo], CENTRES[hi], int(n_notes[i])))
    P.sort(key=lambda p: -p[1])
    garde = []
    for p in P:                                   # fusionne à moins de 1/3 d'octave
        if all(abs(math.log2(p[0] / q[0])) > 1 / 3 for q in garde):
            garde.append(p)
    return garde[:max_n]

# ─────────────────────────── 6. tests de rejet ──────────────────────────────
def _rang_median(notes, fc, tol=1/6):
    ks = [k for fr, r, kk, m in notes for f, k in zip(fr, kk)
          if abs(math.log2(f / fc)) <= tol]
    return int(np.median(ks)) if ks else None

def _bootstrap_groupe(notes, fc, lo, hi, tol=1/6, rng=None):
    """Bootstrap PAR NOTE du résidu médian d'un groupe de rang.
    Par note et non par point : les harmoniques d'une même note sont corrélés."""
    rng = rng or np.random.default_rng(20261023)
    per = []
    for fr, r, ks, m in notes:
        s = (ks >= lo) & (ks <= hi) & (np.abs(np.log2(fr / fc)) <= tol)
        if s.any():
            per.append(r[s])
    npts = sum(len(x) for x in per)
    if len(per) < MIN_NOTES_GROUPE or npts < MIN_PTS_GROUPE:
        return dict(testable=False, n_notes=len(per), n_pts=npts)
    obs = float(np.median(np.concatenate(per)))
    idx = rng.integers(0, len(per), size=(NBOOT, len(per)))
    meds = np.array([np.median(np.concatenate([per[i] for i in idx[b]]))
                     for b in range(NBOOT)])
    ci = float(np.percentile(meds, 5))
    return dict(testable=True, n_notes=len(per), n_pts=npts,
                mediane=round(obs, 2), ic_bas=round(ci, 2),
                passe=bool(obs >= RESIDU_MIN and ci > 0))

def robustesse_fenetre(echantillons, fc, reference=None):
    """Test 2 : le pic se retrouve-t-il à moins de 1/3 d'octave quand la fenêtre
    de redressement s'écarte de ±DELTA_ROB octave de part et d'autre de la référence ?
    Ce test écarte les grandeurs fabriquées par leur propre fenêtre d'analyse,
    défaut du Fp par centroïde.

    Les deux fenêtres de comparaison SUIVENT la référence (27/10/2026, tâche C2).
    Elles étaient fixées à 0,5 et 1,0 : avec une référence à 1,0 le test comparait la
    référence à elle-même d'un côté, et de l'autre à une fenêtre dont le plancher de
    domaine est plus haut — ce qui rejetait d'office tout pic grave nouvellement
    visible."""
    w = HALFWIN if reference is None else reference
    for hw in (w - DELTA_ROB, w + DELTA_ROB):
        if hw <= 0:
            continue
        m, n, p = composite(echantillons, hw)
        pics = [a for a, b, c, d, e in trouver_formants(m, n, p)]
        if not any(abs(math.log2(x / fc)) <= 1 / 3 for x in pics):
            return False
    return True

def verdict(echantillons, notes, fc, rng=None):
    """Applique les deux tests. Rend un dictionnaire de verdict.

    Test 1 — indépendance au rang harmonique : la bande doit rester élevée quel
    que soit le rang qui l'éclaire. Découpage au RANG MÉDIAN (et non à seuil fixe,
    qui perd le test pour les instruments très graves dont tout le signal est à
    rang élevé). Les deux groupes doivent passer le bootstrap.

    Test 2 — robustesse de fenêtre.

    retenu-cr : les deux · fragile-cr : un seul · rejeté : aucun
    non mesurable : moins de deux groupes testables, aucun chiffre produit.
    Conserver le suffixe -cr dans toute publication."""
    km = _rang_median(notes, fc)
    if km is None:
        return dict(hz=round(fc, 1), verdict="non mesurable", groupes=[], rang_median=None)
    G = [_bootstrap_groupe(notes, fc, 1, km, rng=rng),
         _bootstrap_groupe(notes, fc, km + 1, 99, rng=rng)]
    rb = robustesse_fenetre(echantillons, fc)
    nt = sum(1 for g in G if g['testable'])
    npass = sum(1 for g in G if g.get('passe'))
    if nt < 2:
        v = "non mesurable"
    elif npass == 2 and rb:
        v = "retenu-cr"
    elif (npass == 2) != rb:
        v = "fragile-cr"
    else:
        v = "rejeté"
    return dict(hz=round(fc, 1), verdict=v, robustesse=rb, rang_median=km, groupes=G)

# ─────────────────── 7. méthode 6 : enveloppe vraie adaptative ──────────────
def enveloppe_vraie(spectre, f0, iters=20):
    """Cepstre itératif à la Röbel & Rodet, ordre ADAPTÉ à f0.
    Un ordre fixe crée une limite fixe en hauteur : c'est le défaut des specenv
    d'Orchidea (NUM_SMOOTH=160, valide seulement sous 276 Hz)."""
    C = max(3, int(0.7 * QCOMB * BW / f0))
    lg = np.log(np.maximum(spectre, 1e-12))
    V = lg.copy()
    for _ in range(iters):
        V = np.maximum(lg, V)
        c = np.fft.irfft(V)
        c[C:len(c) - C + 1] = 0.0
        V = np.fft.rfft(c).real[:NBINS]
    return V, C

def pic_enveloppe_vraie(spectre, f0, f_haut=6000.0):
    env, C = enveloppe_vraie(spectre, f0)
    seg = env[3:int(f_haut / BW)]
    cand = [i for i in range(1, len(seg) - 1)
            if seg[i] >= seg[i - 1] and seg[i] > seg[i + 1]]
    if not cand:
        return float('nan'), C
    return (max(cand, key=lambda i: seg[i]) + 3) * BW, C

# ─────────────────────────── descripteurs globaux ───────────────────────────
def spearman(a, b):
    ra = np.argsort(np.argsort(a)).astype(float)
    rb = np.argsort(np.argsort(b)).astype(float)
    ra -= ra.mean(); rb -= rb.mean()
    d = math.sqrt((ra ** 2).sum() * (rb ** 2).sum())
    return float((ra * rb).sum() / d) if d else float('nan')

def descripteurs(echantillons):
    """rho, pente, balance impair-pair, enveloppe vraie, etc.

    rho = Spearman(pic d'enveloppe vraie, f0). ATTENTION : rho ne doit JAMAIS
    être publié seul comme indice d'absence de formant. Un rho élevé a deux
    causes opposées — aucune résonance étroite (flûtes) ou beaucoup de résonances
    dont aucune ne domine (cordes). C'est le couple (rho, nombre de formants
    retenus) qui classe. Voir la typologie à trois classes dans 00-METHODE.md."""
    f0s = np.array([x[0] for x in echantillons])
    tv, ordres, pentes, nh, odd = [], [], [], [], []
    for f0, midi, dyn, sp in echantillons:
        v, C = pic_enveloppe_vraie(sp, f0)
        tv.append(v); ordres.append(C)
        fr, am = harmoniques(sp, f0)
        nh.append(fr.size)
        if fr.size >= 5:
            x = np.log2(fr); A = np.vstack([x, np.ones_like(x)]).T
            pentes.append(np.linalg.lstsq(A, am, rcond=None)[0][0])
        if fr.size >= 9:
            odd.append(np.mean([am[k] for k in (2, 4, 6, 8)]) -
                       np.mean([am[k] for k in (1, 3, 5, 7)]))
    tv = np.array(tv, float); g = ~np.isnan(tv)
    sur_harm = float(np.mean([
        min(abs(math.log2(t / (k * f))) for k in (1, 2, 3, 4)) < 0.04
        for t, f in zip(tv[g], f0s[g])])) if g.sum() else float('nan')
    return dict(
        n=len(echantillons), n_notes=len({x[1] for x in echantillons}),
        f0_min=round(float(f0s.min()), 1), f0_max=round(float(f0s.max()), 1),
        rho=round(spearman(f0s[g], tv[g]), 2) if g.sum() > 5 else None,
        pente=round(float(np.median(pentes)), 1) if pentes else None,
        pente_sd=round(float(np.std(pentes)), 1) if pentes else None,
        harmoniques_median=int(np.median(nh)),
        harmoniques_min=int(min(nh)), harmoniques_max=int(max(nh)),
        balance_impair_pair=round(float(np.median(odd)), 1) if odd else None,
        env_vraie=round(float(np.nanmedian(tv)), 0),
        env_vraie_sd=round(float(np.nanstd(tv)), 0),
        ordre_min=int(min(ordres)), ordre_max=int(max(ordres)),
        pics_sur_harmonique=round(sur_harm, 3))

def f0_par_somme_harmonique(spectre, f0_annonce, demi_plage_cents=150, pas=5, kmax=12):
    """Estime f0 en maximisant la somme des amplitudes aux k·f, sur une plage
    autour de f0 annoncé.

    Pourquoi pas le pic le plus proche de f0 : les clarinettes ont une
    fondamentale FAIBLE — c'est leur signature. Un contrôle fondé sur le pic
    voisin de f0 y produit de faux positifs (constaté sur ZSamples : rapport
    annoncé 1,118 alors que la nomenclature était juste)."""
    meilleur, score_max = None, -1.0
    for cents in range(-demi_plage_cents, demi_plage_cents + 1, pas):
        f = f0_annonce * 2 ** (cents / 1200.0)
        s = 0.0
        for k in range(1, kmax + 1):
            fk = k * f
            if fk > FMAX:
                break
            b = int(round(fk / BW))
            if 1 <= b < NBINS:
                s += spectre[b]
        if s > score_max:
            meilleur, score_max = f, s
    return meilleur

def controle_nomenclature(echantillons):
    """Rapport médian entre le f0 réel (somme harmonique) et le f0 déduit du nom
    de fichier. Doit valoir ~1,00. À REFAIRE pour toute nouvelle base :
    s'il vaut 0,89, les noms sont en hauteurs écrites et non réelles.
    La plage de recherche étant de ±150 cents, un écart supérieur sature à
    ~0,92 / 1,09 : un rapport collé à une borne signale une transposition
    franche, à vérifier à la main."""
    r = [f0_par_somme_harmonique(sp, f0) / f0 for f0, m, d, sp in echantillons]
    return round(float(np.median(r)), 3) if r else None

# ─────────────────────────── chaîne complète ────────────────────────────────
def mesurer(chemin_spectres, graine=20261023):
    """Analyse complète d'un fichier de spectres. Rend un dictionnaire prêt à
    sérialiser en JSON et à verser dans l'en-tête YAML d'une fiche."""
    rng = np.random.default_rng(graine)
    ech = lire_spectres(chemin_spectres)
    if len(ech) < 5:
        return dict(erreur=f"{len(ech)} échantillon(s) seulement")
    med, nn, npt = composite(ech)
    notes = par_note(ech)
    cands = trouver_formants(med, nn, npt)
    res = descripteurs(ech)
    res['nomenclature'] = controle_nomenclature(ech)
    res['formants'] = []
    for fc, amp, blo, bhi, nnd in cands:
        v = verdict(ech, notes, fc, rng=rng)
        v.update(residu=round(amp, 2), bande_bas=round(blo, 1), bande_haut=round(bhi, 1),
                 n_notes=nnd)
        res['formants'].append(v)
    res['enveloppe'] = dict(centres=[round(c, 1) for c in CENTRES],
                            mediane=[None if np.isnan(v) else round(float(v), 2) for v in med],
                            n_notes=[int(v) for v in nn], n_points=[int(v) for v in npt])
    return res
