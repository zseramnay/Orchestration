#!/usr/bin/env python3
"""Recalcule le spectre moyen par echantillon en EXCLUANT le transitoire d'attaque.
Sortie : meme format que les fichiers Orchidea `spectrum` (chemin;1024 valeurs)."""
import sys, os, glob
import numpy as np, soundfile as sf

NFFT, HOP, NBINS, SR = 4096, 2048, 1024, 44100
SKIP_MS, MAXDUR_MS, MINDUR_MS = 200, 1500, 250
WIN = np.hanning(NFFT)

def avg_spectrum(path):
    x, sr = sf.read(path, always_2d=True)
    x = x.mean(axis=1).astype(np.float64)
    if sr != SR: return None, f"sr={sr}"
    # debut : 200 ms apres l'onset (premier depassement de 20% du max RMS)
    env = np.abs(x)
    k = max(1, int(sr*0.01))
    rms = np.convolve(env, np.ones(k)/k, mode='same')
    thr = 0.2*rms.max()
    idx = np.nonzero(rms > thr)[0]
    onset = idx[0] if idx.size else 0
    start = onset + int(sr*SKIP_MS/1000)
    end   = min(len(x), start + int(sr*MAXDUR_MS/1000))
    if end - start < int(sr*MINDUR_MS/1000):
        start = onset; end = len(x)
        if end - start < NFFT: return None, "trop court"
    seg = x[start:end]
    frames = []
    for i in range(0, len(seg)-NFFT+1, HOP):
        frames.append(np.abs(np.fft.rfft(seg[i:i+NFFT]*WIN, n=NFFT)))
    if not frames:
        pad = np.zeros(NFFT); pad[:len(seg)] = seg[:NFFT]*WIN[:len(seg)]
        frames = [np.abs(np.fft.rfft(pad, n=NFFT))]
    avg = np.mean(frames, axis=0)[:NBINS]
    s = avg.sum()
    if s > 0: avg = avg / s          # normalisation par l'energie totale
    return avg, None

def run(root, instr_dir, out_path, techniques):
    pats = [os.path.join(root, instr_dir, t, "*.wav") for t in techniques]
    files = sorted(f for p in pats for f in glob.glob(p))
    nok = nerr = 0
    with open(out_path, "w") as fh:
        fh.write("spectrum 4096 2048 1024\n")
        for f in files:
            a, err = avg_spectrum(f)
            if a is None:
                nerr += 1; continue
            rel = "/" + os.path.relpath(f, root)
            fh.write(rel + ";" + ";".join(f"{v:.7g}" for v in a) + "\n")
            nok += 1
    return nok, nerr, len(files)

if __name__ == "__main__":
    root, instr_dir, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    techs = sys.argv[4].split(",")
    print(run(root, instr_dir, out_path, techs))
