#!/usr/bin/env python3
"""
split_specenv_by_instrument.py
Découpe un gros fichier spectrum en fichiers séparés par instrument.

L'instrument est extrait du 2e élément des lignes de chemin, ex :
  /Brass/Bass_Tuba/bisbigliando/BTb-bisb-A#1-mf-N-
  → instrument = "Bass_Tuba"

Chaque fichier résultant commence par la ligne d'en-tête :
  spectrum 4096 2048 1024

Usage :
  python split_spectrum_by_instrument.py INPUT_FILE [OUTPUT_DIR]

  INPUT_FILE : fichier spectrum source (peut faire >200 MB)
  OUTPUT_DIR : dossier de sortie (défaut : INPUT_FILE_par_instrument/)
"""

import sys
import os
import re
from pathlib import Path

HEADER = "spectrum 4096 2048 1024"

# Pattern pour les lignes de chemin : /Family/Instrument/technique/...
PATH_PATTERN = re.compile(r'^/([^/]+)/([^/]+)/')


def extract_instrument(line):
    """Extrait le nom d'instrument (2e élément) d'une ligne de chemin."""
    m = PATH_PATTERN.match(line)
    if m:
        return m.group(2)
    return None


def split_file(input_path, output_dir):
    input_path = Path(input_path)
    if not input_path.exists():
        print(f"ERREUR : fichier introuvable : {input_path}")
        sys.exit(1)

    if output_dir is None:
        output_dir = input_path.parent / f"{input_path.stem}_par_instrument"
    else:
        output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    # État courant
    current_instrument = None
    current_file = None
    open_files = {}  # instrument -> file handle
    counts = {}      # instrument -> nombre de samples

    file_size = input_path.stat().st_size
    bytes_read = 0
    last_pct = -1

    print(f"Lecture de {input_path} ({file_size / 1e6:.0f} MB)...")
    print(f"Sortie dans {output_dir}/\n")

    with open(input_path, 'r', encoding='utf-8', errors='replace') as f:
        for line in f:
            bytes_read += len(line.encode('utf-8', errors='replace'))

            # Progression
            pct = int(bytes_read * 100 / file_size) if file_size > 0 else 0
            if pct >= last_pct + 5:
                last_pct = pct
                print(f"  {pct}%  ({len(open_files)} instruments trouvés)", end='\r')

            stripped = line.strip()

            # Ignorer la ligne d'en-tête globale
            if stripped == HEADER:
                continue

            # Ignorer les lignes vides
            if not stripped:
                continue

            # Ligne de chemin → nouvel échantillon
            instrument = extract_instrument(stripped)
            if instrument:
                current_instrument = instrument

                if instrument not in open_files:
                    out_path = output_dir / f"{instrument}_spectrum.txt"
                    fh = open(out_path, 'w', encoding='utf-8')
                    fh.write(HEADER + '\n')
                    open_files[instrument] = fh
                    counts[instrument] = 0

                current_file = open_files[instrument]
                counts[instrument] += 1
                current_file.write(line)
                continue

            # Ligne de données → écrire dans le fichier courant
            if current_file is not None:
                current_file.write(line)

    # Fermer tous les fichiers
    for fh in open_files.values():
        fh.close()

    # Résumé
    print(f"\n\nTerminé ! {len(counts)} instruments extraits :\n")
    for instr in sorted(counts.keys()):
        out_name = f"{instr}_spectrum.txt"
        out_size = (output_dir / out_name).stat().st_size / 1e6
        print(f"  {instr:30s}  {counts[instr]:5d} samples  ({out_size:.1f} MB)")

    total = sum(counts.values())
    print(f"\n  Total : {total} samples")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_file = sys.argv[1]
    output_directory = sys.argv[2] if len(sys.argv) > 2 else None
    split_file(input_file, output_directory)
