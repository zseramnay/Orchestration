#!/usr/bin/env python3
"""Migration CSV vers v2: renommage des instruments en libelles FR."""

import csv
import os
from collections import Counter


NAME_MAPPING = {
    'Accordion': 'Accordéon',
    'Bass_Clarinet_Bb': 'Clarinette basse en Sib',
    'Bass_Flute': 'Flûte basse',
    'Bass_Trombone': 'Trombone basse',
    'Bass_Tuba': 'Tuba basse',
    'Bass_Tuba+sordina': 'Tuba basse+sourdine',
    'Bassoon': 'Basson',
    'Bassoon+sordina': 'Basson+sourdine',
    'Clarinet_Bb': 'Clarinette en Sib',
    'Clarinet_Eb': 'Clarinette en Mib',
    'Contrabass': 'Contrebasse',
    'Contrabass+sordina': 'Contrebasse+sourdine',
    'Contrabass_Clarinet_Bb': 'Clarinette contrebasse en Sib',
    'Contrabass_Ensemble': 'Ensemble de contrebasses',
    'Contrabass_Flute': 'Flûte contrebasse',
    'Contrabass_Tuba': 'Tuba contrebasse',
    'Contrabassoon': 'Contrebasson',
    'Cow_Bell': 'Cloche de vache',
    'English_Horn': 'Cor anglais',
    'Flute': 'Flûte',
    'Gongs_Thai': 'Gongs thai',
    'Guitar': 'Guitare',
    'Harp': 'Harpe',
    'Horn': 'Cor',
    'Horn+sordina': 'Cor+sourdine',
    'Marimba': 'Marimba',
    'Oboe': 'Hautbois',
    'Oboe+sordina': 'Hautbois+sourdine',
    'Piccolo': 'Petite flûte',
    'Sax_Alto': 'Saxophone alto',
    'Trombone': 'Trombone',
    'Trombone+sordina_cup': 'Trombone+sourdine bol',
    'Trombone+sordina_harmon': 'Trombone+sourdine harmon',
    'Trombone+sordina_straight': 'Trombone+sourdine sèche',
    'Trombone+sordina_wah': 'Trombone+sourdine wah',
    'Trumpet_C': 'Trompette en Ut',
    'Trumpet_C+sordina_cup': 'Trompette en Ut+sourdine bol',
    'Trumpet_C+sordina_harmon': 'Trompette en Ut+sourdine harmon',
    'Trumpet_C+sordina_straight': 'Trompette en Ut+sourdine sèche',
    'Trumpet_C+sordina_wah': 'Trompette en Ut+sourdine wah',
    'Tubular_Bells': 'Cloches tubulaires',
    'Vibraphone': 'Vibraphone',
    'Viola': 'Alto',
    'Viola+sordina': 'Alto+sourdine',
    'Viola+sordina_piombo': 'Alto+sourdine piombo',
    'Viola_Ensemble': 'Ensemble d altos',
    'Viola_Ensemble+sordina': 'Ensemble d altos+sourdine',
    'Violin': 'Violon',
    'Violin+sordina': 'Violon+sourdine',
    'Violin+sordina_piombo': 'Violon+sourdine piombo',
    'Violin_Ensemble': 'Ensemble de violons',
    'Violin_Ensemble+sordina': 'Ensemble de violons+sourdine',
    'Violoncello': 'Violoncelle',
    'Violoncello+sordina': 'Violoncelle+sourdine',
    'Violoncello+sordina_piombo': 'Violoncelle+sourdine piombo',
    'Violoncello_Ensemble': 'Ensemble de violoncelles',
    'Violoncello_Ensemble+sordina': 'Ensemble de violoncelles+sourdine',
}

TARGET_FILES = [
    'Resultats/formants_all_techniques.csv',
    'Resultats/formants_yan_adds.csv',
]


def migrate_csv(input_csv, output_csv):
    print(f"Migration: {input_csv} -> {output_csv}")
    renamed_count = 0
    seen_instruments = Counter()
    unmapped = set()

    with open(input_csv, 'r', encoding='utf-8', newline='') as fin:
        reader = csv.DictReader(fin)
        fieldnames = list(reader.fieldnames or [])
        rows = []

        for row in reader:
            new_row = row.copy()

            for col in ('instrument', "Nom d'instrument"):
                old_name = new_row.get(col)
                if not old_name:
                    continue
                seen_instruments[old_name] += 1
                new_name = NAME_MAPPING.get(old_name)
                if new_name:
                    if new_name != old_name:
                        renamed_count += 1
                    new_row[col] = new_name
                else:
                    unmapped.add(old_name)

            rows.append(new_row)

    with open(output_csv, 'w', encoding='utf-8', newline='') as fout:
        writer = csv.DictWriter(fout, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"  Lignes traitees: {len(rows)}")
    print(f"  Remplacements: {renamed_count}")
    print(f"  Instruments uniques: {len(seen_instruments)}")
    if unmapped:
        print(f"  Non mappes ({len(unmapped)}): {', '.join(sorted(unmapped))}")
    print(f"  OK: {output_csv}")


def migrate_targets():
    for input_path in TARGET_FILES:
        if not os.path.exists(input_path):
            print(f"Fichier introuvable: {input_path}")
            continue
        output_path = input_path.replace('.csv', '_v2.csv')
        migrate_csv(input_path, output_path)
        print()


if __name__ == '__main__':
    print('=== Migration CSV vers v2 ===')
    migrate_targets()
    print('=== Migration terminee ===')

