#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_interference.py
---------------------
Lit un rapport CHIRplus_BC (TXT) et vérifie le risque de brouillage
en fonction d'un seuil ENU fourni par l'utilisateur. Liste les émetteurs
dont l'ENU dépasse le seuil.

Usage:
    python check_interference.py /chemin/vers/rapport.txt --threshold 60
    # Optionnel: exporter les résultats en CSV
    python check_interference.py rapport.txt --threshold 60 --out interferers.csv
"""
import argparse
import re
from typing import List, Tuple, Dict, Any
import csv
import sys
from pathlib import Path

def _find_table_lines(lines: List[str]) -> Tuple[int, int]:
    """
    Trouve la table qui commence par l'en-tête 'ENU  OS  TRANSMITTER ...'
    Retourne (start_idx, end_idx_exclusive) des lignes de la table.
    """
    header_idx = -1
    pattern = re.compile(r'^\s*ENU\s+OS\s+TRANSMITTER', re.IGNORECASE)
    for i, ln in enumerate(lines):
        if pattern.search(ln):
            header_idx = i
            break
    if header_idx == -1:
        raise RuntimeError("En-tête de table non trouvé (ligne qui commence par 'ENU  OS  TRANSMITTER').")
    # La table continue jusqu'à la première ligne vide consécutive ou fin de fichier
    start = header_idx
    end = len(lines)
    for j in range(header_idx+1, len(lines)):
        # Table se termine si on tombe sur une ligne entièrement vide
        if lines[j].strip() == "" and (j+1 < len(lines) and lines[j+1].strip() == ""):
            end = j
            break
    return start, end

def _column_slices(header: str) -> List[Tuple[str, int, int]]:
    """
    Déduit les colonnes (noms + tranches) à partir de l'en-tête aligné.
    On détecte toutes les transitions 'espace -> non-espace' comme début de colonne.
    """
    starts = []
    for i, ch in enumerate(header.rstrip("\n")):
        if (i == 0 and ch != " ") or (i > 0 and ch != " " and header[i-1] == " "):
            starts.append(i)
    # bornes de fin = début de la colonne suivante
    ends = starts[1:] + [len(header.rstrip("\n"))]
    cols = []
    for s, e in zip(starts, ends):
        name = header[s:e].strip()
        # ignorer colonnes vides
        if name:
            cols.append((name, s, e))
    return cols

def _parse_fixed_width_row(line: str, cols: List[Tuple[str, int, int]]) -> Dict[str, str]:
    row: Dict[str, str] = {}
    for name, s, e in cols:
        val = line[s:e].strip()
        row[name] = val
    return row

def _normalize_key(k: str) -> str:
    k = re.sub(r'[^0-9A-Za-z]+', '_', k.strip())
    return k.strip('_')

def parse_report(path: Path) -> List[Dict[str, Any]]:
    txt = path.read_text(encoding="utf-8", errors="ignore")
    lines = txt.splitlines()
    start, end = _find_table_lines(lines)
    header = lines[start]
    data_lines = []
    # données: à partir de la ligne suivant l'en-tête, ignorer éventuelles lignes de séparation
    for ln in lines[start+1:end]:
        if not ln.strip():
            # ligne vide simple autorisée: on continue jusqu'à une seconde vide consécutive détectée dans _find_table_lines
            continue
        # ignorer lignes d'annotation qui ne commencent pas par chiffre de ENU
        if not re.match(r'^\s*\d', ln):
            continue
        data_lines.append(ln)
    # colonnes à partir de l'en-tête
    cols = _column_slices(header)
    rows = []
    for ln in data_lines:
        row = _parse_fixed_width_row(ln, cols)
        # Ajout des clés normalisées en parallèle
        norm_row = { _normalize_key(k): v for k, v in row.items() }
        # pour commodité, garder aussi noms originaux
        norm_row['__raw__'] = row
        rows.append(norm_row)
    return rows

def main():
    ap = argparse.ArgumentParser(description="Vérifie le risque de brouillage (ENU) dans un rapport CHIRplus_BC.")
    ap.add_argument('file', type=Path, help="Chemin du rapport TXT")
    ap.add_argument('--threshold', '-t', type=float, required=True, help="Seuil ENU en dB au-delà duquel on considère un brouillage.")
    ap.add_argument('--out', '-o', type=Path, default=None, help="Chemin de sortie CSV pour la liste des émetteurs dépassant le seuil.")
    args = ap.parse_args()

    try:
        rows = parse_report(args.file)
    except Exception as e:
        print(f"[ERREUR] {e}", file=sys.stderr)
        sys.exit(2)

    # Tenter d'interpréter ENU en float et filtrer
    interferers = []
    for r in rows:
        enu_str = r.get('ENU') or r.get('Enu') or r.get('enu') or r['__raw__'].get('ENU', '')
        try:
            enu = float(enu_str)
        except Exception:
            # certaines lignes peuvent avoir des caractères parasites
            try:
                enu = float(re.findall(r'[-+]?\d+(?:\.\d+)?', enu_str)[0])
            except Exception:
                continue
        if enu > args.threshold:
            r['_ENU_float'] = enu
            interferers.append(r)

    # Tri décroissant par ENU
    interferers.sort(key=lambda x: x.get('_ENU_float', -1.0), reverse=True)

    # Verdict global
    risk = "OUI" if len(interferers) > 0 else "NON"
    print(f"Risque de brouillage selon seuil ENU = {args.threshold:.1f} dB : {risk}")
    print(f"Nombre d'émetteurs dépassant le seuil : {len(interferers)}")
    if interferers:
        # Choisir quelques colonnes clés si présentes
        cols_pref = ['ENU', 'OS', 'TRANSMITTER', 'DIS', 'AZM', 'LONGITUDE', 'LATITUDE', 'ERP', 'f_MHz', 'CHA', 'HEFF', 'POL', 'PROGRAM', 'REMARKS']
        # reconstruire à partir de __raw__
        header = [c for c in cols_pref if c in interferers[0]['__raw__'].keys()]
        # Impression terminal
        print("\nÉmetteurs au-dessus du seuil (triés par ENU décroissant):")
        print(" | ".join(header))
        for r in interferers:
            raw = r['__raw__']
            line = " | ".join(raw.get(h, '') for h in header)
            print(line)
        # Export CSV si demandé
        if args.out is not None:
            with args.out.open('w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(header)
                for r in interferers:
                    raw = r['__raw__']
                    writer.writerow([raw.get(h, '') for h in header])
            print(f"\n[OK] Résultats exportés vers: {args.out}")
    else:
        print("Aucun émetteur ne dépasse le seuil fourni.")

if __name__ == "__main__":
    main()
