#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
batch_interference.py
---------------------
Traitement par lot de rapports CHIRplus_BC (TXT).
Pour chaque fichier, évalue le risque de brouillage en fonction d'un seuil ENU
et produit un classeur Excel contenant :
  - un onglet "Résumé" par site (verdict, nombre d'émetteurs > seuil, pire cas)
  - un onglet "Interférences" listant tous les émetteurs > seuil pour tous les sites

Usage :
    python batch_interference.py dossier_ou_fichiers --threshold 60 --out resultats.xlsx
    # Exemples :
    python batch_interference.py /data/rapports --threshold 60 --out interferences.xlsx
    python batch_interference.py r1.txt r2.txt r3.txt --threshold 55 --out batch.xlsx
"""
import argparse
import re
from pathlib import Path
from typing import List, Dict, Any, Tuple
import pandas as pd

# =============================
# Parsing helpers (repris/alignés avec check_interference.py)
# =============================
def _find_table_lines(lines: List[str]) -> Tuple[int, int]:
    header_idx = -1
    pattern = re.compile(r'^\s*ENU\s+OS\s+TRANSMITTER', re.IGNORECASE)
    for i, ln in enumerate(lines):
        if pattern.search(ln):
            header_idx = i
            break
    if header_idx == -1:
        raise RuntimeError("En-tête de table non trouvé (ligne qui commence par 'ENU  OS  TRANSMITTER').")
    start = header_idx
    end = len(lines)
    for j in range(header_idx+1, len(lines)):
        if lines[j].strip() == "" and (j+1 < len(lines) and lines[j+1].strip() == ""):
            end = j
            break
    return start, end

def _column_slices(header: str):
    starts = []
    h = header.rstrip("\n")
    for i, ch in enumerate(h):
        if (i == 0 and ch != " ") or (i > 0 and ch != " " and h[i-1] == " "):
            starts.append(i)
    ends = starts[1:] + [len(h)]
    cols = []
    for s, e in zip(starts, ends):
        name = h[s:e].strip()
        if name:
            cols.append((name, s, e))
    return cols

def _parse_fixed_width_row(line: str, cols):
    row = {}
    for name, s, e in cols:
        row[name] = line[s:e].strip()
    return row

def parse_table_from_txt(txt: str) -> List[Dict[str, Any]]:
    lines = txt.splitlines()
    start, end = _find_table_lines(lines)
    header = lines[start]
    cols = _column_slices(header)
    rows = []
    for ln in lines[start+1:end]:
        if not ln.strip():
            continue
        if not re.match(r'^\s*\d', ln):
            continue
        rows.append(_parse_fixed_width_row(ln, cols))
    return rows

def extract_site_name(txt: str, default_name: str) -> str:
    # Essayer de récupérer "Interf. Transmit.: POZO NIVES" ou variantes
    patterns = [
        r'Interf\.\s*Transmit\.?\s*:\s*(.+)',
        r'Interfer\.\s*Transmit\.?\s*:\s*(.+)',
        r'Interf.*Transmit.*:\s*(.+)',
    ]
    for pat in patterns:
        m = re.search(pat, txt, flags=re.IGNORECASE)
        if m:
            name = m.group(1).strip()
            # couper à la fin de ligne si reste du texte
            name = name.splitlines()[0].strip()
            return name
    # sinon, essayer "Filename" (chemin) ou utiliser le nom de fichier par défaut
    m = re.search(r'Filename\s*:\s*(.+)', txt, flags=re.IGNORECASE)
    if m:
        candidate = Path(m.group(1).strip()).stem
        if candidate:
            return candidate
    return default_name

def rows_above_threshold(rows: List[Dict[str, Any]], threshold: float) -> List[Dict[str, Any]]:
    out = []
    for r in rows:
        enu_str = r.get('ENU', '')
        try:
            enu = float(enu_str)
        except Exception:
            mm = re.findall(r'[-+]?\d+(?:\.\d+)?', enu_str)
            if not mm:
                continue
            enu = float(mm[0])
        if enu > threshold:
            r = dict(r)  # shallow copy
            r['_ENU_float'] = enu
            out.append(r)
    out.sort(key=lambda x: x['_ENU_float'], reverse=True)
    return out

# =============================
# Batch processing
# =============================
def expand_inputs(paths: List[str]) -> List[Path]:
    files: List[Path] = []
    for p in paths:
        P = Path(p)
        if P.is_dir():
            files.extend(sorted([f for f in P.rglob("*.txt")]))
        elif P.is_file():
            files.append(P)
    # unique, garder l'ordre
    seen = set()
    unique_files = []
    for f in files:
        if f.resolve() not in seen:
            unique_files.append(f)
            seen.add(f.resolve())
    return unique_files

def process_file(file: Path, threshold: float) -> Dict[str, Any]:
    txt = file.read_text(encoding='utf-8', errors='ignore')
    site = extract_site_name(txt, default_name=file.stem)
    table_rows = parse_table_from_txt(txt)
    interfs = rows_above_threshold(table_rows, threshold)
    # Résumé
    risk = len(interfs) > 0
    worst = interfs[0] if risk else None
    summary = {
        'site': site,
        'file': str(file),
        'threshold_ENU_dB': threshold,
        'risk': 'OUI' if risk else 'NON',
        'interferer_count': len(interfs),
        'max_ENU': worst['_ENU_float'] if worst else None,
        'worst_transmitter': worst['TRANSMITTER'] if worst else None,
    }
    # détaillé
    for r in interfs:
        r['site'] = site
        r['file'] = str(file)
    return {'summary': summary, 'details': interfs}

def main():
    ap = argparse.ArgumentParser(description="Traitement par lot de rapports CHIRplus_BC pour évaluer le brouillage (seuil ENU).")
    ap.add_argument('inputs', nargs='+', help="Dossier(s) et/ou fichier(s) TXT à traiter.")
    ap.add_argument('--threshold', '-t', type=float, required=True, help="Seuil ENU en dB.")
    ap.add_argument('--out', '-o', type=Path, required=True, help="Chemin du fichier Excel de sortie (.xlsx).")
    args = ap.parse_args()

    files = expand_inputs(args.inputs)
    if not files:
        raise SystemExit("Aucun fichier .txt trouvé dans les entrées.")
    summaries = []
    all_details = []
    for f in files:
        try:
            res = process_file(f, args.threshold)
            summaries.append(res['summary'])
            all_details.extend(res['details'])
        except Exception as e:
            summaries.append({
                'site': f.stem,
                'file': str(f),
                'threshold_ENU_dB': args.threshold,
                'risk': 'ERREUR',
                'interferer_count': None,
                'max_ENU': None,
                'worst_transmitter': str(e),
            })

    # DataFrames
    df_sum = pd.DataFrame(summaries).sort_values(['risk','interferer_count','max_ENU'], ascending=[False, False, False])
    df_det = pd.DataFrame(all_details)

    # Export Excel
    out_path = args.out
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        df_sum.to_excel(writer, sheet_name='Résumé', index=False)
        if not df_det.empty:
            # Colonnes utiles en tête si présentes
            preferred = ['site','file','ENU','OS','TRANSMITTER','DIS','AZM','LONGITUDE','LATITUDE','ERP','f/MHz','CHA','HEFF','POL','PROGRAM','REMARKS','_ENU_float']
            cols = [c for c in preferred if c in df_det.columns] + [c for c in df_det.columns if c not in preferred]
            df_det[cols].to_excel(writer, sheet_name='Interférences', index=False)
        else:
            # feuille vide avec en-tête
            pd.DataFrame(columns=['site','file','TRANSMITTER','ENU']).to_excel(writer, sheet_name='Interférences', index=False)

    print(f"[OK] Traitement terminé. Fichiers analysés : {len(files)}")
    print(f"Classeur Excel écrit : {out_path}")

if __name__ == "__main__":
    main()
