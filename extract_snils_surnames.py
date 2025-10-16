#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Walks a directory tree, writes a file list (structure), and extracts pairs
SNILS + Surname into a txt file.

Sources for surname per SNILS:
- Filenames: first Cyrillic word that looks like a surname (e.g., "Жукова", "ОСИПОВА").
- Excel file named like "Персональные данные.xlsx|xslx" (B2 cell, first word as surname).

Output examples (one per line, deduplicated):
00899112570 Жукова
00899112570 Шмаль
05327487162 Карасева

Usage:
  python scripts/extract_snils_surnames.py [ROOT_DIR] [--paths OUT1] [--out OUT2]

Defaults:
- ROOT_DIR = current working directory
- OUT1 (all paths listing) = all_paths.txt in ROOT_DIR
- OUT2 (result pairs) = snils_surnames.txt in ROOT_DIR

Notes:
- Attempts to use openpyxl to read B2 from Excel. If not installed, Excel
  extraction is skipped with a console notice.
- SNILS is detected as any path segment of exactly 11 digits; if not found in
  segments, the script searches 11-digit sequences in filenames.
"""

from __future__ import annotations

import os
import re
import sys
import argparse
from typing import Dict, Iterable, List, Optional, Set, Tuple


try:
    import openpyxl  # type: ignore
except Exception:  # pragma: no cover
    openpyxl = None


SNILS_RE = re.compile(r"^\d{11}$")
SNILS_INLINE_RE = re.compile(r"(?<!\d)(\d{11})(?!\d)")

# Heuristics for Cyrillic surname candidates
CYRILLIC_WORD = re.compile(r"^[А-ЯЁа-яё-]{2,}$")
CYRILLIC_CAP_UPPER = re.compile(r"^[А-ЯЁ-]{3,}$")  # e.g., ОСИПОВА
INIT_TOKEN = re.compile(r"^[А-ЯЁ]\.(?:[А-ЯЁ]\.)?$")  # e.g., С.А or Н.


def find_snils_from_path(path: str) -> Optional[str]:
    parts = os.path.normpath(path).split(os.sep)
    # Prefer a directory segment with 11 digits (deepest first)
    for seg in reversed(parts[:-1]):  # exclude filename
        if SNILS_RE.match(seg):
            return seg
    # Fallback: look inside filename
    m = SNILS_INLINE_RE.search(os.path.basename(path))
    if m:
        return m.group(1)
    return None


def tokens_from_name_string(s: str) -> List[str]:
    # Normalize: replace common separators with spaces, collapse multiple spaces
    cleaned = re.sub(r"[\s_]+", " ", s.replace("-", "-")).strip()
    if not cleaned:
        return []
    return cleaned.split(" ")


def looks_like_name_word(tok: str) -> bool:
    # Likely a name-like word: Cyrillic, at least 2 letters (allow hyphen)
    return bool(CYRILLIC_WORD.match(tok))


def extract_surname_candidates_from_text(s: str) -> List[str]:
    """Extract likely surname(s) from a text fragment.

    Strategy:
    1) First pass: pick first Cyrillic token that is followed by a name-like token or initials.
    2) Second pass: pick first token that looks like a surname by casing (Titlecase or ALLCAPS, length >=3).
    """
    base = os.path.splitext(os.path.basename(s))[0]
    tokens = tokens_from_name_string(base)
    if not tokens:
        return []

    # Pass 1: token followed by another name token or initials
    for i, tok in enumerate(tokens):
        if not looks_like_name_word(tok):
            continue
        nxt = tokens[i + 1] if i + 1 < len(tokens) else ""
        if (nxt and (looks_like_name_word(nxt) or INIT_TOKEN.match(nxt))):
            # Prefer uppercase tokens or Titlecase as surnames
            if CYRILLIC_CAP_UPPER.match(tok) or (tok[:1].isupper() and tok[1:].islower()):
                return [normalize_surname(tok)]

    # Pass 2: first plausible standalone surname
    for tok in tokens:
        if not looks_like_name_word(tok):
            continue
        if CYRILLIC_CAP_UPPER.match(tok) or (tok[:1].isupper() and tok[1:].islower()):
            # Avoid too-short tokens (e.g., ЕКГ); prefer length >= 3
            if len(tok.replace("-", "")) >= 3:
                return [normalize_surname(tok)]

    return []


def normalize_surname(s: str) -> str:
    # Normalize casing: Titlecase common form, but preserve all-caps if detected
    if CYRILLIC_CAP_UPPER.match(s):
        return s
    return s.capitalize()


def read_b2_surname_from_excel(path: str) -> Optional[str]:
    if openpyxl is None:
        return None
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.active
            if ws is None:
                return None
            val = ws["B2"].value
        finally:
            wb.close()
        if not val or not isinstance(val, str):
            return None
        # Take the first word as surname
        first = val.strip().split()
        if not first:
            return None
        surname = first[0]
        # Validate Cyrillic
        if not CYRILLIC_WORD.match(surname):
            return None
        return normalize_surname(surname)
    except Exception:
        return None


def is_personal_data_excel(filename: str) -> bool:
    name_lower = filename.lower()
    # Accept both correct and common typo extensions
    return (
        (name_lower.endswith(".xlsx") or name_lower.endswith(".xslx"))
        and os.path.splitext(filename)[0].lower() == "персональные данные"
    )


def walk_and_collect(root: str) -> Tuple[List[str], Dict[str, Set[str]]]:
    all_paths: List[str] = []
    mapping: Dict[str, Set[str]] = {}

    for dirpath, _dirnames, filenames in os.walk(root):
        for fname in filenames:
            full = os.path.join(dirpath, fname)
            all_paths.append(full)

            snils = find_snils_from_path(full)
            if not snils:
                continue

            surnames: Set[str] = mapping.setdefault(snils, set())

            # From filename
            for s in extract_surname_candidates_from_text(fname):
                surnames.add(s)

            # From Excel B2 if file matches criteria
            if is_personal_data_excel(fname):
                s_b2 = read_b2_surname_from_excel(full)
                if s_b2:
                    surnames.add(s_b2)

    return all_paths, mapping


def write_all_paths(paths: Iterable[str], out_path: str) -> None:
    with open(out_path, "w", encoding="utf-8") as f:
        for p in paths:
            f.write(p + "\n")


def write_snils_surnames(mapping: Dict[str, Set[str]], out_path: str) -> None:
    # Sort by SNILS then surname (case-insensitive)
    rows: List[Tuple[str, str]] = []
    for snils, surnames in mapping.items():
        for s in surnames:
            rows.append((snils, s))
    rows.sort(key=lambda x: (x[0], x[1].lower()))

    with open(out_path, "w", encoding="utf-8") as f:
        for snils, surname in rows:
            f.write(f"{snils} {surname}\n")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract SNILS + surname pairs from directory tree"
    )
    parser.add_argument(
        "root",
        nargs="?",
        default=os.getcwd(),
        help="Root directory to scan (default: current directory)",
    )
    parser.add_argument(
        "--paths",
        dest="paths_out",
        default=None,
        help="Output file for all paths (default: ROOT/all_paths.txt)",
    )
    parser.add_argument(
        "--out",
        dest="pairs_out",
        default=None,
        help="Output file for SNILS + surname pairs (default: ROOT/snils_surnames.txt)",
    )

    args = parser.parse_args(argv)
    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        print(f"[ERROR] Root is not a directory: {root}", file=sys.stderr)
        return 2

    paths_out = args.paths_out or os.path.join(root, "all_paths.txt")
    pairs_out = args.pairs_out or os.path.join(root, "snils_surnames.txt")

    if openpyxl is None:
        print(
            "[WARN] openpyxl not available. Excel B2 extraction will be skipped.",
            file=sys.stderr,
        )

    all_paths, mapping = walk_and_collect(root)
    write_all_paths(all_paths, paths_out)
    write_snils_surnames(mapping, pairs_out)

    print(f"[OK] Wrote {len(all_paths)} paths to: {paths_out}")
    print(
        f"[OK] Wrote {sum(len(v) for v in mapping.values())} pairs to: {pairs_out}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

