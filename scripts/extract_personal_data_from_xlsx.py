#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Извлекает СНИЛС (B1) и фамилию (первое слово из ФИО в B2)
из всех файлов с именем "Персональные данные.xlsx" в указанной
директории и её вложенных папках. Результат пишет в txt:

00000000000 Фамилия

Используется только содержимое Excel, структура папок и имена файлов,
кроме точного совпадения названия книги, не анализируются.

Запуск:
  python scripts/extract_personal_data_from_xlsx.py [ROOT_DIR] [--out OUT]

По умолчанию:
- ROOT_DIR = текущая директория
- OUT = scripts/snils_surnames.txt
"""

from __future__ import annotations

import os
import re
import sys
import argparse
from typing import Iterable, List, Optional, Tuple

try:
    import openpyxl  # type: ignore
except Exception:  # pragma: no cover
    openpyxl = None


EXPECTED_XLSX_NAME = "персональные данные.xlsx"
SNILS_DIGITS_RE = re.compile(r"\D+")


def is_target_excel(filename: str) -> bool:
    """True, если это именно "Персональные данные.xlsx" (без учёта регистра)."""
    return filename.lower() == EXPECTED_XLSX_NAME


def read_b1_b2(path: str) -> Tuple[Optional[str], Optional[str]]:
    """Читает B1 (СНИЛС) и B2 (ФИО) из книги."""
    if openpyxl is None:
        return None, None
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        b1 = ws["B1"].value
        b2 = ws["B2"].value
        wb.close()
    except Exception:
        return None, None

    snils = None
    if b1 is not None:
        snils_raw = str(b1)
        snils_digits = SNILS_DIGITS_RE.sub("", snils_raw)
        if len(snils_digits) == 11 and snils_digits.isdigit():
            snils = snils_digits

    fio = str(b2).strip() if b2 is not None else None
    return snils, fio


def surname_from_fio(fio: Optional[str]) -> Optional[str]:
    if not fio:
        return None
    parts = fio.strip().split()
    if not parts:
        return None
    return parts[0]


def find_all_personal_data_xlsx(root: str) -> Iterable[str]:
    for dirpath, _, filenames in os.walk(root):
        for fname in filenames:
            if is_target_excel(fname):
                yield os.path.join(dirpath, fname)


def write_pairs(rows: List[Tuple[str, str]], out_path: str) -> None:
    with open(out_path, "w", encoding="utf-8") as f:
        for snils, surname in rows:
            f.write(f"{snils} {surname}\n")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Собирает СНИЛС (B1) и фамилию (из B2) только из файлов 'Персональные данные.xlsx'"
        )
    )
    parser.add_argument(
        "root",
        nargs="?",
        default=os.getcwd(),
        help="Корневая папка для сканирования (по умолчанию: текущая)",
    )
    parser.add_argument(
        "--out",
        dest="out_path",
        default=None,
        help="Путь к выходному txt (по умолчанию: scripts/snils_surnames.txt)",
    )

    args = parser.parse_args(argv)
    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        print(f"[ERROR] Не папка: {root}", file=sys.stderr)
        return 2

    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = args.out_path or os.path.join(script_dir, "snils_surnames.txt")

    if openpyxl is None:
        print("[ERROR] Модуль openpyxl не установлен", file=sys.stderr)
        return 3

    rows: List[Tuple[str, str]] = []
    files_found = 0
    for xlsx_path in find_all_personal_data_xlsx(root):
        files_found += 1
        snils, fio = read_b1_b2(xlsx_path)
        surname = surname_from_fio(fio)
        if snils and surname:
            rows.append((snils, surname))

    # Не требовалась дедупликация, пишем как есть
    write_pairs(rows, out_path)

    print(f"[OK] Найдено файлов: {files_found}")
    print(f"[OK] Извлечено записей: {len(rows)}")
    print(f"[OK] Результат: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

