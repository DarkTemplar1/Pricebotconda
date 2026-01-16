#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AUTOMAT – liczenie statystyk WYŁĄCZNIE z Excela

USTALENIA:
- plik: Polska.xlsx
- arkusz BAZY:  "Polska (data scalenia plików csv)"
- arkusz RAPORTU: "raport"

ZASADY:
- ZERO internetu, ZERO requests
- STRICT: liczy TYLKO przy 100% adresu
- brak dopasowania w bazie = komunikat
- brak fallbacków po mieście
- wyniki zaokrąglone do 2 miejsc
"""

from __future__ import annotations

import pandas as pd
import numpy as np
from pathlib import Path
import argparse
from math import isnan
from typing import Optional


# =========================
# KONFIG
# =========================

EXCEL_FILE = "Polska.xlsx"

BASE_SHEET = "Polska (data scalenia plików csv)"
RAPORT_SHEET = "raport"

# kolumny adresu
COL_WOJ = "Województwo"
COL_POW = "Powiat"
COL_GMI = "Gmina"
COL_MIA = "Miejscowość"
COL_DZL = "Dzielnica"   # opcjonalna

# dane
COL_OBSZAR = "Obszar"

# kolumny wynikowe (RAPORT)
COL_CENA_M2 = "Średnia cena za m2 ( z bazy)"
COL_CENA_M2_KOR = "Średnia skorygowana cena za m2"
COL_WARTOSC = "Statystyczna wartość nieruchomości"

STRICT_MSG = "BRAK LUB NIEPEŁNY ADRES – WPISZ ADRES MANUALNIE"


# =========================
# POMOCNICZE
# =========================

def _is_missing(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and isnan(v):
        return True
    if isinstance(v, str) and v.strip() in ("", "---"):
        return True
    return False


def _round2(v: Optional[float]) -> Optional[float]:
    try:
        return round(float(v), 2)
    except Exception:
        return None


def _has_full_address(row: pd.Series) -> bool:
    for c in (COL_WOJ, COL_POW, COL_GMI, COL_MIA):
        if _is_missing(row.get(c)):
            return False
    return True


# =========================
# IO
# =========================

def read_sheet(path: Path, sheet: str) -> pd.DataFrame:
    xl = pd.ExcelFile(path, engine="openpyxl")
    if sheet not in xl.sheet_names:
        raise RuntimeError(f"Brak arkusza '{sheet}' w pliku {path.name}")
    return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")


def write_raport(path: Path, df: pd.DataFrame) -> None:
    with pd.ExcelWriter(
        path,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name=RAPORT_SHEET, index=False)


# =========================
# LOGIKA
# =========================

def process(base: pd.DataFrame, raport: pd.DataFrame) -> pd.DataFrame:

    # upewnij się że kolumny wynikowe istnieją
    for c in (COL_CENA_M2, COL_CENA_M2_KOR, COL_WARTOSC):
        if c not in raport.columns:
            raport[c] = None

    for idx, row in raport.iterrows():

        # STRICT – brak pełnego adresu
        if not _has_full_address(row):
            raport.loc[idx, [COL_CENA_M2, COL_CENA_M2_KOR, COL_WARTOSC]] = STRICT_MSG
            continue

        # powierzchnia
        if _is_missing(row.get(COL_OBSZAR)):
            raport.loc[idx, [COL_CENA_M2, COL_CENA_M2_KOR, COL_WARTOSC]] = STRICT_MSG
            continue

        try:
            obszar = float(row[COL_OBSZAR])
        except Exception:
            raport.loc[idx, [COL_CENA_M2, COL_CENA_M2_KOR, COL_WARTOSC]] = STRICT_MSG
            continue

        # ===== FILTR BAZY (100% adresu) =====
        mask = (
            (base[COL_WOJ] == row[COL_WOJ]) &
            (base[COL_POW] == row[COL_POW]) &
            (base[COL_GMI] == row[COL_GMI]) &
            (base[COL_MIA] == row[COL_MIA])
        )

        # dzielnica – tylko jeśli wpisana w raporcie
        if COL_DZL in raport.columns and not _is_missing(row.get(COL_DZL)):
            mask &= (base[COL_DZL] == row[COL_DZL])

        subset = base.loc[mask]

        if subset.empty:
            raport.loc[idx, [COL_CENA_M2, COL_CENA_M2_KOR, COL_WARTOSC]] = STRICT_MSG
            continue

        # ===== LICZENIE CENY ZA M2 =====
        if COL_CENA_M2 not in subset.columns:
            raport.loc[idx, [COL_CENA_M2, COL_CENA_M2_KOR, COL_WARTOSC]] = STRICT_MSG
            continue

        ceny = pd.to_numeric(subset[COL_CENA_M2], errors="coerce").dropna()
        if ceny.empty:
            raport.loc[idx, [COL_CENA_M2, COL_CENA_M2_KOR, COL_WARTOSC]] = STRICT_MSG
            continue

        cena_m2 = ceny.mean()
        cena_kor = cena_m2   # miejsce na Twoją korektę w przyszłości
        wartosc = cena_kor * obszar

        raport.at[idx, COL_CENA_M2] = _round2(cena_m2)
        raport.at[idx, COL_CENA_M2_KOR] = _round2(cena_kor)
        raport.at[idx, COL_WARTOSC] = _round2(wartosc)

    return raport


# =========================
# MAIN
# =========================

def main():
    ap = argparse.ArgumentParser(description="Automat – STRICT, Excel only")
    ap.add_argument(
        "--excel",
        default=EXCEL_FILE,
        help="Ścieżka do Polska.xlsx (domyślnie: Polska.xlsx)"
    )
    args = ap.parse_args()

    path = Path(args.excel).resolve()
    if not path.exists():
        raise FileNotFoundError(path)

    base = read_sheet(path, BASE_SHEET)
    raport = read_sheet(path, RAPORT_SHEET)

    raport = process(base, raport)
    write_raport(path, raport)

    print("✔ AUTOMAT ZAKOŃCZONY (STRICT, Excel only)")


if __name__ == "__main__":
    main()
