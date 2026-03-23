#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import subprocess
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

import openpyxl

MONTHLY_ENTRY_RE = re.compile(
    r"^\s*(\d{2}/\d{2}/\d{4})\s+(.*?)\s+((?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2})(\s+)(\d{2}/\d{2}/\d{4})\s*$"
)
MONTHLY_OPENING_RE = re.compile(
    r"^\s*(\d{2}/\d{2}/\d{4})\s+SOLDE REPORT\s+((?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2})\s*$"
)
MONTHLY_TOTAL_RE = re.compile(
    r"^\s*TOTAL PERIODE\s+((?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2})\s+((?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2})\s*$"
)
MONTHLY_FINAL_RE = re.compile(r"^\s*SOLDE NET\s+((?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2})\s*$")
MONTHLY_DEBIT_GAP = 20

FOLIO_ENTRY_RE = re.compile(
    r"^\s*(\d{2}/\d{2}/\d{4})\s+(.*?)\s+(\d{2}/\d{2}/\d{4})\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\s*$"
)
FOLIO_OPENING_RE = re.compile(
    r"^\s*(\d{2}/\d{2}/\d{4})\s+SOLDE INITIAL\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\s*$"
)
FOLIO_TOTAL_RE = re.compile(
    r"TOTAL MOUVEMENT\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\s*$"
)
FOLIO_FINAL_RE = re.compile(
    r"SOLDE AU\s+(\d{2}/\d{2}/\d{4})\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\s*$"
)
FOLIO_AMOUNT_COLUMN_SPLIT = 125


@dataclass
class Entry:
    date: str
    valeur: str
    libelle: str
    debit: Decimal
    credit: Decimal


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract SRM statement PDFs into a 5-column Excel file."
    )
    parser.add_argument("input_pdf", type=Path)
    parser.add_argument("-o", "--output", type=Path)
    return parser.parse_args()


def default_output_path(input_pdf: Path) -> Path:
    return input_pdf.with_name(f"{input_pdf.stem}_extracted.xlsx")


def parse_amount_comma(text: str) -> Decimal:
    return Decimal(text.replace(",", ""))


def parse_amount_space(text: str) -> Decimal:
    return Decimal(text.replace(" ", "").replace(",", "."))


def extract_text(input_pdf: Path) -> list[str]:
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", str(input_pdf), "-"],
            check=True,
            capture_output=True,
            text=True,
        )
    except FileNotFoundError as exc:
        raise RuntimeError(
            "Le binaire systeme 'pdftotext' est introuvable. "
            "Ce projet depend de Poppler et ne peut pas extraire le PDF sans cet outil."
        ) from exc
    return result.stdout.splitlines()


def write_workbook(entries: list[Entry], opening_label: str, opening_amount: Decimal, output_xlsx: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Date", "Valeur", "Libellé", "Débit", "Crédit"])
    ws.append([None, None, opening_label, None, float(opening_amount)])

    for entry in entries:
        ws.append(
            [
                entry.date,
                entry.valeur,
                entry.libelle,
                float(entry.debit) if entry.debit else None,
                float(entry.credit) if entry.credit else None,
            ]
        )

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 58
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    wb.save(output_xlsx)


def extract_monthly(lines: list[str]) -> tuple[list[Entry], str, Decimal, Decimal, Decimal, Decimal]:
    opening_date: str | None = None
    opening_amount: Decimal | None = None
    total_debit_stmt: Decimal | None = None
    total_credit_stmt: Decimal | None = None
    final_stmt: Decimal | None = None
    entries: list[Entry] = []

    for line in lines:
        if opening_amount is None:
            match = MONTHLY_OPENING_RE.match(line)
            if match:
                opening_date = match.group(1)
                opening_amount = parse_amount_comma(match.group(2))
                continue

        match = MONTHLY_TOTAL_RE.match(line)
        if match:
            total_debit_stmt = parse_amount_comma(match.group(1))
            total_credit_stmt = parse_amount_comma(match.group(2))
            continue

        match = MONTHLY_FINAL_RE.match(line)
        if match:
            final_stmt = parse_amount_comma(match.group(1))
            continue

        match = MONTHLY_ENTRY_RE.match(line)
        if not match:
            continue

        amount = parse_amount_comma(match.group(3))
        gap = len(match.group(4))
        entries.append(
            Entry(
                date=match.group(1),
                valeur=match.group(5),
                libelle=" ".join(match.group(2).split()),
                debit=amount if gap >= MONTHLY_DEBIT_GAP else Decimal("0"),
                credit=amount if gap < MONTHLY_DEBIT_GAP else Decimal("0"),
            )
        )

    if None in (opening_date, opening_amount, total_debit_stmt, total_credit_stmt, final_stmt):
        raise RuntimeError("Could not parse monthly SRM statement totals")

    total_debit = sum((entry.debit for entry in entries), Decimal("0"))
    total_credit = sum((entry.credit for entry in entries), Decimal("0"))
    credit_with_opening = total_credit + opening_amount
    final_computed = credit_with_opening - total_debit

    if total_debit != total_debit_stmt or credit_with_opening != total_credit_stmt or final_computed != final_stmt:
        raise RuntimeError(
            "Verification failed: "
            f"debit diff={total_debit - total_debit_stmt}, "
            f"credit diff={credit_with_opening - total_credit_stmt}, "
            f"final diff={final_computed - final_stmt}"
        )

    return entries, f"SOLDE REPORT {opening_date}", opening_amount, total_debit, credit_with_opening, final_stmt


def extract_folio(lines: list[str]) -> tuple[list[Entry], str, Decimal, Decimal, Decimal, Decimal]:
    opening_date: str | None = None
    opening_amount: Decimal | None = None
    total_debit_stmt: Decimal | None = None
    total_credit_stmt: Decimal | None = None
    final_stmt: Decimal | None = None
    entries: list[Entry] = []

    for line in lines:
        if opening_amount is None:
            match = FOLIO_OPENING_RE.match(line)
            if match:
                opening_date = match.group(1)
                opening_amount = parse_amount_space(match.group(2))
                continue

        match = FOLIO_TOTAL_RE.search(line)
        if match:
            total_debit_stmt = parse_amount_space(match.group(1))
            total_credit_stmt = parse_amount_space(match.group(2))
            continue

        match = FOLIO_FINAL_RE.search(line)
        if match:
            final_stmt = parse_amount_space(match.group(2))
            continue

        match = FOLIO_ENTRY_RE.match(line)
        if not match:
            continue

        amount_match = re.search(r"(?:\d{1,3}(?: \d{3})*|\d+),\d{2}", line)
        if amount_match is None:
            continue

        amount = parse_amount_space(match.group(4))
        entries.append(
            Entry(
                date=match.group(1),
                valeur=match.group(3),
                libelle=" ".join(match.group(2).split()),
                debit=amount if amount_match.start() < FOLIO_AMOUNT_COLUMN_SPLIT else Decimal("0"),
                credit=amount if amount_match.start() >= FOLIO_AMOUNT_COLUMN_SPLIT else Decimal("0"),
            )
        )

    if None in (opening_date, opening_amount, total_debit_stmt, total_credit_stmt, final_stmt):
        raise RuntimeError("Could not parse folio SRM statement totals")

    total_debit = sum((entry.debit for entry in entries), Decimal("0"))
    total_credit = sum((entry.credit for entry in entries), Decimal("0"))
    credit_with_opening = total_credit + opening_amount
    final_computed = credit_with_opening - total_debit

    if total_debit != total_debit_stmt or credit_with_opening != total_credit_stmt or final_computed != final_stmt:
        raise RuntimeError(
            "Verification failed: "
            f"debit diff={total_debit - total_debit_stmt}, "
            f"credit diff={credit_with_opening - total_credit_stmt}, "
            f"final diff={final_computed - final_stmt}"
        )

    return entries, f"SOLDE INITIAL {opening_date}", opening_amount, total_debit, credit_with_opening, final_stmt


def main() -> None:
    args = parse_args()
    input_pdf = args.input_pdf
    output_xlsx = args.output or default_output_path(input_pdf)
    lines = extract_text(input_pdf)
    full_text = "\n".join(lines)

    if "RELEVE MENSUEL DE COMPTE" in full_text:
        variant = "monthly"
        entries, opening_label, opening_amount, total_debit, total_credit_stmt, final_stmt = extract_monthly(lines)
    else:
        variant = "folio"
        entries, opening_label, opening_amount, total_debit, total_credit_stmt, final_stmt = extract_folio(lines)

    write_workbook(entries, opening_label, opening_amount, output_xlsx)

    print(f"Input: {input_pdf}")
    print(f"Detected SRM variant: {variant}")
    print(f"Rows extracted (transactions): {len(entries)}")
    print(f"Opening amount: {opening_amount:,.2f}")
    print(f"Verified debit total: {total_debit:,.2f}")
    print(f"Verified credit total (including opening): {total_credit_stmt:,.2f}")
    print(f"Verified final balance: {final_stmt:,.2f}")
    print(f"Wrote: {output_xlsx}")


if __name__ == "__main__":
    main()
