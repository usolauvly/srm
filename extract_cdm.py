#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import subprocess
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

import openpyxl

ENTRY_RE = re.compile(
    r"^\s*(\d{2}\s+\d{2}\s+\d{2})\s+(\d{2}\s+\d{2}\s+\d{2})\s+(.*?)\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\s*$"
)
TOTAL_RE = re.compile(
    r"TOTAL MOUVEMENTS\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\s*$"
)
ANCIEN_RE = re.compile(
    r"ANCIEN SOLDE AU\s+(\d{2}/\d{2}/\d{4})\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})"
)
REPORTE_RE = re.compile(r"SOLDE REPORTE\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})")
FINAL_RE = re.compile(
    r"NOUVEAU SOLDE AU\s+(\d{2}/\d{2}/\d{4})\s+((?:\d{1,3}(?: \d{3})*|\d+),\d{2})"
)
AMOUNT_RE = re.compile(r"(?:\d{1,3}(?: \d{3})*|\d+),\d{2}")
AMOUNT_COLUMN_SPLIT = 160


@dataclass
class Entry:
    date: str
    valeur: str
    libelle: str
    debit: Decimal
    credit: Decimal


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract Crédit du Maroc statements into a 5-column Excel file."
    )
    parser.add_argument("input_pdf", type=Path)
    parser.add_argument("-o", "--output", type=Path)
    return parser.parse_args()


def default_output_path(input_pdf: Path) -> Path:
    return input_pdf.with_name(f"{input_pdf.stem}_extracted.xlsx")


def parse_amount(text: str) -> Decimal:
    return Decimal(text.replace(" ", "").replace(",", "."))


def format_short_date(text: str) -> str:
    day, month, year = text.split()
    return f"{day}/{month}/20{year}"


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


def write_workbook(entries: list[Entry], opening_label: str, opening_signed: Decimal, output_xlsx: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Date", "Valeur", "Libellé", "Débit", "Crédit"])
    ws.append(
        [
            None,
            None,
            opening_label,
            float(-opening_signed) if opening_signed < 0 else None,
            float(opening_signed) if opening_signed > 0 else None,
        ]
    )

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


def main() -> None:
    args = parse_args()
    input_pdf = args.input_pdf
    output_xlsx = args.output or default_output_path(input_pdf)
    lines = extract_text(input_pdf)

    entries: list[Entry] = []
    opening_label: str | None = None
    opening_signed: Decimal | None = None
    total_stmt: tuple[Decimal, Decimal] | None = None
    final_signed: Decimal | None = None

    for line in lines:
        if opening_signed is None:
            match = ANCIEN_RE.search(line)
            if match:
                amount_match = AMOUNT_RE.search(line)
                if amount_match is None:
                    continue
                amount = parse_amount(match.group(2))
                opening_signed = amount if amount_match.start() >= AMOUNT_COLUMN_SPLIT else -amount
                opening_label = f"ANCIEN SOLDE AU {match.group(1)}"
                continue

        match = TOTAL_RE.search(line)
        if match:
            total_stmt = (parse_amount(match.group(1)), parse_amount(match.group(2)))
            continue

        match = FINAL_RE.search(line)
        if match:
            amount_match = AMOUNT_RE.search(line)
            if amount_match is None:
                continue
            amount = parse_amount(match.group(2))
            final_signed = amount if amount_match.start() >= AMOUNT_COLUMN_SPLIT else -amount
            continue

        if opening_signed is None:
            match = REPORTE_RE.search(line)
            if match:
                amount_match = AMOUNT_RE.search(line)
                if amount_match is None:
                    continue
                amount = parse_amount(match.group(1))
                opening_signed = amount if amount_match.start() >= AMOUNT_COLUMN_SPLIT else -amount
                opening_label = "SOLDE REPORTE"
                continue

        match = ENTRY_RE.match(line)
        if not match:
            continue

        amount_match = AMOUNT_RE.search(line)
        if amount_match is None:
            continue

        amount = parse_amount(match.group(4))
        entries.append(
            Entry(
                date=format_short_date(" ".join(match.group(1).split())),
                valeur=format_short_date(" ".join(match.group(2).split())),
                libelle=" ".join(match.group(3).split()),
                debit=amount if amount_match.start() < AMOUNT_COLUMN_SPLIT else Decimal("0"),
                credit=amount if amount_match.start() >= AMOUNT_COLUMN_SPLIT else Decimal("0"),
            )
        )

    if opening_signed is None or opening_label is None:
        raise RuntimeError("Opening balance not found")
    if total_stmt is None:
        raise RuntimeError("TOTAL MOUVEMENTS line not found")
    if final_signed is None:
        raise RuntimeError("NOUVEAU SOLDE AU line not found")

    total_debit = sum((entry.debit for entry in entries), Decimal("0"))
    total_credit = sum((entry.credit for entry in entries), Decimal("0"))
    computed_final = opening_signed - total_debit + total_credit

    diff_debit = total_debit - total_stmt[0]
    diff_credit = total_credit - total_stmt[1]
    diff_final = computed_final - final_signed

    if diff_debit or diff_credit or diff_final:
        raise RuntimeError(
            "Verification failed: "
            f"debit diff={diff_debit}, credit diff={diff_credit}, final diff={diff_final}"
        )

    write_workbook(entries, opening_label, opening_signed, output_xlsx)

    print(f"Input: {input_pdf}")
    print(f"Rows extracted (transactions): {len(entries)}")
    print(f"Opening balance: {abs(opening_signed):,.2f}")
    print(f"Transactions total debit: {total_debit:,.2f}")
    print(f"Transactions total credit: {total_credit:,.2f}")
    print(f"Statement TOTAL MOUVEMENTS debit: {total_stmt[0]:,.2f}")
    print(f"Statement TOTAL MOUVEMENTS credit: {total_stmt[1]:,.2f}")
    print(f"Verified final balance: {abs(final_signed):,.2f}")
    print(f"Wrote: {output_xlsx}")


if __name__ == "__main__":
    main()
