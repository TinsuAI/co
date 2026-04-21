#!/usr/bin/env python3

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from zipfile import ZipFile
import argparse
import xml.etree.ElementTree as ET

import openpyxl


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

DEFAULT_WORKBOOK = Path(
    "data/extracted/Growatt-20260421/Growatt/"
    "tru lui CO final  SXXK - 2025 commercial-MAC - Huyền đúng.xlsm"
)

SHIPMENTS = {
    "GIN01426B282": {
        "PV01.0117300": 130.0,
        "PV01.0117400": 87.0,
        "PV02.0228801": 144.0,
        "PV02.0228901": 87.0,
        "PV02.0229000": 29.0,
        "PV02.0229100": 173.0,
    },
    "GIN01426C171": {
        "PV00.0048400": 309.0,
        "PV00.0048500": 659.0,
        "PV01.0117600": 346.0,
    },
}

SEEDED_CODES = {
    "020.0023000",
    "030.0095100",
    "030.0054901",
    "020.0004802",
    "018.0531701",
    "020.0028800",
    "006.0033900",
    "007.0033200",
    "007.0068500",
    "030.0122700",
}


def family(name: str | None) -> str:
    if not name:
        return ""
    return str(name).split("#&", 1)[0].strip()


def load_shared_strings(archive: ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall(f"{NS}si"):
        parts = [text.text or "" for text in si.iter(f"{NS}t")]
        strings.append("".join(parts))
    return strings


def parse_dm_bom(workbook_path: Path, target_models: set[str]) -> dict[str, list[tuple[str, float]]]:
    bom: dict[str, list[tuple[str, float]]] = defaultdict(list)
    with ZipFile(workbook_path) as archive:
        shared_strings = load_shared_strings(archive)
        for _, elem in ET.iterparse(archive.open("xl/worksheets/sheet4.xml"), events=("end",)):
            if elem.tag != f"{NS}row":
                continue
            row_num = int(elem.attrib["r"])
            if row_num < 7:
                elem.clear()
                continue
            values: dict[str, str] = {}
            for cell in elem.findall(f"{NS}c"):
                ref = cell.attrib["r"]
                col = "".join(ch for ch in ref if ch.isalpha())
                raw = cell.find(f"{NS}v")
                if raw is None:
                    continue
                value = raw.text or ""
                if cell.attrib.get("t") == "s":
                    value = shared_strings[int(value)]
                values[col] = value
            model = values.get("A")
            code = values.get("F")
            qty = values.get("I")
            if model in target_models and code and qty is not None:
                bom[model].append((code, float(qty)))
            elem.clear()
    return bom


def load_nk2_stock(workbook_path: Path):
    workbook = openpyxl.load_workbook(
        workbook_path,
        data_only=True,
        read_only=True,
        keep_vba=True,
    )
    sheet = workbook["NK2"]
    stock_by_code: dict[str, list[dict[str, object]]] = defaultdict(list)
    meta: dict[str, dict[str, str]] = {}
    for row in sheet.iter_rows(min_row=5, values_only=True):
        code = row[4]
        if code is None:
            continue
        code = str(code).strip()
        hs = str(row[5]).strip() if row[5] is not None else ""
        name = row[6]
        origin = row[7]
        unit = float(row[8] or 0)
        stock = float(row[17] or 0)
        if stock <= 0:
            continue
        stock_by_code[code].append(
            {
                "hs": hs,
                "name": name,
                "origin": origin,
                "unit": unit,
                "stock": stock,
            }
        )
        meta.setdefault(code, {"hs": hs, "name": str(name or ""), "family": family(name)})
    return stock_by_code, meta


def effective_unit_cost(rows: list[dict[str, object]], need: float) -> float | None:
    remaining = need
    total_cost = 0.0
    for row in sorted(rows, key=lambda item: float(item["unit"])):
        available = float(row["stock"])
        take = min(remaining, available)
        if take <= 0:
            continue
        total_cost += take * float(row["unit"])
        remaining -= take
        if remaining <= 1e-9:
            return total_cost / need
    return None


def total_stock(rows: list[dict[str, object]]) -> float:
    return sum(float(row["stock"]) for row in rows)


def best_origin(rows: list[dict[str, object]]) -> str:
    return str(sorted(rows, key=lambda item: float(item["unit"]))[0]["origin"] or "")


def analyze(workbook_path: Path, top_n: int) -> None:
    target_models = {model for shipment in SHIPMENTS.values() for model in shipment}
    bom = parse_dm_bom(workbook_path, target_models)
    stock_by_code, meta = load_nk2_stock(workbook_path)

    codes_by_hs: dict[str, set[str]] = defaultdict(set)
    for code, details in meta.items():
        codes_by_hs[details["hs"]].add(code)

    for shipment_id, models in SHIPMENTS.items():
        demand_by_code: dict[str, float] = defaultdict(float)
        for model, quantity in models.items():
            for code, per_unit in bom.get(model, []):
                demand_by_code[code] += quantity * per_unit

        print(f"\nShipment {shipment_id}")
        print(f"Models: {', '.join(f'{model} x {int(qty)}' for model, qty in models.items())}")
        print(f"Distinct BOM codes: {len(demand_by_code)}")

        seeded_in_shipment = sorted(code for code in demand_by_code if code in SEEDED_CODES)
        if seeded_in_shipment:
            print("\nSeeded replacement codes present in shipment:")
            for code in seeded_in_shipment:
                hs = meta.get(code, {}).get("hs", "")
                need = demand_by_code[code]
                current_rows = stock_by_code.get(code, [])
                current_cost = effective_unit_cost(current_rows, need) if current_rows else None
                current_cost_text = f"{current_cost:.6f}" if current_cost is not None else "n/a"
                print(
                    f"- {code} | need={need:.4f} | hs={hs} | "
                    f"current_ec={current_cost_text} | "
                    f"stock={total_stock(current_rows):.4f}"
                )
                candidates = []
                for candidate in codes_by_hs.get(hs, set()):
                    if candidate == code or candidate not in stock_by_code:
                        continue
                    candidate_cost = effective_unit_cost(stock_by_code[candidate], need)
                    if candidate_cost is None:
                        continue
                    if current_cost is not None and candidate_cost >= current_cost:
                        continue
                    candidates.append(
                        (
                            candidate_cost,
                            candidate,
                            total_stock(stock_by_code[candidate]),
                            best_origin(stock_by_code[candidate]),
                            meta[candidate]["name"],
                        )
                    )
                for candidate_cost, candidate, stock, origin, name in sorted(candidates)[:5]:
                    print(
                        f"  -> {candidate} | ec={candidate_cost:.6f} | "
                        f"stock={stock:.4f} | origin={origin} | {name}"
                    )

        broad_opportunities = []
        family_matched_opportunities = []
        for code, need in demand_by_code.items():
            if code not in stock_by_code or code not in meta:
                continue
            current_cost = effective_unit_cost(stock_by_code[code], need)
            if current_cost is None:
                continue
            hs = meta[code]["hs"]
            current_family = meta[code]["family"]
            best_broad = None
            best_family = None
            for candidate in codes_by_hs.get(hs, set()):
                if candidate == code or candidate not in stock_by_code:
                    continue
                candidate_cost = effective_unit_cost(stock_by_code[candidate], need)
                if candidate_cost is None or candidate_cost >= current_cost:
                    continue
                saving = (current_cost - candidate_cost) * need
                item = (
                    saving,
                    code,
                    candidate,
                    need,
                    hs,
                    current_family,
                    current_cost,
                    candidate_cost,
                    total_stock(stock_by_code[candidate]),
                    best_origin(stock_by_code[candidate]),
                )
                if best_broad is None or item[0] > best_broad[0]:
                    best_broad = item
                if meta[candidate]["family"] == current_family:
                    if best_family is None or item[0] > best_family[0]:
                        best_family = item
            if best_broad:
                broad_opportunities.append(best_broad)
            if best_family:
                family_matched_opportunities.append(best_family)

        print("\nTop same-HS opportunities:")
        for item in sorted(broad_opportunities, reverse=True)[:top_n]:
            saving, code, candidate, need, hs, current_family, current_cost, candidate_cost, stock, origin = item
            print(
                f"- {code} -> {candidate} | hs={hs} | family={current_family} | "
                f"need={need:.4f} | current_ec={current_cost:.6f} | cand_ec={candidate_cost:.6f} | "
                f"saving={saving:.2f} | cand_stock={stock:.4f} | cand_origin={origin}"
            )

        print("\nTop same-HS + same-family opportunities:")
        for item in sorted(family_matched_opportunities, reverse=True)[:top_n]:
            saving, code, candidate, need, hs, current_family, current_cost, candidate_cost, stock, origin = item
            print(
                f"- {code} -> {candidate} | hs={hs} | family={current_family} | "
                f"need={need:.4f} | current_ec={current_cost:.6f} | cand_ec={candidate_cost:.6f} | "
                f"saving={saving:.2f} | cand_stock={stock:.4f} | cand_origin={origin}"
            )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--top", type=int, default=15)
    args = parser.parse_args()
    analyze(args.workbook, args.top)


if __name__ == "__main__":
    main()
