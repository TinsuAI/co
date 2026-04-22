#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import urllib.parse
import urllib.request
from pathlib import Path

import openpyxl


BASE_URL = "https://www.customs.gov.vn/bridge?url=/customs/api/"
DEFAULT_OUTPUT = Path("data/reference/DS_ty_gia_ngoai_te.xlsx")
USER_AGENT = "Mozilla/5.0"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--language", default="TIENG_VIET")
    return parser.parse_args()


def fetch_json(endpoint: str) -> dict[str, object]:
    url = f"{BASE_URL}{endpoint}"
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


def main() -> None:
    args = parse_args()
    currency_payload = fetch_json(f"GetListDongTienTyGia&language={urllib.parse.quote(args.language)}")
    usd_payload = fetch_json(f"GetListUSDRate&language={urllib.parse.quote(args.language)}")
    other_payload = fetch_json(f"GetListOtherRate&language={urllib.parse.quote(args.language)}")

    currency_names = {
        item["DONG_TIEN"]: item["TEN_DONG_TIEN"]
        for item in currency_payload.get("d", [])
        if item.get("DONG_TIEN")
    }

    rows: list[tuple[str, str, str, str]] = []
    for item in usd_payload.get("d", []):
        rows.append(
            (
                str(item.get("LOAI_NGOAI_TE") or "USD"),
                currency_names.get("USD", "Đô-la Mỹ"),
                str(item.get("HIEU_LUC_TU_NGAY") or ""),
                str(item.get("TY_GIA") or ""),
            )
        )
    for item in other_payload.get("d", []):
        code = str(item.get("LOAI_NGOAI_TE") or "")
        rows.append(
            (
                code,
                str(item.get("TEN_NGOAI_TE") or currency_names.get(code, "")),
                str(item.get("HIEU_LUC_TU_NGAY") or ""),
                str(item.get("TY_GIA") or ""),
            )
        )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = "Danh sách tỷ giá ngoại tệ"
    sheet.append([None, "Nguyên tệ", "Ngày hiệu lực", "Tỷ giá"])
    for code, currency_name, effective_date, rate_text in rows:
        sheet.append([code, currency_name, effective_date, rate_text])
    workbook.save(args.output)

    print(f"Wrote {args.output}")
    print(f"Currencies: {len(currency_names)}")
    print(f"Rows: {len(rows)}")


if __name__ == "__main__":
    main()
