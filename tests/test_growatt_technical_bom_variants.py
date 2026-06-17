from __future__ import annotations

import csv
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))


def load_module(module_name: str, file_name: str):
    spec = importlib.util.spec_from_file_location(module_name, SCRIPTS_DIR / file_name)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load {file_name} for tests")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


builder = load_module(
    "build_growatt_technical_bom_variants_test",
    "build-growatt-technical-bom-variants.py",
)
baseline = load_module(
    "growatt_rvc_baseline_test",
    "growatt-rvc-baseline.py",
)


class GrowattTechnicalBomVariantTests(unittest.TestCase):
    def test_build_variant_rows_maps_per_code_csv_to_case_schema(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            per_code_dir = Path(tmpdir)
            (per_code_dir / "PV01.0117300.csv").write_text(
                "\n".join(
                    [
                        "rootProductCode,leafComponentCode,leafComponentDescription,brand,unit,quantity,samplePath,pathCount,sourceFileCount,sourceFiles",
                        "PV01.0117300,001.0000100,Resistor A,BRAND-A,ST,1.5,PV01 -> A,1,1,file-a.xlsx",
                        "PV01.0117300,001.0000100,Resistor A,BRAND-A,ST,0.5,PV01 -> B,2,2,file-a.xlsx;file-b.xlsx",
                        "PV01.0117300,005.0001100,Capacitor B,BRAND-B,ST,3.0,PV01 -> C,1,1,file-c.xlsx",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )

            rows_out, code_summaries = builder.build_variant_rows(per_code_dir)

        self.assertEqual(len(rows_out), 3)
        self.assertEqual(rows_out[0]["bom_code"], "PV01.0117300")
        self.assertEqual(
            rows_out[0]["bom_variant_id"],
            "PV01.0117300__technical_flatten_20260423",
        )
        self.assertEqual(rows_out[0]["dm_row_no"], 1)
        self.assertEqual(rows_out[1]["ordinal_key"], "T2PV01.0117300")
        self.assertEqual(rows_out[2]["material_code"], "005.0001100")
        self.assertEqual(rows_out[2]["qty_per_unit"], 3.0)
        self.assertEqual(code_summaries[0]["row_count"], 3)
        self.assertEqual(code_summaries[0]["unique_material_count"], 2)
        self.assertEqual(code_summaries[0]["duplicate_leaf_material_count"], 1)

    def test_baseline_loader_accepts_generated_technical_variant_csv(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / "technical-bom-variants.csv"
            with csv_path.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.DictWriter(
                    handle,
                    fieldnames=[
                        "bom_source_id",
                        "bom_source_kind",
                        "bom_code",
                        "product_family_code",
                        "bom_variant_id",
                        "dm_row_no",
                        "ordinal_key",
                        "material_code",
                        "qty_per_unit",
                    ],
                )
                writer.writeheader()
                writer.writerows(
                    [
                        {
                            "bom_source_id": "technical-priority-2026-04-23",
                            "bom_source_kind": "technical_flatten",
                            "bom_code": "PV01.0117300",
                            "product_family_code": "PV01.0117300",
                            "bom_variant_id": "PV01.0117300__technical_flatten_20260423",
                            "dm_row_no": 1,
                            "ordinal_key": "T1PV01.0117300",
                            "material_code": "001.0000100",
                            "qty_per_unit": 1.5,
                        },
                        {
                            "bom_source_id": "technical-priority-2026-04-23",
                            "bom_source_kind": "technical_flatten",
                            "bom_code": "PV01.0117300",
                            "product_family_code": "PV01.0117300",
                            "bom_variant_id": "PV01.0117300__technical_flatten_20260423",
                            "dm_row_no": 2,
                            "ordinal_key": "T2PV01.0117300",
                            "material_code": "005.0001100",
                            "qty_per_unit": 3.0,
                        },
                    ]
                )

            variants, variants_by_model = baseline.load_bom_variants(csv_path)

        self.assertEqual(len(variants), 1)
        variant = variants[0]
        self.assertEqual(variant.variant_id, "PV01.0117300__technical_flatten_20260423")
        self.assertEqual(variant.model_code, "PV01.0117300")
        self.assertEqual(variant.start_row, 1)
        self.assertEqual(variant.end_row, 2)
        self.assertEqual(variant.line_count, 2)
        self.assertEqual(
            [line.material_code for line in variants_by_model["PV01.0117300"][0].lines],
            ["001.0000100", "005.0001100"],
        )


if __name__ == "__main__":
    unittest.main()
