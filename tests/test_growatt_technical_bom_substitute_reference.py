from __future__ import annotations

import importlib.util
import sys
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
    "build_growatt_technical_bom_substitute_reference_test",
    "build-growatt-technical-bom-substitute-reference.py",
)


class GrowattTechnicalBomSubstituteReferenceTests(unittest.TestCase):
    def test_collect_code_reference_rows_keeps_only_real_substitute_signals(self) -> None:
        payload = {
            "canonicalSourceChain": [
                {
                    "parentCode": "B700.TEST",
                    "normalizedFileCode": "B700.TEST",
                    "kind": "btp",
                    "filePath": "/tmp/B700.TEST.xlsx",
                    "rows": [
                        {
                            "组件物料": "019.0009300",
                            "组件物料描述": "Only same-level marker",
                            "组件物料品牌": "",
                            "单位": "ST",
                            "标准用量": "1.000000",
                            "替代项目组": "",
                            "可替代物料": "",
                            "优先级": "0",
                            "策略": "",
                            "使用概率": "0",
                            "位置号": "CN1",
                            "同层同组替代标识": "CN1",
                            "替代组合标识": "",
                            "临时替代预留行号": "",
                        },
                        {
                            "组件物料": "012.0004100",
                            "组件物料描述": "Primary part",
                            "组件物料品牌": "BRAND-A",
                            "单位": "ST",
                            "标准用量": "1.000000",
                            "替代项目组": "A1",
                            "可替代物料": "012.0001100,012.0004201",
                            "优先级": "1",
                            "策略": "1",
                            "使用概率": "100",
                            "位置号": "NTC1",
                            "同层同组替代标识": "B700.TEST/A1/00/00",
                            "替代组合标识": "A1-01-00-00",
                            "临时替代预留行号": "",
                        },
                        {
                            "组件物料": "006.0005600",
                            "组件物料描述": "Combo-only part",
                            "组件物料品牌": "BRAND-B",
                            "单位": "ST",
                            "标准用量": "2.000000",
                            "替代项目组": "",
                            "可替代物料": "",
                            "优先级": "0",
                            "策略": "",
                            "使用概率": "0",
                            "位置号": "C99",
                            "同层同组替代标识": "",
                            "替代组合标识": "A9-02-00-00",
                            "临时替代预留行号": "",
                        },
                    ],
                }
            ]
        }

        row_signals, explicit_links, summary = builder.collect_code_reference_rows("PV02.TEST", payload)

        self.assertEqual(len(row_signals), 2)
        self.assertEqual(len(explicit_links), 2)
        self.assertEqual(summary["substitute_signal_row_count"], 2)
        self.assertEqual(summary["explicit_substitute_link_count"], 2)
        self.assertEqual(summary["reference_group_count"], 2)

        first = row_signals[0]
        self.assertEqual(first["component_code"], "012.0004100")
        self.assertEqual(first["signal_class"], "group_with_explicit_materials")
        self.assertEqual(first["substitute_material_codes"], "012.0001100;012.0004201")
        self.assertEqual(first["reference_group_key"], "PV02.TEST|B700.TEST|group|A1")

        second = row_signals[1]
        self.assertEqual(second["component_code"], "006.0005600")
        self.assertEqual(second["signal_class"], "combo_marker_only")
        self.assertEqual(second["reference_group_key"], "PV02.TEST|B700.TEST|combo|A9-02-00-00")

    def test_build_reference_groups_aggregates_member_and_explicit_codes(self) -> None:
        row_signals = [
            {
                "reference_group_key": "PV02.TEST|B700.TEST|group|A1",
                "root_product_code": "PV02.TEST",
                "source_parent_code": "B700.TEST",
                "source_file_code": "B700.TEST",
                "source_file_name": "B700.TEST.xlsx",
                "source_kind": "btp",
                "substitute_group": "A1",
                "component_code": "012.0004100",
                "substitute_materials_raw": "012.0001100;012.0004201",
                "substitute_priority": "1",
                "substitute_strategy": "1",
                "usage_probability": "100",
                "same_level_group_marker": "B700.TEST/A1/00/00",
                "substitute_combo_marker": "A1-01-00-00",
                "temp_substitute_row_no": "",
                "signal_class": "group_with_explicit_materials",
                "position_no": "NTC1",
            },
            {
                "reference_group_key": "PV02.TEST|B700.TEST|group|A1",
                "root_product_code": "PV02.TEST",
                "source_parent_code": "B700.TEST",
                "source_file_code": "B700.TEST",
                "source_file_name": "B700.TEST.xlsx",
                "source_kind": "btp",
                "substitute_group": "A1",
                "component_code": "012.0004200",
                "substitute_materials_raw": "",
                "substitute_priority": "2",
                "substitute_strategy": "1",
                "usage_probability": "100",
                "same_level_group_marker": "B700.TEST/A1/00/00",
                "substitute_combo_marker": "A1-01-00-00",
                "temp_substitute_row_no": "",
                "signal_class": "group_and_combo_semantics",
                "position_no": "NTC2",
            },
        ]

        groups = builder.build_reference_groups(row_signals)

        self.assertEqual(len(groups), 1)
        group = groups[0]
        self.assertEqual(group["member_row_count"], 2)
        self.assertEqual(group["member_component_count"], 2)
        self.assertEqual(group["explicit_substitute_material_count"], 2)
        self.assertEqual(group["priority_values"], "1;2")
        self.assertEqual(group["strategy_values"], "1")
        self.assertEqual(group["usage_probability_values"], "100")
        self.assertEqual(group["same_level_group_markers"], "B700.TEST/A1/00/00")
        self.assertEqual(group["substitute_combo_markers"], "A1-01-00-00")


if __name__ == "__main__":
    unittest.main()
