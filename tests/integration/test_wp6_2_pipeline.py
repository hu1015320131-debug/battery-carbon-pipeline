from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from carbon_excel_pipeline.capability.pipeline import run_wp6_2_capability_detection


class WP62PipelineTests(unittest.TestCase):
    def test_pipeline_consumes_only_wp6_1_artifacts_and_writes_required_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            run = Path(temp) / "RUN-WP62-SYNTHETIC"
            import_dir = run / "01_import"
            (run / "02_capability").mkdir(parents=True)
            import_dir.mkdir()
            fingerprint = "A" * 64
            recognition = {
                "recognition_status": "RECOGNIZED",
                "input_fingerprint": fingerprint,
                "best_candidate_sheet": "Data",
                "best_candidate_header_row": 1,
            }
            field_mappings = [
                {
                    "semantic_field": "Reported_Activity_Value",
                    "detected_unit": "kg/year",
                },
                {"semantic_field": "EF_Value", "detected_unit": "kgCO2e/kg"},
            ]
            mappings = [
                {
                    "sheet_name": "Data",
                    "header_row": 1,
                    "field_mappings": field_mappings,
                }
            ]
            records = {
                "workbook_name": "synthetic.xlsx",
                "input_fingerprint": fingerprint,
                "sheet_name": "Data",
                "sheet_index": 0,
                "header_row": 1,
                "denominator_definition": "synthetic records",
                "units": {
                    "Reported_Activity_Value": "kg/year",
                    "EF_Value": "kgCO2e/kg",
                },
                "records": [
                    {
                        "Source_Row": 2,
                        "values": {"Reported_Activity_Value": 10, "EF_Value": 2},
                        "formula_fields": [],
                    }
                ],
            }
            for name, payload in (
                ("recognition_summary.json", recognition),
                ("semantic_field_mapping.json", mappings),
                ("recognized_records.json", records),
            ):
                (import_dir / name).write_text(
                    json.dumps(payload, ensure_ascii=False), encoding="utf-8"
                )
            result = run_wp6_2_capability_detection(run)
            self.assertEqual(result["status"], "PASS")
            required = {
                "dataset_capabilities.json",
                "record_capabilities.csv",
                "activity_path_decisions.csv",
                "capability_summary.json",
                "WP6-2_数据能力识别报告.md",
            }
            self.assertTrue(required.issubset({item.name for item in (run / "02_capability").iterdir()}))


if __name__ == "__main__":
    unittest.main()
