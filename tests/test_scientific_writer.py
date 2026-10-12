from io import BytesIO
import json
from urllib.error import URLError
import unittest

from scientific_writer import (
    LABELS,
    ScientificWriterError,
    ScientificWriterResponseError,
    edit_notes,
    extract_source_sections,
    parse_response,
    validate_measurements,
)


NOTES = """项目：薄膜稳定性测试
目标：评估 25 °C 和 60 °C 条件下薄膜的稳定性。
实验工作：样品在 60 °C 下保存 48 h。每 12 h 测量厚度。
结果：初始厚度为 120 nm，48 h 后为 117 nm。
讨论：The thickness decrease may indicate a small structural change, but the cause was not determined.
下一步：重复实验三次，并增加 25 °C 对照组。"""


def fake_opener(request, timeout):
    request_body = json.loads(request.data.decode("utf-8"))
    assert request_body["format"]["required"] == list(LABELS)
    payload = {
        "response": json.dumps({
            "Objectives": "Evaluate thin-film stability at 25 °C and 60 °C.",
            "Experimental Work": "Samples were stored at 60 °C for 48 h; thickness was measured every 12 h.",
            "Results & Observations": "Initial thickness was 120 nm and was 117 nm after 48 h.",
            "Discussion": "The decrease may indicate a small structural change; the cause was not determined.",
            "Challenges": "The control condition is ambiguous.",
            "Next Steps": "Repeat the experiment three times and add a 25 °C control group.",
        }),
        "done": True,
    }
    return BytesIO(json.dumps(payload).encode("utf-8"))


class ScientificWriterTest(unittest.TestCase):
    def test_local_response_parsing_and_faithful_sections(self):
        result = edit_notes(NOTES, fake_opener)
        self.assertEqual(result.sections["Results & Observations"].count("120 nm"), 1)
        self.assertFalse(result.validation.has_warnings)
        self.assertIn("cause was not determined", result.sections["Discussion"])

    def test_missing_measurement_is_flagged(self):
        report = validate_measurements(NOTES, "The thickness was 117 nm after 48 h.")
        self.assertTrue(report.has_warnings)
        self.assertTrue(any("120 nm" in warning for warning in report.warnings))

    def test_malformed_response_is_rejected(self):
        with self.assertRaisesRegex(ScientificWriterResponseError, "invalid JSON"):
            parse_response("not json")

    def test_incomplete_and_non_string_responses_are_rejected(self):
        with self.assertRaisesRegex(ScientificWriterResponseError, "missing fields"):
            parse_response(json.dumps({"Objectives": "Only one section"}))
        incomplete = {label: "" for label in LABELS}
        incomplete["Discussion"] = 7
        with self.assertRaisesRegex(ScientificWriterResponseError, "non-string"):
            parse_response(json.dumps(incomplete))

    def test_invalid_api_json_is_distinguished_from_connectivity(self):
        def invalid_api_json(request, timeout):
            return BytesIO(b"not-json")

        with self.assertRaisesRegex(ScientificWriterResponseError, "API response level"):
            edit_notes(NOTES, invalid_api_json)

    def test_truncated_ollama_response_is_rejected(self):
        def truncated(request, timeout):
            return BytesIO(json.dumps({"response": "{}", "done": False}).encode("utf-8"))

        with self.assertRaisesRegex(ScientificWriterResponseError, "incomplete"):
            edit_notes(NOTES, truncated)

    def test_unavailable_ollama_is_reported(self):
        def unavailable(request, timeout):
            raise URLError("offline")

        with self.assertRaisesRegex(ScientificWriterError, "127.0.0.1:11434"):
            edit_notes(NOTES, unavailable)

    def test_chinese_headings_and_missing_objectives_are_reported(self):
        notes = "目标：评估涂层稳定性。\n实验：样品保存 25°C 24 h。\n结果：信号保持稳定。"
        sections = extract_source_sections(notes)
        self.assertEqual(sections["Objectives"], "评估涂层稳定性。")
        self.assertEqual(sections["Experimental Work"], "样品保存 25°C 24 h。")

        def missing_objectives(request, timeout):
            payload = {
                "response": json.dumps({
                    "Objectives": "",
                    "Experimental Work": "Samples stored at 25 °C for 24 h.",
                    "Results & Observations": "Signal remained stable.",
                    "Discussion": "",
                    "Challenges": "",
                    "Next Steps": "",
                }),
                "done": True,
            }
            return BytesIO(json.dumps(payload).encode("utf-8"))

        result = edit_notes(notes, missing_objectives)
        self.assertEqual(result.validation.missing_source_sections, ("Objectives",))
        self.assertEqual(result.validation.source_excerpts[0], ("Objectives", "评估涂层稳定性。"))
        self.assertEqual(result.sections["Objectives"], "")

    def test_measurement_normalization_and_duplicate_categories(self):
        normalized = validate_measurements("25°C; 25 °C; 120 nm", "25 °C; 25°C; 120 nm")
        self.assertFalse(normalized.has_warnings)

        missing = validate_measurements("25°C; 25°C", "25 °C")
        self.assertTrue(any("Duplicate-count difference" in warning for warning in missing.warnings))

        added = validate_measurements("25°C", "25 °C; 30 °C")
        self.assertTrue(any("Added measurement" in warning for warning in added.warnings))

        altered = validate_measurements("25°C", "30°C")
        self.assertTrue(any("Altered measurement" in warning for warning in altered.warnings))
