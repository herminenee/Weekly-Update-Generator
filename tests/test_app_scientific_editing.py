import json
from pathlib import Path
import unittest
from unittest.mock import patch

from scientific_writer import LABELS, ScientificWriterError

from streamlit.testing.v1 import AppTest


class AppScientificEditingTest(unittest.TestCase):
    def test_notes_are_displayed_and_failure_preserves_project_content(self):
        app_path = str(Path(__file__).resolve().parents[1] / "app.py")
        app = AppTest.from_file(app_path).run()
        self.assertFalse(app.exception)

        notes = "结果：初始厚度为 120 nm，48 h 后为 117 nm。"
        with patch(
            "scientific_writer._request_ollama",
            side_effect=ScientificWriterError("synthetic local failure"),
        ):
            app.text_area(key="raw_notes").set_value(notes).run()
            self.assertFalse(app.exception)
            self.assertIn(notes, app.code[0].value)
            app.button(key="run_local_editing").click().run()

        self.assertFalse(app.exception)
        self.assertIn(notes, app.code[0].value)
        self.assertEqual(app.session_state.projects[0]["sections"]["Experimental Work"], "")

    def test_missing_source_section_warning_renders_without_exception(self):
        app_path = str(Path(__file__).resolve().parents[1] / "app.py")
        app = AppTest.from_file(app_path).run()
        notes = (
            "目标：评估薄膜稳定性。\n"
            "实验：样品在 25°C 下保存 24 h。\n"
            "结果：信号保持稳定。\n"
            "讨论：原因尚未确定。"
        )
        response_sections = {
            label: "" for label in LABELS
        }
        response_sections.update({
            "Experimental Work": "Samples were stored at 25 °C for 24 h.",
            "Results & Observations": "Signal remained stable.",
            "Discussion": "The cause was not determined.",
        })

        def mocked_request(notes_value, opener):
            self.assertEqual(notes_value, notes)
            return json.dumps(response_sections)

        with patch("scientific_writer._request_ollama", side_effect=mocked_request):
            app.text_area(key="raw_notes").set_value(notes).run()
            app.button(key="run_local_editing").click().run()

        self.assertFalse(app.exception)
        self.assertIn("Objectives", "\n".join(item.value for item in app.error))
        self.assertIn("评估薄膜稳定性", "\n".join(item.value for item in app.code))
        self.assertIn("missing_source_sections", app.session_state.ai_result.validation.__dict__)
        self.assertEqual(
            app.session_state.ai_result.validation.missing_source_sections,
            ("Objectives",),
        )
