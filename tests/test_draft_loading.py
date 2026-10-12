import json
from datetime import date
from pathlib import Path
import tempfile
import unittest

from streamlit.testing.v1 import AppTest


class DraftLoadingTest(unittest.TestCase):
    def test_load_draft_after_widgets_rendered_and_switch_projects(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            tmp_path = Path(temporary_directory)
            draft_path = tmp_path / "saved.json"
            draft_path.write_text(
                json.dumps(
                    {
                        "version": 1,
                        "title": "Loaded weekly report",
                        "author": "Loaded author",
                        "report_date": date(2026, 9, 30).isoformat(),
                        "previous": "Review calibration",
                        "projects": [
                            {
                                "id": "loaded-alpha",
                                "name": "Loaded Alpha",
                                "sections": {
                                    "Objectives": "Alpha objective",
                                    "Experimental Work": "Alpha work",
                                    "Results & Observations": "",
                                    "Discussion": "",
                                    "Challenges": "",
                                    "Next Steps": "Alpha next",
                                },
                            },
                            {
                                "id": "loaded-beta",
                                "name": "Loaded Beta",
                                "sections": {
                                    "Objectives": "Beta objective",
                                    "Experimental Work": "",
                                    "Results & Observations": "Beta result",
                                    "Discussion": "",
                                    "Challenges": "",
                                    "Next Steps": "",
                                },
                            },
                        ],
                    }
                ),
                encoding="utf-8",
            )

            app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "app.py")).run()
            self.assertFalse(app.exception)

            app.text_input(key="draft_folder").set_value(str(tmp_path)).run()
            self.assertFalse(app.exception)
            app.button(key="load_draft").click().run()

            self.assertFalse(app.exception)
            self.assertEqual(app.text_input(key="report_title").value, "Loaded weekly report")
            self.assertEqual(app.text_input(key="report_author").value, "Loaded author")
            self.assertEqual(app.date_input(key="report_date").value, date(2026, 9, 30))
            self.assertEqual(app.text_area(key="previous_next_steps").value, "Review calibration")
            self.assertEqual(app.session_state.projects[0]["name"], "Loaded Alpha")
            self.assertEqual(app.session_state.projects[1]["name"], "Loaded Beta")
            self.assertEqual(app.text_area(key="section_loaded-alpha_1_0").value, "Alpha objective")

            app.button(key="select_project_loaded-beta").click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state.selected_project, 1)
            self.assertEqual(app.text_area(key="section_loaded-beta_1_0").value, "Beta objective")

            app.selectbox(key="selected_draft").select(draft_path.name).run()
            app.button(key="load_draft").click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.text_input(key="report_title").value, "Loaded weekly report")
            self.assertEqual(app.session_state.selected_project, 0)
            self.assertEqual(app.session_state.projects[0]["sections"]["Next Steps"], "Alpha next")
