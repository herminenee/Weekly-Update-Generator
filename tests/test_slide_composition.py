from contextlib import redirect_stderr, redirect_stdout
from datetime import date
from io import BytesIO
from io import StringIO
from pathlib import Path
import runpy
import unittest

from pptx import Presentation


def load_app_namespace():
    with StringIO() as log:
        with redirect_stdout(log), redirect_stderr(log):
            return runpy.run_path(str(Path(__file__).resolve().parents[1] / "app.py"))


class SlideCompositionTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = load_app_namespace()

    def short_projects(self):
        labels = self.app["LABELS"]
        return [
            {
                "id": "alpha",
                "name": "Alpha",
                "sections": {label: "" for label in labels} | {
                    "Objectives": "Assess coating stability.",
                    "Experimental Work": "Stored samples at 25 °C for 24 h.",
                },
            },
            {
                "id": "beta",
                "name": "Beta",
                "sections": {label: "" for label in labels} | {
                    "Results & Observations": "Signal remained stable.",
                    "Next Steps": "Repeat the measurement.",
                },
            },
        ]

    def test_concise_reduces_slides_and_preserves_project_order(self):
        plan = self.app["build_slide_plan"](self.short_projects(), mode="Concise")
        self.app["validate_slide_plan"](plan, self.short_projects())
        self.assertEqual([slide["project"] for slide in plan if slide["project"]], ["Alpha", "Beta"])
        self.assertEqual(len(plan), 3)

        data = self.app["make_ppt"](
            self.short_projects(), "Synthetic report", date(2026, 10, 12), "Tester", "", None, "Concise"
        )
        deck = Presentation(BytesIO(data))
        texts = "\n".join(
            shape.text for slide in deck.slides for shape in slide.shapes if hasattr(shape, "text")
        )
        self.assertEqual(len(deck.slides), 3)
        self.assertIn("Assess coating stability.", texts)
        self.assertIn("Repeat the measurement.", texts)
        self.assertIn("Alpha", texts)
        self.assertIn("Beta", texts)

    def test_optional_project_dividers_and_template_dimensions(self):
        projects = self.short_projects()
        without_dividers = self.app["build_slide_plan"](projects, mode="Concise", include_dividers=False)
        with_dividers = self.app["build_slide_plan"](projects, mode="Concise", include_dividers=True)
        self.assertEqual(len(without_dividers), 3)
        self.assertEqual(len(with_dividers), 3)

        labels = self.app["LABELS"]
        long_project = {
            "id": "divider",
            "name": "Divider Project",
            "sections": {label: "" for label in labels} | {
                "Experimental Work": "\n".join(f"Long detail {i} remains." for i in range(12)),
            },
        }
        no_divider = self.app["build_slide_plan"]([long_project], mode="Detailed", include_dividers=False)
        divider = self.app["build_slide_plan"]([long_project], mode="Detailed", include_dividers=True)
        self.assertEqual(len(divider), len(no_divider) + 1)
        self.assertEqual(divider[1]["content"], [])

        template = Presentation()
        template.slide_width = self.app["Inches"](10)
        template.slide_height = self.app["Inches"](7)
        template_bytes = BytesIO()
        template.save(template_bytes)
        data = self.app["make_ppt"](
            projects, "Synthetic report", date(2026, 10, 12), "Tester", "",
            template_bytes.getvalue(), "Concise", False,
        )
        imported = Presentation(BytesIO(data))
        self.assertEqual(imported.slide_width, template.slide_width)
        self.assertEqual(imported.slide_height, template.slide_height)

    def test_detailed_splits_long_content_without_omission(self):
        labels = self.app["LABELS"]
        long_lines = [f"Observation {index} recorded at {index} nm." for index in range(1, 13)]
        project = {
            "id": "long",
            "name": "Long Project",
            "sections": {label: "" for label in labels} | {
                "Results & Observations": "\n".join(long_lines),
                "Discussion": "The cause was not determined.",
            },
        }
        concise = self.app["build_slide_plan"]([project], mode="Concise")
        detailed = self.app["build_slide_plan"]([project], mode="Detailed")
        self.app["validate_slide_plan"](concise, [project])
        self.app["validate_slide_plan"](detailed, [project])
        self.assertGreaterEqual(len(detailed), len(concise))
        detailed_text = "\n".join(line for slide in detailed for _, line in slide["content"])
        for line in long_lines:
            self.assertIn(line, detailed_text)
        self.assertIn("The cause was not determined.", detailed_text)
        self.assertGreater(len(detailed), 2)

    def test_titles_are_project_only_and_long_text_stays_inside_safe_bounds(self):
        labels = self.app["LABELS"]
        paragraph = " ".join(
            f"Long experimental observation {index} remains approved text." for index in range(30)
        )
        project = {
            "id": "long-title",
            "name": "Very Long Project Name Used Only As A Primary Title",
            "sections": {label: "" for label in labels} | {
                "Experimental Work": paragraph,
                "Results & Observations": "Observed result was not interpreted.",
            },
        }
        plan = self.app["build_slide_plan"]([project], mode="Detailed")
        self.app["validate_slide_plan"](plan, [project])
        self.assertTrue(all("Experimental Work" not in slide["title"] for slide in plan))
        self.assertTrue(all("Results & Observations" not in slide["title"] for slide in plan))
        titles = [slide["title"] for slide in plan if slide["project"]]
        self.assertTrue(all(title.startswith(project["name"]) for title in titles))

        data = self.app["make_ppt"](
            [project], "Synthetic report", date(2026, 10, 12), "Tester", "", None, "Detailed"
        )
        deck = Presentation(BytesIO(data))
        footer_top = self.app["SLIDE_HEIGHT"] - self.app["FOOTER_HEIGHT"] - self.app["BOTTOM_MARGIN"]
        for slide in list(deck.slides)[1:]:
            for shape in slide.shapes:
                if not hasattr(shape, "text") or not shape.text:
                    continue
                top = shape.top / 914400
                bottom = (shape.top + shape.height) / 914400
                if "Tester" in shape.text:
                    self.assertGreaterEqual(top, footer_top - .01)
                else:
                    self.assertLessEqual(bottom, footer_top + .01)

        body_chunks = [
            line for slide in plan for label, line in slide["content"]
            if label == "Experimental Work"
        ]
        self.assertEqual("".join(body_chunks), paragraph)

    def test_reference_style_assigns_verified_fonts_colors_and_dimensions(self):
        projects = self.short_projects()
        data = self.app["make_ppt"](
            projects, "Reference style report", date(2026, 10, 12), "Tester", "",
            None, "Concise", False, "Reference Style",
        )
        deck = Presentation(BytesIO(data))
        self.assertAlmostEqual(deck.slide_width / 914400, 13.333, places=2)
        self.assertAlmostEqual(deck.slide_height / 914400, 7.5, places=2)

        content_shapes = [
            shape for slide in deck.slides for shape in slide.shapes
            if hasattr(shape, "text") and shape.text
        ]
        title_shape = next(shape for shape in content_shapes if shape.text == "Alpha")
        heading_shape = next(shape for shape in content_shapes if shape.text == "Objectives")
        body_shape = next(shape for shape in content_shapes if "Assess coating stability." in shape.text)
        footer_shape = next(shape for shape in content_shapes if "Tester" in shape.text)

        self.assertEqual(title_shape.text_frame.paragraphs[0].runs[0].font.name, "Times New Roman")
        self.assertEqual(heading_shape.text_frame.paragraphs[0].runs[0].font.name, "Bitstream Charter")
        self.assertEqual(body_shape.text_frame.paragraphs[0].runs[0].font.name, "Bitstream Charter")
        self.assertEqual(footer_shape.text_frame.paragraphs[0].runs[0].font.name, "Bitstream Charter")

        def rgb(shape):
            return str(shape.text_frame.paragraphs[0].runs[0].font.color.rgb)

        self.assertEqual(rgb(title_shape), "17324D")
        self.assertEqual(rgb(heading_shape), "C00000")
        self.assertEqual(rgb(body_shape), "17324D")
        self.assertEqual(rgb(footer_shape), "667784")
        self.assertAlmostEqual(title_shape.left / 914400, 0.70, places=2)
        self.assertLess(body_shape.top + body_shape.height, footer_shape.top)

    def test_reference_style_paginates_larger_text_without_content_loss(self):
        labels = self.app["LABELS"]
        paragraph = "\n".join(
            f"Reference-mode observation {index} remains approved at {index} nm."
            for index in range(18)
        )
        project = {
            "id": "reference-long",
            "name": "Reference Long Project",
            "sections": {label: "" for label in labels} | {
                "Results & Observations": paragraph,
            },
        }
        plan = self.app["build_slide_plan"](
            [project], mode="Detailed", style_name="Reference Style"
        )
        self.app["validate_slide_plan"](plan, [project])
        planned = "".join(line for slide in plan for _, line in slide["content"])
        self.assertEqual(planned, "".join(self.app["section_lines"](paragraph)))
        self.assertGreater(len(plan), 2)

    def test_safe_default_style_remains_selectable(self):
        style = self.app["resolve_style"]("Safe Default Style")
        self.assertEqual(style["name"], "Safe Default Style")
        self.assertEqual(style["title_font"], "Aptos")

    def test_short_content_uses_checked_two_column_layout(self):
        project = self.short_projects()[0]
        plan = self.app["build_slide_plan"]([project], mode="Concise", style_name="Reference Style")
        content_slide = next(slide for slide in plan if slide["project"])
        self.assertEqual(content_slide["layout"], "compact")

        data = self.app["make_ppt"](
            [project], "Compact report", date(2026, 10, 12), "Tester", "",
            None, "Concise", False, "Reference Style",
        )
        deck = Presentation(BytesIO(data))
        slide = list(deck.slides)[1]
        body_shapes = [
            shape for shape in slide.shapes
            if hasattr(shape, "text") and shape.text.startswith("• ")
        ]
        self.assertEqual(len(body_shapes), 2)
        self.assertNotEqual(body_shapes[0].left, body_shapes[1].left)
        footer_top = next(
            shape.top for shape in slide.shapes
            if hasattr(shape, "text") and "Tester" in shape.text
        )
        self.assertTrue(all(shape.top + shape.height < footer_top for shape in body_shapes))

    def test_long_content_stays_single_column_and_preserves_exact_text(self):
        labels = self.app["LABELS"]
        paragraph = "\n".join(
            f"Long paragraph observation {index} remains unchanged at {index} nm."
            for index in range(20)
        )
        project = {
            "id": "single-column",
            "name": "Single Column Project",
            "sections": {label: "" for label in labels} | {
                "Experimental Work": paragraph,
            },
        }
        plan = self.app["build_slide_plan"]([project], mode="Detailed", style_name="Reference Style")
        self.assertTrue(all(slide["layout"] == "single" for slide in plan if slide["project"]))
        self.app["validate_slide_plan"](plan, [project])
        actual = "".join(line for slide in plan for _, line in slide["content"])
        self.assertEqual(actual, "".join(self.app["section_lines"](paragraph)))

    def test_thin_film_five_sections_use_adaptive_reference_pagination(self):
        labels = self.app["LABELS"]
        project = {
            "id": "thin-film",
            "name": "Thin-Film Stability",
            "sections": {label: "" for label in labels} | {
                "Objectives": "Evaluate thin-film stability at 25 °C and 60 °C.",
                "Experimental Work": (
                    "Samples were stored at 60 °C for 48 h; "
                    "thickness was measured every 12 h."
                ),
                "Results & Observations": (
                    "Initial thickness was 120 nm and was 117 nm after 48 h."
                ),
                "Discussion": (
                    "The decrease may indicate a small structural change; "
                    "the cause was not determined."
                ),
                "Next Steps": (
                    "Repeat the experiment three times and add a 25 °C control group."
                ),
            },
        }
        plan = self.app["build_slide_plan"](
            [project], mode="Concise", style_name="Reference Style"
        )
        self.app["validate_slide_plan"](plan, [project])
        content_slides = [slide for slide in plan if slide["project"]]
        self.assertEqual(len(content_slides), 1)
        self.assertEqual(content_slides[0]["layout"], "compact")
        self.assertEqual(
            [label for label, _ in content_slides[0]["content"]],
            [
                "Objectives",
                "Experimental Work",
                "Results & Observations",
                "Discussion",
                "Next Steps",
            ],
        )

        data = self.app["make_ppt"](
            [project], "Thin-Film Weekly Update", date(2026, 10, 12),
            "Tester", "", None, "Concise", False, "Reference Style",
        )
        deck = Presentation(BytesIO(data))
        self.assertEqual(len(deck.slides), 2)
        exported = "\n".join(
            shape.text for slide in deck.slides for shape in slide.shapes
            if hasattr(shape, "text")
        )
        for section_text in project["sections"].values():
            if section_text:
                self.assertIn(section_text, exported)
