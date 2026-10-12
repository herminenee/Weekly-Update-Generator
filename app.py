"""Local-first Weekly Update Generator. No network calls or AI APIs."""
from __future__ import annotations

from datetime import date
from io import BytesIO
import json
import re
from pathlib import Path
from uuid import uuid4

import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from scientific_writer import LABELS as WRITER_LABELS
from scientific_writer import ScientificWriterError, edit_notes

st.set_page_config(page_title="Weekly Update Generator", page_icon="🧪", layout="wide")

HEADINGS = [
    ("Objectives", r"^(objective|aim|goal|purpose|research question)s?\s*[:：]|^(目标|目的)\s*[:：]"),
    ("Experimental Work", r"^(experimental|experiment|procedure|method|protocol|preparation|work performed)s?\s*[:：]|^(实验|实验工作|方法)\s*[:：]"),
    ("Results & Observations", r"^(result|observation|finding|measurement|data)s?\s*[:：]|^(结果|观察|数据)\s*[:：]"),
    ("Discussion", r"^(discussion|analysis|interpretation|hypothesis)s?\s*[:：]|^(讨论|分析|解释)\s*[:：]"),
    ("Challenges", r"^(challenge|issue|problem|limitation|difficulty)s?\s*[:：]|^(问题|挑战|困难)\s*[:：]"),
    ("Next Steps", r"^(next step|next week|future work|plan|to do|todo)s?\s*[:：]|^(下一步|后续工作|计划)\s*[:：]"),
]
LABELS = [name for name, _ in HEADINGS]
REFERENCE_TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "Weekly_Update_20261005.pptx"

SAFE_DEFAULT_STYLE = {
    "name": "Safe Default Style",
    "title_font": "Aptos",
    "heading_font": "Aptos",
    "body_font": "Aptos",
    "footer_font": "Aptos",
    "title_size": 29,
    "heading_size": 13,
    "body_size": 17,
    "footer_size": 11,
    "title_color": RGBColor(26, 37, 53),
    "heading_color": RGBColor(37, 83, 124),
    "body_color": RGBColor(26, 37, 53),
    "footer_color": RGBColor(91, 106, 121),
    "left_margin": 0.75,
    "right_margin": 0.75,
    "top_margin": 0.85,
    "title_height": 0.65,
    "body_top": 1.65,
    "footer_height": 0.42,
    "footer_gap": 0.18,
    "bottom_margin": 0.28,
    "body_chars_per_line": 68,
}

REFERENCE_STYLE = {
    "name": "Reference Style",
    "title_font": "Times New Roman",
    "heading_font": "Bitstream Charter",
    "body_font": "Bitstream Charter",
    "footer_font": "Bitstream Charter",
    "title_size": 25.5,
    "heading_size": 20.25,
    "body_size": 20.25,
    "footer_size": 13.5,
    "title_color": RGBColor(23, 50, 77),
    "heading_color": RGBColor(192, 0, 0),
    "body_color": RGBColor(23, 50, 77),
    "footer_color": RGBColor(102, 119, 132),
    "left_margin": 0.70,
    "right_margin": 0.70,
    "top_margin": 0.40,
    "title_height": 0.62,
    "body_top": 1.55,
    "footer_height": 0.32,
    "footer_gap": 0.18,
    "bottom_margin": 0.30,
    "body_chars_per_line": 54,
}

STYLE_NAMES = ("Reference Style", "Safe Default Style")


def reference_style_available() -> bool:
    return REFERENCE_TEMPLATE_PATH.is_file()


def resolve_style(style_name: str | None = None) -> dict[str, object]:
    selected = style_name or ("Reference Style" if reference_style_available() else "Safe Default Style")
    if selected == "Reference Style" and reference_style_available():
        return REFERENCE_STYLE
    return SAFE_DEFAULT_STYLE


def decode_upload(upload):
    if not upload:
        return ""
    raw = upload.getvalue()
    for codec in ("utf-8-sig", "utf-16", "latin-1"):
        try:
            return raw.decode(codec)
        except UnicodeDecodeError:
            pass
    return ""


def split_notes(text: str) -> dict[str, str]:
    """Conservative sectioning; unclassified text is preserved in Experimental Work."""
    result = {name: [] for name in LABELS}
    current = "Experimental Work"
    for line in text.splitlines():
        stripped = line.strip().lstrip("#").strip()
        found = False
        for name, pattern in HEADINGS:
            match = re.match(pattern, stripped, flags=re.I)
            if match:
                current = name
                remaining = stripped[match.end():].strip()
                if remaining:
                    result[current].append(remaining)
                found = True
                break
        if not found and stripped:
            result[current].append(stripped)
    return {name: "\n".join(lines) for name, lines in result.items()}


def clean_lines(text: str) -> list[str]:
    return [re.sub(r"^\s*[-*•\d.)]+\s*", "", ln).strip() for ln in text.splitlines() if ln.strip()]


def text_box(slide, x, y, w, h, value, size=18, color=None, bold=False, font_name="Aptos"):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    paragraph = tf.paragraphs[0]
    run = paragraph.add_run()
    run.text = value
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color or SAFE_DEFAULT_STYLE["body_color"]
    return shape


def empty_project(number: int) -> dict[str, object]:
    return {
        "id": f"project-{uuid4().hex}",
        "name": f"Research Project {number}",
        "sections": {name: "" for name in LABELS},
    }


def normalize_projects(projects: list[dict[str, object]]) -> list[dict[str, object]]:
    normalized = []
    used_ids = set()
    for index, project in enumerate(projects, start=1):
        raw_project = project if isinstance(project, dict) else {}
        sections = raw_project.get("sections", {})
        project_id = str(raw_project.get("id") or f"project-{uuid4().hex}")
        if project_id in used_ids:
            project_id = f"{project_id}-{uuid4().hex}"
        used_ids.add(project_id)
        normalized.append({
            "id": project_id,
            "name": str(raw_project.get("name") or f"Research Project {index}"),
            "sections": {
                label: str(sections.get(label, "")) if isinstance(sections, dict) else ""
                for label in LABELS
            },
        })
    return normalized or [empty_project(1)]


def draft_payload(projects, title, author, report_date, previous,
                  presentation_mode="Concise", include_dividers=False,
                  presentation_style=None):
    return {
        "version": 2,
        "title": title,
        "author": author,
        "report_date": report_date.isoformat(),
        "previous": previous,
        "presentation_mode": presentation_mode,
        "include_dividers": include_dividers,
        "presentation_style": presentation_style or (
            "Reference Style" if reference_style_available() else "Safe Default Style"
        ),
        "projects": projects,
    }


def load_draft(path: Path) -> dict[str, object]:
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict) or not isinstance(payload.get("projects"), list):
        raise ValueError("Draft file must contain a projects list.")
    return payload


def save_draft(folder: Path, payload: dict[str, object]) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    stamp = date.today().strftime("%Y%m%d")
    candidate = folder / f"weekly_update_{stamp}.json"
    counter = 2
    while candidate.exists():
        candidate = folder / f"weekly_update_{stamp}_{counter}.json"
        counter += 1
    with candidate.open("x", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    return candidate


def move_project(projects: list[dict[str, object]], index: int, offset: int) -> int:
    target = index + offset
    if 0 <= target < len(projects):
        projects[index], projects[target] = projects[target], projects[index]
        return target
    return index


SECTION_ORDER = tuple(LABELS)
CONCISE_BUDGET = 520
DETAILED_BUDGET = 300
SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5
LEFT_MARGIN = SAFE_DEFAULT_STYLE["left_margin"]
RIGHT_MARGIN = SAFE_DEFAULT_STYLE["right_margin"]
TOP_MARGIN = SAFE_DEFAULT_STYLE["top_margin"]
TITLE_HEIGHT = SAFE_DEFAULT_STYLE["title_height"]
BODY_TOP = SAFE_DEFAULT_STYLE["body_top"]
FOOTER_HEIGHT = SAFE_DEFAULT_STYLE["footer_height"]
FOOTER_GAP = SAFE_DEFAULT_STYLE["footer_gap"]
BOTTOM_MARGIN = SAFE_DEFAULT_STYLE["bottom_margin"]
BODY_BOTTOM = SLIDE_HEIGHT - FOOTER_HEIGHT - FOOTER_GAP - BOTTOM_MARGIN
BODY_HEIGHT = BODY_BOTTOM - BODY_TOP
BODY_WIDTH = SLIDE_WIDTH - LEFT_MARGIN - RIGHT_MARGIN
BODY_FONT_SIZE = SAFE_DEFAULT_STYLE["body_size"]
SECTION_FONT_SIZE = SAFE_DEFAULT_STYLE["heading_size"]


def estimate_lines(text: str, chars_per_line: int = 68) -> int:
    """Conservative line estimate; PowerPoint remains the final renderer."""
    return max(1, sum(max(1, (len(line) + chars_per_line - 1) // chars_per_line)
                       for line in text.splitlines() or [""]))


def split_text_for_layout(text: str, max_chars: int) -> list[str]:
    if len(text) <= max_chars:
        return [text]
    return [text[start:start + max_chars] for start in range(0, len(text), max_chars)]


def section_lines(content: object) -> list[str]:
    return clean_lines(str(content)) if str(content).strip() else []


def content_item_height(line: str, style: dict[str, object], chars_per_line: int) -> float:
    body_lines = estimate_lines(line, chars_per_line)
    body_height = max(
        0.52,
        body_lines * (float(style["body_size"]) / 72) * 1.25 + 0.08,
    )
    heading_height = float(style["heading_size"]) / 72 * 1.25
    return heading_height + body_height + 0.12


def body_layout_capacity(style: dict[str, object]) -> float:
    return (
        SLIDE_HEIGHT
        - float(style["footer_height"])
        - float(style["bottom_margin"])
        - float(style["body_top"])
        - float(style["footer_gap"])
    )


def content_fits_single_column(content: list[tuple[str, str]], style: dict[str, object]) -> bool:
    return sum(
        content_item_height(line, style, int(style["body_chars_per_line"]))
        for _, line in content
    ) <= body_layout_capacity(style) + 0.01


def compact_column_groups(content: list[tuple[str, str]]) -> list[list[tuple[str, str]]]:
    midpoint = (len(content) + 1) // 2
    return [content[:midpoint], content[midpoint:]]


def build_slide_plan(projects: list[dict[str, object]], previous: str = "",
                     mode: str = "Concise", include_dividers: bool = False,
                     style_name: str | None = None) -> list[dict[str, object]]:
    """Build a deterministic, lossless slide plan from approved project text."""
    style = resolve_style(style_name)
    budget = CONCISE_BUDGET if mode == "Concise" else DETAILED_BUDGET
    plan = [{"title": "Cover", "project": "", "sections": [], "content": []}]

    previous_lines = section_lines(previous)
    for start in range(0, len(previous_lines), 5):
        plan.append({
            "title": "Previous Week's Next Steps" + (" (continued)" if start else ""),
            "project": "",
            "sections": ["Previous Week's Next Steps"],
            "content": [("Previous Week's Next Steps", line) for line in previous_lines[start:start + 5]],
        })

    for project in projects:
        project_name = str(project.get("name", "")).strip() or "Untitled Research Project"
        sections = project.get("sections", {})
        if not isinstance(sections, dict):
            continue
        populated = [(label, section_lines(sections.get(label, "")))
                     for label in SECTION_ORDER if section_lines(sections.get(label, ""))]
        if not populated:
            continue

        section_groups = []
        current_labels = []
        current_lines = []
        current_length = 0
        project_line_count = sum(len(lines) for _, lines in populated)
        project_char_count = sum(len(line) for _, lines in populated for line in lines)
        project_compact_eligible = project_line_count <= 6 and project_char_count <= 420
        current_requires_single = not project_compact_eligible
        max_items = 6 if mode == "Concise" else 4
        for label, lines in populated:
            for line in lines:
                chunks = split_text_for_layout(line, 300 if mode == "Concise" else 210)
                for chunk in chunks:
                    item_length = len(label) + len(chunk) + 2
                    candidate = current_lines + [(label, chunk)]
                    candidate_slide = {"content": candidate}
                    candidate_requires_single = current_requires_single or len(chunks) > 1
                    candidate_fits = content_fits_single_column(candidate, style)
                    if not candidate_requires_single and project_compact_eligible:
                        candidate_fits = candidate_fits or compact_layout_allowed(
                            candidate_slide, style
                        )
                    if current_lines and (
                        current_length + item_length > budget
                        or len(current_lines) >= max_items
                        or not candidate_fits
                    ):
                        section_groups.append((
                            current_labels, current_lines, current_requires_single
                        ))
                        current_labels, current_lines, current_length = [], [], 0
                        current_requires_single = not project_compact_eligible
                    if label not in current_labels:
                        current_labels.append(label)
                    current_lines.append((label, chunk))
                    current_length += item_length
                    current_requires_single = current_requires_single or len(chunks) > 1
        if current_lines:
            section_groups.append((current_labels, current_lines, current_requires_single))

        # A divider is useful only when this project needs more than one content slide.
        if include_dividers and len(section_groups) > 1:
            plan.append({
                "title": project_name,
                "project": project_name,
                "sections": [],
                "content": [],
            })
        for group_index, (group_labels, group_content, requires_single) in enumerate(section_groups):
            title = project_name if len(section_groups) == 1 else f"{project_name} — Part {group_index + 1}"
            slide = {
                "title": title,
                "project": project_name,
                "sections": group_labels,
                "content": group_content,
            }
            slide["layout"] = (
                "single"
                if requires_single
                else ("compact" if compact_layout_allowed(slide, style) else "single")
            )
            plan.append(slide)
    return plan


def validate_slide_plan(plan: list[dict[str, object]], projects: list[dict[str, object]]) -> None:
    expected = [
        line for slide in plan if not slide["project"]
        for label, line in slide["content"]
    ]
    for project in projects:
        sections = project.get("sections", {})
        if isinstance(sections, dict):
            for label in SECTION_ORDER:
                expected.extend(section_lines(sections.get(label, "")))
    actual = [line for slide in plan for label, line in slide["content"]]
    if "".join(expected) != "".join(actual):
        raise ValueError("Slide plan omitted or duplicated approved scientific content.")


def plan_summary(plan: list[dict[str, object]]) -> list[str]:
    return [
        f"{index}. {slide['title']}"
        + (f" — {', '.join(slide['sections'])}" if slide["sections"] else "")
        for index, slide in enumerate(plan, start=1)
    ]


def compact_layout_allowed(slide_plan: dict[str, object], style: dict[str, object]) -> bool:
    """Allow two columns only for short, balanced content."""
    content = slide_plan.get("content", [])
    if len(content) < 2 or len(content) > 6:
        return False
    lines = [str(line) for _, line in content]
    if sum(len(line) for line in lines) > 420 or max(len(line) for line in lines) > 150:
        return False
    compact_chars_per_line = max(20, int(style["body_chars_per_line"]) // 2)
    groups = compact_column_groups(content)
    return all(
        sum(
            content_item_height(line, style, compact_chars_per_line)
            for _, line in group
        ) <= body_layout_capacity(style) + 0.01
        for group in groups
    )


def apply_pending_draft() -> None:
    pending_path = st.session_state.pop("pending_draft_path", None)
    if not pending_path:
        return
    try:
        payload = load_draft(Path(pending_path))
        st.session_state.projects = normalize_projects(payload["projects"])
        st.session_state.selected_project = 0
        st.session_state.report_title = str(payload.get("title", "Weekly Research Update"))
        st.session_state.report_author = str(payload.get("author", ""))
        st.session_state.report_date = date.fromisoformat(
            str(payload.get("report_date", date.today().isoformat()))
        )
        st.session_state.previous_next_steps = str(payload.get("previous", ""))
        st.session_state.presentation_mode = str(payload.get("presentation_mode", "Concise"))
        st.session_state.include_dividers = bool(payload.get("include_dividers", False))
        loaded_style = str(payload.get(
            "presentation_style",
            "Reference Style" if reference_style_available() else "Safe Default Style",
        ))
        st.session_state.presentation_style = (
            loaded_style if loaded_style in STYLE_NAMES
            else ("Reference Style" if reference_style_available() else "Safe Default Style")
        )
        st.session_state.editor_version = st.session_state.get("editor_version", 0) + 1
        st.session_state.draft_load_error = ""
    except (OSError, ValueError, TypeError, json.JSONDecodeError) as exc:
        st.session_state.draft_load_error = f"Could not load draft: {exc}"


def request_draft_load(path: str) -> None:
    st.session_state.pending_draft_path = path


apply_pending_draft()

if "report_date" not in st.session_state:
    st.session_state.report_date = date.today()
if "report_title" not in st.session_state:
    st.session_state.report_title = "Weekly Research Update"
if "report_author" not in st.session_state:
    st.session_state.report_author = "Huiqin (Hermine) Ni"
if "ai_result" not in st.session_state:
    st.session_state.ai_result = None
if "ai_error" not in st.session_state:
    st.session_state.ai_error = ""
if "presentation_mode" not in st.session_state:
    st.session_state.presentation_mode = "Concise"
if "include_dividers" not in st.session_state:
    st.session_state.include_dividers = False
if "presentation_style" not in st.session_state:
    st.session_state.presentation_style = (
        "Reference Style" if reference_style_available() else "Safe Default Style"
    )


def make_ppt(projects: list[dict[str, object]], title: str, report_date: date, author: str,
             previous: str, template_bytes: bytes | None, mode: str = "Concise",
             include_dividers: bool = False, style_name: str | None = None) -> bytes:
    style = resolve_style(style_name)
    if template_bytes:
        prs = Presentation(BytesIO(template_bytes))
        # Use the uploaded deck's slide dimensions and theme. To avoid exposing
        # previous weeks' private content, remove all old slides.
        slide_ids = prs.slides._sldIdLst
        for sid in list(slide_ids):
            prs.part.drop_rel(sid.rId)
            slide_ids.remove(sid)
        layout = next((l for l in prs.slide_layouts if l.name.lower().startswith("blank")), prs.slide_layouts[-1])
    else:
        prs = Presentation()
        prs.slide_width = Inches(SLIDE_WIDTH)
        prs.slide_height = Inches(SLIDE_HEIGHT)
        layout = prs.slide_layouts[6]

    width = prs.slide_width / Inches(1)
    height = prs.slide_height / Inches(1)
    scale_y = height / SLIDE_HEIGHT
    render_left = float(style["left_margin"])
    render_right = float(style["right_margin"])
    render_top = float(style["top_margin"]) * scale_y
    render_title_height = float(style["title_height"]) * scale_y
    render_footer_height = float(style["footer_height"]) * scale_y
    render_bottom_margin = float(style["bottom_margin"]) * scale_y
    render_footer_top = height - render_footer_height - render_bottom_margin
    render_body_top = float(style["body_top"]) * scale_y
    render_body_bottom = render_footer_top - float(style["footer_gap"]) * scale_y
    render_body_width = width - render_left - render_right

    def new_slide(heading: str, number: int):
        s = prs.slides.add_slide(layout)
        text_box(s, render_left, render_top, width - render_left - render_right,
                 render_title_height, heading, style["title_size"],
                 style["title_color"], True, style["title_font"])
        text_box(s, render_left, render_footer_top,
                 width - render_left - render_right, render_footer_height,
                 f"{author}   •   {report_date:%b %d, %Y}   •   {number:02d}",
                 style["footer_size"], style["footer_color"], False, style["footer_font"])
        return s

    s = prs.slides.add_slide(layout)
    text_box(s, 1.0, 2.0, width-2, .9, title, 38, style["title_color"], True, style["title_font"])
    text_box(s, 1.0, 3.0, width-2, .5, "Group Meeting", 22, style["heading_color"], False, style["heading_font"])
    text_box(s, 1.0, 3.6, width-2, .5, f"{author}  |  {report_date:%b %d, %Y}",
             18, style["footer_color"], False, style["footer_font"])

    plan = build_slide_plan(projects, previous, mode, include_dividers, style["name"])
    validate_slide_plan(plan, projects)
    idx = 2
    for slide_plan in plan[1:]:
        if not slide_plan["content"]:
            s = prs.slides.add_slide(layout)
            text_box(s, 1.0, 2.35, width - 2.0, 1.0, slide_plan["title"],
                     34, style["title_color"], True, style["title_font"])
            text_box(s, 1.0, 3.45, width - 2.0, .5, "Research Project",
                     20, style["heading_color"], False, style["heading_font"])
            text_box(s, render_left, render_footer_top,
                     width - render_left - render_right, render_footer_height,
                     f"{author}   •   {report_date:%b %d, %Y}   •   {idx:02d}",
                     style["footer_size"], style["footer_color"], False, style["footer_font"])
            idx += 1
            continue
        s = new_slide(str(slide_plan["title"]), idx)
        content = slide_plan["content"]
        compact = slide_plan.get("layout") == "compact"
        groups = compact_column_groups(content) if compact else [content]
        column_gap = 0.38 if compact else 0
        column_width = (render_body_width - column_gap) / 2 if compact else render_body_width
        for column_index, group in enumerate(groups):
            x = render_left + column_index * (column_width + column_gap)
            y = render_body_top
            for label, line in group:
                body_lines = estimate_lines(line, int(style["body_chars_per_line"] / (2 if compact else 1)))
                body_height = max(.52, body_lines * (float(style["body_size"]) / 72) * 1.25 + .08)
                heading_height = float(style["heading_size"]) / 72 * 1.25
                required_height = heading_height + body_height + .12
                if y + required_height > render_body_bottom + .01:
                    raise ValueError(f"Slide plan exceeds safe text area for '{slide_plan['title']}'.")
                text_box(s, x, y, column_width, heading_height, label,
                         style["heading_size"], style["heading_color"], True, style["heading_font"])
                text_box(s, x + .08, y + heading_height,
                         column_width - .08, body_height, "• " + line,
                         style["body_size"], style["body_color"], False, style["body_font"])
                y += required_height
        idx += 1
    output = BytesIO()
    prs.save(output)
    return output.getvalue()


st.title("Weekly Research Update Generator")
st.caption("Local-first • Editable slide text • Image layout stays under your control")
st.info("This version organizes notes by explicit headings. It does not use an AI model, translate Chinese, infer scientific conclusions, or send data to an external service.")

with st.sidebar:
    st.header("Report")
    report_date = st.date_input("Meeting date", key="report_date")
    title = st.text_input("Presentation title", key="report_title")
    author = st.text_input("Presented by", key="report_author")
    st.header("Research projects")
    if "projects" not in st.session_state:
        st.session_state.projects = [empty_project(1)]
    st.session_state.projects = normalize_projects(st.session_state.projects)
    if "selected_project" not in st.session_state:
        st.session_state.selected_project = 0

    if st.button("➕ Add research project", key="add_project"):
        st.session_state.projects.append(empty_project(len(st.session_state.projects) + 1))
        st.session_state.selected_project = len(st.session_state.projects) - 1
        st.rerun()

    for project_index, project in enumerate(st.session_state.projects):
        project_name = str(project["name"]).strip() or f"Research Project {project_index + 1}"
        project_key = str(project["id"])
        cols = st.columns([4, 1, 1, 1])
        if cols[0].button(f"✎ {project_name}", key=f"select_project_{project_key}"):
            st.session_state.selected_project = project_index
            st.rerun()
        if cols[1].button("↑ Move up", key=f"move_up_{project_key}", disabled=project_index == 0):
            st.session_state.selected_project = move_project(st.session_state.projects, project_index, -1)
            st.rerun()
        if cols[2].button("↓ Move down", key=f"move_down_{project_key}",
                          disabled=project_index == len(st.session_state.projects) - 1):
            st.session_state.selected_project = move_project(st.session_state.projects, project_index, 1)
            st.rerun()
        if cols[3].button("🗑 Delete", key=f"remove_project_{project_key}",
                          disabled=len(st.session_state.projects) == 1):
            st.session_state.pending_delete = project_key

    pending_delete = st.session_state.get("pending_delete")
    if pending_delete:
        pending_index = next((i for i, project in enumerate(st.session_state.projects)
                              if project["id"] == pending_delete), None)
        if pending_index is not None:
            st.warning(f"Delete '{st.session_state.projects[pending_index]['name']}'? This removes its edited text from the current draft.")
            confirm_cols = st.columns(2)
            if confirm_cols[0].button("Confirm delete", key="confirm_delete"):
                st.session_state.projects.pop(pending_index)
                st.session_state.selected_project = min(st.session_state.selected_project, len(st.session_state.projects) - 1)
                st.session_state.pending_delete = None
                st.rerun()
            if confirm_cols[1].button("Cancel delete", key="cancel_delete"):
                st.session_state.pending_delete = None
                st.rerun()

    selected_project = min(st.session_state.selected_project, len(st.session_state.projects) - 1)
    st.session_state.selected_project = selected_project
    selected_project_key = str(st.session_state.projects[selected_project]["id"])
    selected_name = st.text_input(
        "Selected project name",
        value=str(st.session_state.projects[selected_project]["name"]),
        key=f"project_name_{selected_project_key}",
    )
    st.session_state.projects[selected_project]["name"] = selected_name
    st.header("Optional files")
    uploaded_notes = st.file_uploader("Lab notes (.txt / .md)", type=["txt", "md"])
    template = st.file_uploader("Existing PPTX for slide theme (optional)", type=["pptx"])
    st.caption("Using a PPTX preserves its theme and dimensions, not its exact custom slide layouts. Existing slides are removed from the output.")

if "raw_notes" not in st.session_state:
    st.session_state.raw_notes = ""
notes = st.text_area(
    "Paste raw notes (headings such as Results:, Discussion:, Next Steps: are recognized)",
    height=230,
    placeholder="Objectives:\n...\nExperimental Work:\n...\nResults:\n...\nNext Steps:\n...",
    key="raw_notes",
)
if uploaded_notes:
    notes += "\n" + decode_upload(uploaded_notes)

if st.button("Organize notes into sections", type="primary"):
    st.session_state.projects[selected_project]["sections"] = split_notes(notes)
    st.session_state.editor_version = st.session_state.get("editor_version", 0) + 1

st.subheader("Faithful Scientific Editing (local draft)")
st.caption(
    "Uses only the local Ollama model qwen2.5:3b at 127.0.0.1. "
    "Output is unverified draft text and never replaces sections automatically."
)
with st.expander("Original notes supplied to the local editor", expanded=bool(notes.strip())):
    st.code(notes, language="text")
if st.button("Run Faithful Scientific Editing locally", key="run_local_editing"):
    try:
        st.session_state.ai_result = edit_notes(notes)
        st.session_state.ai_error = ""
    except ScientificWriterError as exc:
        st.session_state.ai_result = None
        st.session_state.ai_error = str(exc)

if st.session_state.ai_error:
    st.error(st.session_state.ai_error)
ai_result = st.session_state.ai_result
if ai_result:
    st.warning("AI-generated content is unverified draft text. Review it against the original notes before using it.")
    if ai_result.validation.has_warnings:
        st.error("Measurement validation warnings require human review:\n\n" + "\n".join(
            f"- {warning}" for warning in ai_result.validation.warnings
        ))
        if ai_result.validation.source_excerpts:
            st.warning("These source sections contain substantive text but the proposal left them empty. Review the original excerpts; nothing was filled automatically.")
            for label, excerpt in ai_result.validation.source_excerpts:
                st.markdown(f"**Original {label} excerpt**")
                st.code(excerpt, language="text")
    else:
        st.success("No numeric/unit differences were detected by the deterministic check.")
    proposed_sections = {}
    ai_columns = st.columns(2)
    for i, label in enumerate(WRITER_LABELS):
        with ai_columns[i % 2]:
            proposed_sections[label] = st.text_area(
                f"Proposed {label}",
                value=ai_result.sections[label],
                height=130,
                key=f"ai_proposed_{selected_project_key}_{label}",
            )
    if st.button("Copy reviewed proposals into project sections", key="apply_ai_proposals"):
        st.session_state.projects[selected_project]["sections"] = proposed_sections
        st.session_state.editor_version = st.session_state.get("editor_version", 0) + 1
        st.session_state.ai_result = None
        st.rerun()

st.subheader(f"Review and edit: {st.session_state.projects[selected_project]['name']}")
st.caption("Edit all sections yourself before exporting. No result or next step is invented.")
updated = {}
columns = st.columns(2)
for i, label in enumerate(LABELS):
    with columns[i%2]:
        value = st.session_state.projects[selected_project]["sections"].get(label, "")
        with st.expander(label, expanded=bool(value.strip())):
            updated[label] = st.text_area(
                f"Edit {label}",
                value=value,
                height=155,
                key=f"section_{selected_project_key}_{st.session_state.get('editor_version', 0)}_{i}",
            )
st.session_state.projects[selected_project]["sections"] = updated

with st.expander("Compare with last week (optional)"):
    previous = st.text_area("Paste previous week's Next Steps", height=120, key="previous_next_steps")
    st.caption("Previous steps are included as a reference slide, not marked complete automatically.")

st.subheader("Presentation composition")
presentation_mode = st.radio(
    "Slide detail",
    ("Concise", "Detailed"),
    horizontal=True,
    help="Concise combines short related sections. Detailed uses smaller content groups and more slides.",
    key="presentation_mode",
)
presentation_style = st.selectbox(
    "PowerPoint style",
    STYLE_NAMES,
    index=STYLE_NAMES.index(st.session_state.presentation_style)
    if st.session_state.presentation_style in STYLE_NAMES else 0,
    key="presentation_style",
    help="Reference Style uses verified metadata from the local reference deck. Safe Default Style uses the application's fallback typography.",
)
include_dividers = st.checkbox(
    "Include project divider slides",
    key="include_dividers",
    help="Concise mode omits dividers by default; enable this when separate project title slides are useful.",
)
try:
    current_slide_plan = build_slide_plan(
        st.session_state.projects, previous, presentation_mode, include_dividers,
        presentation_style,
    )
    validate_slide_plan(current_slide_plan, st.session_state.projects)
    with st.expander(f"Slide plan preview ({len(current_slide_plan)} slides)", expanded=True):
        for summary in plan_summary(current_slide_plan):
            st.write(summary)
except ValueError as exc:
    st.error(f"Could not build a complete slide plan: {exc}")
    current_slide_plan = []

with st.sidebar:
    st.header("Local drafts")
    draft_folder = st.text_input("Draft folder", ".local_drafts", key="draft_folder")
    draft_path = Path(draft_folder).expanduser()
    if st.button("💾 Save draft", key="save_draft"):
        try:
            saved = save_draft(draft_path, draft_payload(
                st.session_state.projects, title, author, report_date, previous,
                presentation_mode, include_dividers, presentation_style))
            st.success(f"Saved new draft: {saved.name}")
        except OSError as exc:
            st.error(f"Could not save draft: {exc}")

    draft_files = sorted(draft_path.glob("*.json")) if draft_path.is_dir() else []
    selected_draft = st.selectbox(
        "Draft to load",
        draft_files,
        format_func=lambda path: path.name,
        key="selected_draft",
    ) if draft_files else None
    st.button(
        "📂 Load draft",
        key="load_draft",
        disabled=selected_draft is None,
        on_click=request_draft_load,
        args=(str(selected_draft),) if selected_draft else (),
    )
    if st.session_state.get("draft_load_error"):
        st.error(st.session_state.pop("draft_load_error"))

if st.button("Generate editable PowerPoint"):
    if not any(any(str(value).strip() for value in project["sections"].values())
               for project in st.session_state.projects):
        st.warning("Add or organize some research notes first.")
    elif not current_slide_plan:
        st.error("Fix the slide plan before exporting.")
    else:
        try:
            data = make_ppt(
                st.session_state.projects, title, report_date, author, previous,
                template.getvalue() if template else None, presentation_mode, include_dividers,
                presentation_style,
            )
            st.success("PowerPoint generated. Add your own photos and finalize the layout in PowerPoint.")
            st.download_button("Download .pptx", data=data,
                               file_name=f"Weekly_Update_{report_date:%Y%m%d}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as exc:
            st.error(f"Could not generate PPTX: {exc}")

st.divider()
st.caption("Confidentiality: keep the repository private and run this app on your own computer. Streamlit uploads remain in this app's process memory; there are no external API calls in this code.")
