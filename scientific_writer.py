"""Local-only faithful scientific editing through Ollama."""
from __future__ import annotations

from dataclasses import dataclass
import json
import re
from collections import Counter
from typing import Callable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

OLLAMA_URL = "http://127.0.0.1:11434/api/generate"
MODEL = "qwen2.5:3b"
LABELS = (
    "Objectives",
    "Experimental Work",
    "Results & Observations",
    "Discussion",
    "Challenges",
    "Next Steps",
)
RESPONSE_SCHEMA = {
    "type": "object",
    "properties": {label: {"type": "string"} for label in LABELS},
    "required": list(LABELS),
    "additionalProperties": False,
}


class ScientificWriterError(RuntimeError):
    """Raised when local inference cannot produce a usable response."""


class ScientificWriterResponseError(ScientificWriterError):
    """Raised when Ollama responds but its content is unusable."""


@dataclass(frozen=True)
class ValidationReport:
    source_measurements: tuple[str, ...]
    proposed_measurements: tuple[str, ...]
    warnings: tuple[str, ...]
    missing_source_sections: tuple[str, ...] = ()
    source_excerpts: tuple[tuple[str, str], ...] = ()

    @property
    def has_warnings(self) -> bool:
        return bool(self.warnings)


@dataclass(frozen=True)
class EditingResult:
    original_notes: str
    sections: dict[str, str]
    validation: ValidationReport


def build_prompt(notes: str) -> str:
    return f"""You are a scientific language editor.

Faithful Scientific Editing mode only.
Translate and condense only the supplied Chinese/English notes into concise scientific English.
Do not add experimental conditions, mechanisms, results, interpretations, or next steps.
Preserve every number, unit, time point, comparison, and uncertainty.
If a statement is ambiguous, preserve the ambiguity and do not guess.
Return valid JSON only with exactly these string keys:
{", ".join(LABELS)}
Use an empty string when a section is not present in the notes.

Supplied notes:
{notes}
"""


def parse_response(response_text: str) -> dict[str, str]:
    try:
        parsed = json.loads(response_text)
    except json.JSONDecodeError as exc:
        raise ScientificWriterResponseError(
            "Ollama returned invalid JSON scientific editing output."
        ) from exc
    if not isinstance(parsed, dict):
        raise ScientificWriterResponseError("Ollama response was not a JSON object.")
    missing = [label for label in LABELS if label not in parsed]
    if missing:
        raise ScientificWriterResponseError(
            "Ollama response is incomplete; missing fields: " + ", ".join(missing)
        )
    invalid = [label for label in LABELS if not isinstance(parsed[label], str)]
    if invalid:
        raise ScientificWriterResponseError(
            "Ollama response has non-string fields: " + ", ".join(invalid)
        )
    return {label: parsed[label].strip() for label in LABELS}


SOURCE_HEADINGS = (
    ("Objectives", r"^(?:objectives?|aims?|goals?|purpose|research questions?|目标|目的)\s*[:：]"),
    ("Experimental Work", r"^(?:experimental work|experiments?|methods?|procedures?|protocols?|实验|实验工作|方法)\s*[:：]"),
    ("Results & Observations", r"^(?:results?|observations?|findings?|measurements?|data|结果|观察|数据)\s*[:：]"),
    ("Discussion", r"^(?:discussion|analysis|interpretation|hypothesis|讨论|分析|解释)\s*[:：]"),
    ("Challenges", r"^(?:challenges?|issues?|problems?|limitations?|difficulties?|问题|挑战|困难)\s*[:：]"),
    ("Next Steps", r"^(?:next steps?|next week|future work|plans?|to do|下一步|后续工作|计划)\s*[:：]"),
)


def extract_source_sections(notes: str) -> dict[str, str]:
    """Extract explicitly headed source content without inventing unheaded sections."""
    result = {label: [] for label in LABELS}
    current = None
    for line in notes.splitlines():
        stripped = line.strip().lstrip("#").strip()
        matched = False
        for label, pattern in SOURCE_HEADINGS:
            match = re.match(pattern, stripped, flags=re.IGNORECASE)
            if match:
                current = label
                remainder = stripped[match.end():].strip()
                if remainder:
                    result[label].append(remainder)
                matched = True
                break
        if not matched and stripped and current:
            result[current].append(stripped)
    return {label: "\n".join(lines).strip() for label, lines in result.items()}


def _measurement_tokens(text: str) -> tuple[str, ...]:
    number = r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][-+]?\d+)?"
    unit = r"(?:°?\s*[CFK]|nm|μm|um|mm|cm|m|mg|g|kg|mL|L|μL|uL|h|min|s|%)"
    tokens = re.findall(rf"{number}\s*{unit}\b", text, flags=re.IGNORECASE)
    return tuple(re.sub(r"\s+", " ", token).strip() for token in tokens)


def _canonical_measurement(token: str) -> tuple[str, str]:
    match = re.match(
        r"^\s*([-+]?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][-+]?\d+)?)\s*(.*)\s*$",
        token,
    )
    if not match:
        return token.lower(), ""
    number = f"{float(match.group(1)):g}"
    unit = re.sub(r"\s+", "", match.group(2)).lower()
    unit = {"°c": "c", "c": "c", "°f": "f", "f": "f", "°k": "k", "k": "k",
            "um": "μm", "ul": "μl"}.get(unit, unit)
    return number, unit


def validate_measurements(source: str, proposed: str) -> ValidationReport:
    source_tokens = _measurement_tokens(source)
    proposed_tokens = _measurement_tokens(proposed)
    source_counts = Counter(_canonical_measurement(token) for token in source_tokens)
    proposed_counts = Counter(_canonical_measurement(token) for token in proposed_tokens)
    warnings = []
    source_by_unit = Counter(unit for _, unit in source_counts.elements())
    proposed_by_unit = Counter(unit for _, unit in proposed_counts.elements())
    for key, count in source_counts.items():
        if proposed_counts[key] < count:
            number, unit = key
            label = f"{number} {unit}"
            if proposed_by_unit[unit] >= source_by_unit[unit]:
                warnings.append(f"Altered measurement: {label}")
            elif count > 1:
                warnings.append(f"Duplicate-count difference: {label} ({count} source occurrences)")
            else:
                warnings.append(f"Missing measurement: {label}")
    for key, count in proposed_counts.items():
        if count > source_counts[key]:
            number, unit = key
            label = f"{number} {unit}"
            if source_counts[key] and count > source_counts[key]:
                warnings.append(f"Duplicate-count difference: {label} ({count} proposed occurrences)")
            else:
                warnings.append(f"Added measurement: {label}")
    return ValidationReport(source_tokens, proposed_tokens, tuple(sorted(set(warnings))))


def _request_ollama(notes: str, opener: Callable[..., object] = urlopen) -> str:
    body = json.dumps({
        "model": MODEL,
        "prompt": build_prompt(notes),
        "stream": False,
        "format": RESPONSE_SCHEMA,
        "options": {"temperature": 0, "num_ctx": 2048, "num_predict": 700},
    }).encode("utf-8")
    request = Request(OLLAMA_URL, data=body, headers={"Content-Type": "application/json"})
    try:
        with opener(request, timeout=120) as response:
            raw_payload = response.read().decode("utf-8")
    except (HTTPError, URLError, TimeoutError, OSError) as exc:
        raise ScientificWriterError(
            "Local Ollama is unavailable at http://127.0.0.1:11434. "
            "Start Ollama and verify qwen2.5:3b is available."
        ) from exc
    try:
        payload = json.loads(raw_payload)
    except json.JSONDecodeError as exc:
        raise ScientificWriterResponseError(
            "Ollama returned invalid JSON at the API response level."
        ) from exc
    if not isinstance(payload, dict) or not isinstance(payload.get("response"), str):
        raise ScientificWriterResponseError("Ollama returned an unexpected response shape.")
    if payload.get("done") is False:
        raise ScientificWriterResponseError("Ollama response was incomplete or truncated.")
    return payload["response"]


def edit_notes(notes: str, opener: Callable[..., object] = urlopen) -> EditingResult:
    if not notes.strip():
        raise ScientificWriterError("Supply notes before running Faithful Scientific Editing.")
    response = _request_ollama(notes, opener)
    sections = parse_response(response)
    proposed_text = "\n".join(value for value in sections.values() if value)
    validation = validate_measurements(notes, proposed_text)
    source_sections = extract_source_sections(notes)
    missing = tuple(label for label, value in source_sections.items()
                    if len(value.strip()) >= 3 and not sections.get(label, "").strip())
    excerpts = tuple((label, value[:240]) for label, value in source_sections.items()
                     if label in missing)
    if missing:
        validation = ValidationReport(
            validation.source_measurements,
            validation.proposed_measurements,
            validation.warnings + tuple(
                f"Source section has substantive content but proposal is empty: {label}"
                for label in missing
            ),
            missing,
            excerpts,
        )
    return EditingResult(notes, sections, validation)
