"""Parsers for reading quiz questions from various text formats.

Four input formats are supported:
    1. Original  — numbered/unnumbered questions with lettered options
    2. Hemis     — brace-wrapped, ====/+++++ separated
    3. Continuous — ++++ separated, options use continuously incrementing letters
    4. Table doc  — Word .docx with one bordered table per question
"""

import random
import re
from pathlib import Path

from docx import Document

from models import Option, Question

LETTERS = "abcdefghijklmnopqrstuvwxyz"

_OPTION_PATTERN = re.compile(r"^([a-z])\)\s+(\*?)(.+)$")
_NUMBERED_QUESTION_PATTERN = re.compile(r"^(\d+)\.\s+(.+)$")

_UNNUMBERED_PLACEHOLDER = 0


# ---------------------------------------------------------------------------
# Original format
# ---------------------------------------------------------------------------

def parse_original(filepath: Path) -> list[Question]:
    """Parse the standard quiz text format (numbered or unnumbered questions)."""
    content = filepath.read_text(encoding="utf-8")

    questions: list[Question] = []
    current: Question | None = None
    pending_text: str | None = None

    for raw_line in content.split("\n"):
        line = raw_line.strip()
        if not line:
            continue

        option_match = _OPTION_PATTERN.match(line)
        if option_match:
            if current is None:
                if pending_text:
                    current = Question(number=_UNNUMBERED_PLACEHOLDER, text=pending_text)
                    pending_text = None
                else:
                    continue

            current.options.append(
                Option(
                    letter=option_match.group(1),
                    text=option_match.group(3).strip(),
                    is_correct=option_match.group(2) == "*",
                )
            )
            continue

        question_match = _NUMBERED_QUESTION_PATTERN.match(line)
        if question_match:
            _save_if_complete(current, questions)
            current = Question(
                number=int(question_match.group(1)),
                text=question_match.group(2),
            )
            pending_text = None
            continue

        _save_if_complete(current, questions)
        current = None
        pending_text = line

    _save_if_complete(current, questions)
    _assign_sequential_numbers(questions)
    return questions


# ---------------------------------------------------------------------------
# Hemis format  (brace-wrapped, ==== / +++++ separated)
# ---------------------------------------------------------------------------

def parse_hemis(filepath: Path) -> list[Question]:
    """Parse the hemis format: { ... ==== ... +++++ ... }."""
    content = filepath.read_text(encoding="utf-8").strip()

    # Strip outer braces
    if content.startswith("{"):
        content = content[1:]
    if content.endswith("}"):
        content = content[:-1]
    content = content.strip().replace("\r\n", "\n")

    questions: list[Question] = []
    question_blocks = re.split(r"\n[ \t]*\+{4,5}[ \t]*\n", content)

    for block in question_blocks:
        block = block.strip()
        if not block:
            continue

        parts = [p.strip() for p in re.split(r"\n[ \t]*====[ \t]*\n", block) if p.strip()]
        if not parts:
            continue

        question_text = parts[0]
        if re.fullmatch(r"[+=]+", question_text) or question_text.startswith("#"):
            continue

        options = []
        for i, raw in enumerate(parts[1:]):
            if i >= len(LETTERS):
                break
            is_correct = raw.startswith("#")
            text = raw[1:].strip() if is_correct else raw.strip()
            options.append(Option(letter=LETTERS[i], text=text, is_correct=is_correct))

        questions.append(Question(
            number=len(questions) + 1,
            text=question_text,
            options=options,
        ))

    return questions


# ---------------------------------------------------------------------------
# Continuous format  (++++ separated, letters increment across questions)
# ---------------------------------------------------------------------------

def parse_continuous(filepath: Path) -> list[Question]:
    """Parse the continuous format where option letters span across questions."""
    content = filepath.read_text(encoding="utf-8")
    blocks = re.split(r"\s*\+{4}\s*", content)

    questions: list[Question] = []

    for block in blocks:
        block = block.strip()
        if not block:
            continue

        question_text = None
        raw_options: list[dict] = []

        for line in block.splitlines():
            line = line.strip()
            if not line:
                continue

            opt_match = _OPTION_PATTERN.match(line)
            if opt_match:
                raw_options.append({
                    "text": opt_match.group(3).strip(),
                    "is_correct": opt_match.group(2) == "*",
                })
                continue

            num_match = _NUMBERED_QUESTION_PATTERN.match(line)
            if num_match:
                question_text = num_match.group(1)
                continue

            if not raw_options:
                question_text = line

        if not question_text or not raw_options:
            continue

        options = [
            Option(letter=LETTERS[i], text=o["text"], is_correct=o["is_correct"])
            for i, o in enumerate(raw_options)
        ]

        questions.append(Question(
            number=len(questions) + 1,
            text=question_text,
            options=options,
        ))

    return questions


# ---------------------------------------------------------------------------
# Table doc format  (Word .docx, one single-column table per question)
# ---------------------------------------------------------------------------

def parse_table_doc(filepath: Path) -> list[Question]:
    """Parse a Word .docx produced by write_table_doc() back to questions.

    Each table has:
      row 0 → question text
      row 1 → correct answer text
      row 2+ → incorrect answer texts
    """
    doc = Document(filepath)
    questions: list[Question] = []

    for table in doc.tables:
        if len(table.rows) < 2:
            continue

        question_text = table.rows[0].cells[0].text.strip()
        correct_text = table.rows[1].cells[0].text.strip()
        if not question_text or not correct_text:
            continue

        texts = [correct_text] + [
            row.cells[0].text.strip()
            for row in table.rows[2:]
            if row.cells[0].text.strip()
        ]
        random.shuffle(texts)
        options = [
            Option(letter=LETTERS[i], text=text, is_correct=(text == correct_text))
            for i, text in enumerate(texts)
            if i < len(LETTERS)
        ]

        questions.append(Question(
            number=len(questions) + 1,
            text=question_text,
            options=options,
        ))

    return questions


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _save_if_complete(question: Question | None, dest: list[Question]) -> None:
    if question is not None and question.options:
        dest.append(question)


def _assign_sequential_numbers(questions: list[Question]) -> None:
    for index, question in enumerate(questions, start=1):
        if question.number == _UNNUMBERED_PLACEHOLDER:
            question.number = index