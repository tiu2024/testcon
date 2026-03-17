#!/usr/bin/env python3
"""CLI entry point for the quiz processor."""

import argparse
import sys
from pathlib import Path

from parser import parse_original, parse_hemis, parse_continuous
from writers import (
    write_table_doc,
    write_hemis_text,
    write_choices_doc,
    write_questions_doc,
    write_original_text,
)

OUTPUT_DIR = Path("quiz_output")

# (writer_function, file_suffix)
_EXPORT_FORMATS = {
    "table":     (write_table_doc,     "_table.docx"),
    "hemis":     (write_hemis_text,    "_hemis.txt"),
    "choices":   (write_choices_doc,   "_choices.docx"),
    "questions": (write_questions_doc, "_questions.docx"),
}


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="quiz_processor",
        description="Convert multiple-choice questions between formats.",
    )
    p.add_argument("input_file", type=Path, help="Input text file")

    export = p.add_argument_group(
        "export formats",
        "Convert original format → output. If none specified, all are generated.",
    )
    export.add_argument("--table",     action="store_true", help="Word doc with bordered tables")
    export.add_argument("--hemis",     action="store_true", help="Text file in hemis format")
    export.add_argument("--choices",   action="store_true", help="Word doc with questions and choices")
    export.add_argument("--questions", action="store_true", help="Word doc with questions only")

    convert = p.add_argument_group(
        "reverse conversion",
        "Convert other formats back to the original text format.",
    )
    convert.add_argument("--from-hemis",      action="store_true", help="Hemis → original text")
    convert.add_argument("--from-continuous",  action="store_true", help="Continuous → original text")

    return p


def main(argv: list[str] | None = None) -> None:
    args = _build_parser().parse_args(argv)

    input_path: Path = args.input_file
    if not input_path.exists():
        print(f"Error: '{input_path}' not found.", file=sys.stderr)
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)
    base_name = input_path.stem

    # --- Reverse conversions (hemis/continuous → original) -----------------
    if args.from_hemis or args.from_continuous:
        parser_fn = parse_hemis if args.from_hemis else parse_continuous
        questions = parser_fn(input_path)

        if not questions:
            print("No questions found in the input file.", file=sys.stderr)
            sys.exit(1)

        out_path = OUTPUT_DIR / f"{base_name}_original.txt"
        write_original_text(questions, out_path)
        print(f"Parsed {len(questions)} question(s) → {out_path}")
        return

    # --- Forward conversion (original → export formats) --------------------
    questions = parse_original(input_path)
    if not questions:
        print("No questions found in the input file.", file=sys.stderr)
        sys.exit(1)

    print(f"Parsed {len(questions)} question(s) from '{input_path}'")

    selected = {name for name in _EXPORT_FORMATS if getattr(args, name)}
    if not selected:
        selected = set(_EXPORT_FORMATS)

    for name in sorted(selected):
        writer_fn, suffix = _EXPORT_FORMATS[name]
        out_path = OUTPUT_DIR / f"{base_name}{suffix}"
        writer_fn(questions, out_path)
        print(f"  Created: {out_path}")


if __name__ == "__main__":
    main()