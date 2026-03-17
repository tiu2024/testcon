# Quiz Processor

Converts plain-text multiple-choice questions between formats: Word documents, hemis format, and back.

## Installation

Requires Python 3.10+.

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Generate all export formats from original quiz text
python main.py quiz.txt

# Generate only specific formats
python main.py --table --choices quiz.txt

# Convert hemis format back to original
python main.py --from-hemis quiz_hemis.txt

# Convert continuous-letter format back to original
python main.py --from-continuous quiz_continuous.txt
```

### Export flags (original → output)

| Flag          | Output                                                |
|---------------|-------------------------------------------------------|
| `--table`     | `.docx` — each question in a bordered table           |
| `--hemis`     | `.txt` — brace-wrapped, `====`/`+++++` separators     |
| `--choices`   | `.docx` — questions with all answer choices            |
| `--questions` | `.docx` — numbered questions only                      |

When no flags are given, all four are generated.

### Reverse conversion flags

| Flag                | Converts                            |
|---------------------|-------------------------------------|
| `--from-hemis`      | Hemis format → original text        |
| `--from-continuous` | Continuous-letter format → original |

All output goes to `quiz_output/`.

## Input Format (original)

```
1. What is the capital of France?
 a) London
 b) *Paris
 c) Berlin
 d) Madrid
```

- Questions may be numbered (`1. ...`) or unnumbered (auto-numbered on output).
- Options are lettered (`a)`, `b)`, …). Correct answer prefixed with `*`.
- Blank lines between questions are optional.

## Project Structure

```
cli.py           # Entry point and argument parsing
parser.py        # Parsers for original, hemis, and continuous formats
writers.py       # All output writers (Word and text)
models.py        # Question and Option dataclasses
requirements.txt
```

## License

Free to use and modify.