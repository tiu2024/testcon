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
python main.py input/quiz.txt

# Generate only specific formats
python main.py input/quiz.txt --table --choices

# Convert hemis format back to original
python main.py input/quiz_hemis.txt --from-hemis

# Convert continuous-letter format back to original
python main.py input/quiz_continuous.txt --from-continuous

# Convert table .docx back to original text
python main.py input/quiz_table.docx --from-table
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

| Flag                | Converts                                  |
|---------------------|-------------------------------------------|
| `--from-hemis`      | Hemis format -> original text             |
| `--from-continuous` | Continuous-letter format -> original text |
| `--from-table`      | Table `.docx` format -> original text     |

## Output

Each run creates a subfolder inside `output/` named after the input file. For example:

```
python main.py input/test.txt
# Writes to: output/test/
#   output/test/test_table.docx
#   output/test/test_hemis.txt
#   output/test/test_choices.docx
#   output/test/test_questions.docx
```

## Input Format (original)

```
1. What is the capital of France?
 a) London
 b) *Paris
 c) Berlin
 d) Madrid
```

- Questions may be numbered (`1. ...`) or unnumbered (auto-numbered on output).
- Options are lettered (`a)`, `b)`, ...). Correct answer prefixed with `*`.
- Blank lines between questions are optional.

## Project Structure

```
input/           # Source quiz files
output/          # Generated output (one subfolder per input file)
main.py          # Entry point and argument parsing
parser.py        # Parsers for original, hemis, continuous, and table formats
writers.py       # All output writers (Word and text)
models.py        # Question and Option dataclasses
requirements.txt
```

## License

Free to use and modify.
