# Quiz Processor

A flexible Python tool to convert multiple choice questions from text format into multiple output formats including Word documents and custom text formats.

## Features

- Parses text files containing multiple choice questions
- Supports multiple output formats controlled by command-line flags
- All outputs are organized in a `quiz_output` folder
- Four available formats:
  1. **Table format (-f)** - Word document with each question in a table
  2. **Hash format (-n)** - Text file with custom bracket/separator format
  3. **With choices (-v)** - Word document with questions and all choices
  4. **Questions only (-nv)** - Word document with only questions

## Installation

1. Make sure you have Python 3.6 or higher installed
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Generate all formats (default when no flags are provided):

```bash
python quiz_processor.py input.txt
```

This generates all four formats in the `quiz_output/` folder:
- `input_table_format.docx` - Questions in table format
- `input_new_format.txt` - Questions in hash/bracket format
- `input_with_choices.docx` - Questions with all choices
- `input_questions.docx` - Questions only

### Using Flags

Generate specific formats only by using flags:

```bash
# Generate table format only
python quiz_processor.py -f input.txt

# Generate questions only
python quiz_processor.py -nv input.txt

# Generate multiple specific formats
python quiz_processor.py -f -v input.txt
python quiz_processor.py -nv -v -n input.txt

# Generate all formats explicitly
python quiz_processor.py -f -n -v -nv input.txt
```

### Available Flags

- `-f` / `--table-format`: Word document with table format (question in row 1, correct answer in row 2, other options below)
- `-n` / `--hash-format`: Text file with custom format using `{}`, `====`, and `+++++` separators
- `-v` / `--with-choices`: Word document with questions and all answer choices (no correct answer markers)
- `-nv` / `--no-choices`: Word document with questions only

## Input Format

The input text file can use either numbered or unnumbered questions:

### Format 1: Numbered Questions
```
1. Question 1
 a) Option 1
 b) *Option 2
 c) Option 3
 d) Option 4

2. Question 2
 a) Option 1
 b) Option 2
 c) Option 3
 d) *Option 4
```

### Format 2: Unnumbered Questions
```
Question 1
 a) Option 1
 b) *Option 2
 c) Option 3
 d) Option 4

Question 2
 a) Option 1
 b) Option 2
 c) Option 3
 d) *Option 4
```

**Format rules:**
- Questions can be numbered (e.g., `1.`, `2.`, etc.) or unnumbered
- Unnumbered questions will be automatically numbered in the output
- Options must be lettered (e.g., `a)`, `b)`, etc.)
- The correct answer is marked with an asterisk `*` before the option text
- Blank lines between questions are optional
- Both formats can be mixed in the same file

## Output Formats

### 1. Table Format (-f)

Word document where each question is in a single-column table:
- **Row 1:** Question text
- **Row 2:** Correct answer (no label)
- **Row 3+:** Other options in original order (no labels)

Example:
```
┌─────────────────────────────────┐
│ What is the capital of France?  │
├─────────────────────────────────┤
│ Paris                           │  ← correct answer
├─────────────────────────────────┤
│ London                          │  ← other options
├─────────────────────────────────┤
│ Berlin                          │
├─────────────────────────────────┤
│ Madrid                          │
└─────────────────────────────────┘
```

### 2. Hash Format (-n)

Text file with custom bracket and separator format:
```
{
Question text
====
Option 1
====
#Correct option
====
Option 3
====

+++++
Next question...
}
```

### 3. With Choices (-v)

Word document with questions followed by all answer choices (a, b, c, d) without asterisk markers.

### 4. Questions Only (-nv)

Word document containing only the numbered questions without any options.

## Examples

### Generate All Formats

```bash
python quiz_processor.py my_quiz.txt
```

Generates in `quiz_output/`:
- `my_quiz_table_format.docx`
- `my_quiz_new_format.txt`
- `my_quiz_with_choices.docx`
- `my_quiz_questions.docx`

### Generate Table Format Only

```bash
python quiz_processor.py -f exam.txt
```

Generates only: `quiz_output/exam_table_format.docx`

### Generate Study Materials

```bash
python quiz_processor.py -nv -v study.txt
```

Generates:
- `quiz_output/study_questions.docx` (for practice)
- `quiz_output/study_with_choices.docx` (for review)

## Requirements

- Python 3.6+
- python-docx 1.1.2

## Output Directory

All generated files are automatically saved in the `quiz_output/` folder in your current working directory. The folder is created automatically if it doesn't exist.

## License

This is a simple utility tool. Feel free to use and modify as needed.
