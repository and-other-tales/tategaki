#OtherTales Tategaki
##Genkō Yōshi Tategaki Converter

Convert Japanese plain text into a properly formatted vertical Genkō Yōshi (原稿用紙) DOCX, following authentic manuscript-paper rules (tategaki/縦書き).

## Features

- Vertical writing (tategaki, 縦書き) from right to left, columns progressing right to left.
- One character per square (masume/升目), including small kana, long vowels, repetition marks.
- Precise punctuation placement and 禁則処理 (forbidden line breaks) for quotes, brackets, commas, periods, etc.
- Dialogue formatting: each speaker’s line starts on a new column; quotes positioned at column top/bottom.
- Scene breaks: auto-detected blank lines insert centered `＊　＊　＊` markers with empty columns.
- Advanced number rules:
  - Numbers 1–10 as kanji (一〜十).
  - Larger numbers in full-width Arabic digits (１５, １００).
  - Dates in 年月日 format (e.g. ２０２３年十月五日).
  - Times in 時分 notation (e.g. 九時３０分).
- Customizable grid: columns per page and squares per column.
- Automatic page numbering with half-width Arabic numerals at top outer margin (odd pages right, even pages left).

## Installation

Requires Python 3.8+ and the Noto Sans JP font installed on your system (supports vertical text).

```bash
git clone <repository>
cd <repository>
pip install -r requirements.txt
pip install rich chardet
# Install Noto Sans JP font on your system (e.g., Ubuntu/Debian: sudo apt install fonts-noto-sans-jp; macOS Homebrew: brew tap homebrew/cask-fonts && brew install --cask font-noto-sans-jp)
```

## Usage

```bash
python main.py [OPTIONS] INPUT_FILE.txt
```

### Options

| Option               | Description                                                                 |
|----------------------:|:----------------------------------------------------------------------------|
| `-o`, `--output`      | Output DOCX file path (default: `<input>_genkou_yoshi.docx`)                 |
| `-a`, `--author`      | Author name (if not specified in text)                                       |
| `--font`              | Font name for DOCX output (default: `Noto Sans JP`)                           |
| `--columns`           | Number of columns per page (default: 5)                                      |
| `--squares`           | Number of squares per column (default: 12)                                   |
| `--paragraph-split`   | Paragraph split mode: `blank` (default) or `single`                          |
| `--chapter-pagebreak` | Insert a page break before each chapter                                      |

## Example

```bash
python main.py my_novel.txt -o novel_genkou.docx --columns 5 --squares 12
```
