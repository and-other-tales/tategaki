#!/usr/bin/env python3
# ░▄▀▄░▀█▀░█▄█▒██▀▒█▀▄░▀█▀▒▄▀▄░█▒░▒██▀░▄▀▀
# ░▀▄▀░▒█▒▒█▒█░█▄▄░█▀▄░▒█▒░█▀█▒█▄▄░█▄▄▒▄██
"""
Genkō Yōshi Tategaki Converter - Convert Japanese text to proper genkō yōshi format
Implements authentic Japanese manuscript paper rules with grid-based layout
"""

import re
import argparse
import unicodedata
import sys
try:
    import chardet
except ImportError:
    chardet = None
from pathlib import Path
from docx import Document
from docx.shared import Mm, Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from rich.console import Console, Group
from rich.progress import Progress, BarColumn, TextColumn, TimeRemainingColumn, TaskProgressColumn
from rich.panel import Panel
from rich.live import Live


class GenkouYoshiGrid:
    """Manages the genkō yōshi grid layout and positioning"""
    
    def __init__(self, squares_per_column=12, max_columns_per_page=5):
        self.squares_per_column = squares_per_column
        self.max_columns_per_page = max_columns_per_page
        self.current_column = 1
        self.current_square = 1
        self.columns = {}  # Store content for each column/square
        self.pages = []  # Store completed pages
        self.current_page = 1
        
    def move_to_column(self, column_num, square_num=1):
        """Move to a specific column and square"""
        self.current_column = column_num
        self.current_square = square_num
        
    def advance_square(self):
        """Move to the next square in the current column"""
        self.current_square += 1
        if self.current_square > self.squares_per_column:
            self.current_column += 1
            self.current_square = 1
            
    def advance_column(self, square_num=1):
        """Move to the next column at specified square"""
        self.current_column += 1
        self.current_square = square_num
        
        # Check if we need a new page
        if self.current_column > self.max_columns_per_page:
            self.finish_page()
            self.current_column = 1
            self.current_square = square_num
        
    def skip_columns(self, num_columns):
        """Skip a number of columns (for spacing)"""
        self.current_column += num_columns
        self.current_square = 1
        
    def place_character(self, char, share_with_previous=False):
        """Place a character in the current square"""
        # Check if we need a new page
        if self.current_column > self.max_columns_per_page:
            self.finish_page()
            self.current_column = 1
            self.current_square = 1
            
        col_key = self.current_column
        square_key = self.current_square
        
        if col_key not in self.columns:
            self.columns[col_key] = {}
            
        if share_with_previous and square_key > 1:
            # Share square with previous character (for punctuation rules)
            prev_square = square_key - 1
            if prev_square in self.columns[col_key]:
                self.columns[col_key][prev_square] += char
            else:
                self.columns[col_key][square_key] = char
                self.advance_square()
        else:
            self.columns[col_key][square_key] = char
            self.advance_square()
            
    def finish_page(self):
        """Finish current page and start a new one"""
        if self.columns:
            self.pages.append({
                'page_num': self.current_page,
                'columns': self.columns.copy()
            })
            self.columns = {}
            self.current_page += 1
            
    def get_all_pages(self):
        """Get all pages including current page"""
        pages = self.pages.copy()
        if self.columns:  # Add current page if it has content
            pages.append({
                'page_num': self.current_page,
                'columns': self.columns.copy()
            })
        return pages
            
    def is_at_column_top(self):
        """Check if we're at the top of a column (square 1)"""
        return self.current_square == 1


class JapaneseTextProcessor:
    """Process Japanese text for genkō yōshi layout"""
    
    @staticmethod
    def identify_text_structure(text, paragraph_split_mode='blank'):
        """Identify novel title, subtitle, author, body text, and subheadings (chapters)
        paragraph_split_mode: 'blank' (default) or 'single' (split on every newline)
        """
        text = text.replace('\\r\\n', '\n').replace('\\n', '\n')
        text = text.replace('\r\n', '\n').replace('\r', '\n')
        lines = text.split('\n')
        structure = {
            'novel_title': None,
            'subtitle': None,
            'author': None,
            'body_paragraphs': [],
            'subheadings': []  # List of (chapter_title, [paragraphs])
        }
        # Detect explicit metadata markers for title, subtitle, author and remove those lines from content
        metadata_indices = set()
        title_pattern = re.compile(r'^(?:題名|タイトル)\s*[:：]\s*(.+)')
        subtitle_pattern = re.compile(r'^(?:副題|サブタイトル)\s*[:：]\s*(.+)')
        author_pattern = re.compile(r'^(?:作者|著者)\s*[:：]\s*(.+)')
        for idx, line in enumerate(lines):
            stripped = line.strip()
            if structure['novel_title'] is None:
                m = title_pattern.match(stripped)
                if m:
                    structure['novel_title'] = m.group(1).strip()
                    metadata_indices.add(idx)
                    continue
            if structure['subtitle'] is None:
                m = subtitle_pattern.match(stripped)
                if m:
                    structure['subtitle'] = m.group(1).strip()
                    metadata_indices.add(idx)
                    continue
            if structure['author'] is None:
                m = author_pattern.match(stripped)
                if m:
                    structure['author'] = m.group(1).strip()
                    metadata_indices.add(idx)
                    continue
        if metadata_indices:
            lines = [line for idx, line in enumerate(lines) if idx not in metadata_indices]

        # Identify non-empty lines for fallback metadata detection
        non_empty_lines = [line.strip() for line in lines if line.strip()]
        chapter_pattern = re.compile(r'^\s*第[0-9一二三四五六七八九十百千]+章.*')
        # Use a regex that treats lines with only whitespace (including full-width) as blank
        blankline_split = lambda txt: [p.strip() for p in re.split(r'(?:\n[\s\u3000]*\n)+', txt)]

        # Fallback for title, subtitle, author if not set by explicit markers
        pos = 0
        if structure['novel_title'] is None and non_empty_lines:
            structure['novel_title'] = non_empty_lines[0]
            pos = 1
        if structure['subtitle'] is None and pos < len(non_empty_lines) and not chapter_pattern.match(non_empty_lines[pos]):
            structure['subtitle'] = non_empty_lines[pos]
            pos += 1
        if structure['author'] is None and pos < len(non_empty_lines) and not chapter_pattern.match(non_empty_lines[pos]):
            structure['author'] = non_empty_lines[pos]
            pos += 1

        count = 0
        start_idx = 0
        for idx, line in enumerate(lines):
            if line.strip():
                count += 1
            if count == pos:
                start_idx = idx + 1
                break
        remaining_lines = lines[start_idx:]
        current_chapter = None
        buffer = []
        for line in remaining_lines:
            if chapter_pattern.match(line.strip()):
                if current_chapter is not None:
                    if paragraph_split_mode == 'single':
                        paragraphs = [p.strip() for p in buffer if p.strip()]
                    else:
                        paragraphs = blankline_split('\n'.join(buffer))
                    structure['subheadings'].append((current_chapter, paragraphs))
                    buffer = []
                current_chapter = line.strip()
            else:
                buffer.append(line)
        if current_chapter is not None:
            if paragraph_split_mode == 'single':
                paragraphs = [p.strip() for p in buffer if p.strip()]
            else:
                paragraphs = blankline_split('\n'.join(buffer))
            structure['subheadings'].append((current_chapter, paragraphs))
        else:
            if paragraph_split_mode == 'single':
                paragraphs = [p.strip() for p in remaining_lines if p.strip()]
            else:
                paragraphs = blankline_split('\n'.join(remaining_lines))
            structure['body_paragraphs'] = paragraphs
        return structure
        
    @staticmethod
    def get_vertical_char_equivalent(char):
        """Get the vertical equivalent of a character if it exists"""
        vertical_mappings = {
            '。': '︒',  # Vertical ideographic full stop
            '、': '︑',  # Vertical ideographic comma  
            '（': '︵',  # Vertical left parenthesis
            '）': '︶',  # Vertical right parenthesis
            '「': '﹁',  # Vertical left corner bracket
            '」': '﹂',  # Vertical right corner bracket
            '『': '﹃',  # Vertical left double angle bracket
            '』': '﹄',  # Vertical right double angle bracket
            '【': '︻',  # Vertical left tortoise shell bracket
            '】': '︼',  # Vertical right tortoise shell bracket
            '！': '︕',  # Vertical exclamation mark
            '？': '︖',  # Vertical question mark
            '：': '︓',  # Vertical colon
            '；': '︔',  # Vertical semicolon
        }
        return vertical_mappings.get(char, char)
        
    @staticmethod
    def is_punctuation(char):
        """Check if character is punctuation that follows special rules"""
        punctuation_chars = {'。', '、', '！', '？', '：', '；', '︒', '︑', '︕', '︖', '︓', '︔'}
        return char in punctuation_chars


class GenkouYoshiDocumentBuilder:
    """Build DOCX document following genkō yōshi rules"""
    
    def __init__(self, font_name='Noto Sans JP', squares_per_column=12, max_columns_per_page=5, chapter_pagebreak=False):
        self.doc = Document()
        self.grid = GenkouYoshiGrid(squares_per_column=squares_per_column, max_columns_per_page=max_columns_per_page)
        self.text_processor = JapaneseTextProcessor()
        self.font_name = font_name
        self.chapter_pagebreak = chapter_pagebreak
        self.setup_page_layout()
        
    @staticmethod
    def convert_to_fullwidth_number(num):
        """Convert Arabic numeral to full-width character"""
        fullwidth_map = {
            '0': '０', '1': '１', '2': '２', '3': '３', '4': '４',
            '5': '５', '6': '６', '7': '７', '8': '８', '9': '９'
        }
        return ''.join(fullwidth_map.get(char, char) for char in str(num))
    def _apply_number_rules(self, text: str) -> str:
        """Apply Japanese number writing rules: small numbers to kanji, larger to full-width digits,
        dates in 年月日 format, times with 時分 notation."""
        numbers_to_kanji = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五',
                          6: '六', 7: '七', 8: '八', 9: '九', 10: '十'}

        def repl_date(m):
            y, mo, da = m.group(1), m.group(2), m.group(3)
            y_str = self.convert_to_fullwidth_number(y)
            mo_n = int(mo)
            mo_str = numbers_to_kanji.get(mo_n, self.convert_to_fullwidth_number(mo))
            da_n = int(da)
            da_str = numbers_to_kanji.get(da_n, self.convert_to_fullwidth_number(da))
            return f"{y_str}年{mo_str}月{da_str}日"
        text = re.sub(r"(\d{1,4})年(\d{1,2})月(\d{1,2})日", repl_date, text)

        def repl_time(m):
            hh, mm = m.group(1), m.group(2)
            hh_n = int(hh)
            mm_n = int(mm)
            hh_str = numbers_to_kanji.get(hh_n, self.convert_to_fullwidth_number(hh))
            mm_str = numbers_to_kanji.get(mm_n, self.convert_to_fullwidth_number(mm))
            return f"{hh_str}時{mm_str}分"
        text = re.sub(r"(\d{1,2}):(\d{1,2})", repl_time, text)

        def repl_num(m):
            s = m.group(0)
            try:
                n = int(s)
            except ValueError:
                return s
            return numbers_to_kanji.get(n, self.convert_to_fullwidth_number(s))
        text = re.sub(r"\d+", repl_num, text)
        return text
        
    def setup_page_layout(self):
        """Setup genkō yōshi page layout for 178x111mm format"""
        section = self.doc.sections[0]
        
        # 178x111mm document dimensions (portrait)
        section.page_width = Mm(111)
        section.page_height = Mm(178)
        section.orientation = WD_ORIENTATION.PORTRAIT
        
        # Specified margins for genkou yoshi format
        section.top_margin = Mm(12.7)    # Top margin 12.7mm
        section.bottom_margin = Mm(12.7) # Bottom margin 12.7mm  
        section.left_margin = Mm(6.4)    # Left margin 6.4mm
        section.right_margin = Mm(6.4)   # Right margin 6.4mm
        
        # Set document vertical text direction
        self._set_document_vertical_text_direction(section)
        
    def create_genkou_yoshi_document(self, text):
        """Create document following genkō yōshi rules"""
        structure = self.text_processor.identify_text_structure(text)
        
        # Rule 1: Novel title on 1st column, first character in 4th square
        if structure['novel_title']:
            self.place_novel_title(structure['novel_title'])
            
        # Rule 2: Subtitle (chapter title) on 2nd column, 4th square
        if structure['subtitle']:
            self.place_subtitle(structure['subtitle'])
            
        # Rule 3: Author on 2nd column, 1st square (if present)
        if structure['author']:
            self.place_author(structure['author'])
            
        # Rule 4: First sentence begins on 3rd column, 2nd square
        self.grid.move_to_column(3, 2)
        
        paragraphs_to_process = structure['body_paragraphs']
        for i, paragraph in enumerate(paragraphs_to_process):
            if paragraph == '':
                # Scene break: leave one full empty column
                self.grid.advance_column(1)
                continue
            # Each new paragraph begins on 2nd square of new column
            if i > 0:
                self.grid.advance_column(2)
            self.place_paragraph(paragraph)
            
        # Finish any remaining page
        self.grid.finish_page()
        
        # Generate the actual DOCX content
        self.generate_docx_content()
        
    def place_novel_title(self, title):
        """Place novel title on 1st column, 4th square"""
        self.grid.move_to_column(1, 4)
        
        for char in title:
            vertical_char = self.text_processor.get_vertical_char_equivalent(char)
            self.grid.place_character(vertical_char)
            
    def place_subtitle(self, subtitle):
        """Place subtitle (chapter title) on 2nd column, 4th square"""
        self.grid.move_to_column(2, 4)
        
        for char in subtitle:
            vertical_char = self.text_processor.get_vertical_char_equivalent(char)
            self.grid.place_character(vertical_char)
            
    def place_author(self, author):
        """Place author on 2nd column with proper spacing"""
        self.grid.move_to_column(2, 1)
        
        # Split author into family and given name (assumes space separation)
        name_parts = author.split()
        
        for i, part in enumerate(name_parts):
            if i > 0:
                # 1 square between family and given name
                self.grid.advance_square()
                
            for char in part:
                vertical_char = self.text_processor.get_vertical_char_equivalent(char)
                self.grid.place_character(vertical_char)
                
        # 1 empty square below
        self.grid.advance_square()
        
    def place_subheading(self, subheading):
        """Place subheading with 1 empty column before and after, starting at 3rd square"""
        # 1 empty column before
        self.grid.advance_column(1)
        
        # Start at 3rd square of new column
        self.grid.move_to_column(self.grid.current_column, 3)
        
        for char in subheading:
            vertical_char = self.text_processor.get_vertical_char_equivalent(char)
            self.grid.place_character(vertical_char)
            
        # 1 empty column after
        self.grid.advance_column(1)
        
    def place_paragraph(self, paragraph, indent=True):
        """Place paragraph text following punctuation rules, with optional indentation"""
        paragraph = self._apply_number_rules(paragraph)
        is_dialogue = paragraph.lstrip().startswith('「')
        if is_dialogue:
            self.grid.advance_column(1)
            indent = False
        if indent and paragraph and not paragraph.startswith('　'):
            paragraph = '　' + paragraph
        
        # Break long paragraphs into sentences for better column flow
        sentences = re.split(r'([。！？])', paragraph)
        
        # Recombine sentences with their punctuation
        current_sentence = ""
        for i, part in enumerate(sentences):
            if part in ['。', '！', '？']:
                current_sentence += part
                # Process this sentence
                self.place_sentence(current_sentence.strip())
                current_sentence = ""
            else:
                current_sentence += part
                
        # Handle any remaining text
        if current_sentence.strip():
            self.place_sentence(current_sentence.strip())
        if paragraph.rstrip().endswith('」') and is_dialogue:
            # advance squares to end of column
            remaining = self.grid.squares_per_column - (self.grid.current_square - 1)
            for _ in range(remaining):
                self.grid.advance_square()
    
    def place_sentence(self, sentence):
        """Place a sentence with proper character handling"""
        opening_forbidden = {'「', '（', '［'}
        closing_forbidden = {'」', '）', '］'}
        for char in sentence:
            if self.grid.current_square > self.grid.squares_per_column:
                self.grid.advance_column(1)
            if char in opening_forbidden and self.grid.current_square == self.grid.squares_per_column:
                self.grid.advance_column(1)
            if char in closing_forbidden and self.grid.is_at_column_top() and self.grid.current_column > 1:
                self.grid.advance_column(1)
            vertical_char = self.text_processor.get_vertical_char_equivalent(char)
            if self.text_processor.is_punctuation(char):
                if self.grid.is_at_column_top() and self.grid.current_column > 1:
                    prev = self.grid.current_column - 1
                    if prev in self.grid.columns and self.grid.squares_per_column in self.grid.columns[prev]:
                        self.grid.columns[prev][self.grid.squares_per_column] += vertical_char
                    else:
                        self.grid.place_character(vertical_char)
                else:
                    self.grid.place_character(vertical_char)
            else:
                self.grid.place_character(vertical_char)
                
    def generate_docx_content(self, progress=None, task=None):
        """Generate the actual DOCX document from the grid"""
        pages = self.grid.get_all_pages()

        if not pages:
            return

        # Reverse page order for print compatibility
        pages.reverse()

        if progress and task is not None:
            total_pages = len(pages)
            progress.reset(task, total=total_pages, completed=0, description="Generating pages...")

        # Calculate page numbering (right-to-left, starting from main text)
        total_pages = len(pages)
        page_numbers = {}
        
        # Determine which pages contain main text (exclude title pages)
        main_text_start_page = self._find_main_text_start_page(pages)
        
        for i, page_data in enumerate(pages):
            # Only number pages starting from main text
            if page_data['page_num'] >= main_text_start_page:
                # Calculate page number starting from 1 for main text
                page_numbers[page_data['page_num']] = total_pages - i - (main_text_start_page - 1)
            else:
                # Title pages and other preliminary pages don't get numbered
                page_numbers[page_data['page_num']] = None
            
        for page_idx, page_data in enumerate(pages):
            if progress and task is not None:
                progress.update(task, advance=1)
            if page_idx > 0:
                # Add page break between pages
                self.doc.add_page_break()
                
            # Add page number following tategaki novel specifications
            display_page_num = page_numbers[page_data['page_num']]
            
            if display_page_num is not None:
                page_num_para = self.doc.add_paragraph()
                if display_page_num % 2 == 1:
                    page_num_para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                else:
                    page_num_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                page_num_run = page_num_para.add_run(str(display_page_num))
                page_num_run.font.name = self.font_name
                page_num_run.font.size = Pt(10)
                
            
            columns = page_data['columns']
            # Always use 5 columns for the table
            max_column = self.grid.max_columns_per_page
            table = self.doc.add_table(rows=self.grid.squares_per_column, cols=max_column)
            table.autofit = False
            
            # Set table style with transparent borders
            table.style = 'Table Grid'
            
            # Center the table
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            # Make borders transparent/white
            self._set_table_borders_white(table)
            
            # Configure table for genkō yōshi layout with 12.7mm cell size
            for row_idx in range(self.grid.squares_per_column):
                row = table.rows[row_idx]
                row.height = Mm(12.7)  # Cell size 12.7mm as specified
                
                for col_idx in range(max_column):
                    cell = row.cells[col_idx]
                    cell.width = Mm(12.7)  # Cell size 12.7mm as specified
                    
                    # Calculate actual column number (right-to-left)
                    actual_col = max_column - col_idx
                    actual_square = row_idx + 1
                    
                    # Get character for this position
                    if actual_col in columns and actual_square in columns[actual_col]:
                        char = columns[actual_col][actual_square]
                        
                        # Add character to cell
                        p = cell.paragraphs[0]
                        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        self._set_vertical_text_direction(p)
                        
                        run = p.add_run(char)
                        font = run.font
                        font.name = self.font_name
                        font.size = Pt(12)  # Appropriate font size for 12.7mm cells
        
    def setup_multi_column_layout(self):
        """Setup multi-column layout for proper tategaki display"""
        section = self.doc.sections[0]
        sectPr = section._sectPr
        
        # Create columns element
        cols = OxmlElement('w:cols')
        cols.set(qn('w:num'), '1')  # Single column for now, but with vertical text
        sectPr.append(cols)
        
    def _set_vertical_text_direction(self, paragraph):
        """Set vertical text direction for paragraph"""
        p_element = paragraph._element
        
        pPr = p_element.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            p_element.insert(0, pPr)
            
        textDirection = OxmlElement('w:textDirection')
        textDirection.set(qn('w:val'), 'tbRl')
        pPr.append(textDirection)
        
        return paragraph
        
    def _set_document_vertical_text_direction(self, section):
        """Set document-wide vertical text direction"""
        sectPr = section._sectPr
        
        textDirection = OxmlElement('w:textDirection')
        textDirection.set(qn('w:val'), 'tbRl')
        sectPr.append(textDirection)
        
        bidi = OxmlElement('w:bidi')
        sectPr.append(bidi)
        
        return section
        
    def _set_table_borders_white(self, table):
        """Set table borders to white/transparent"""
        # Access the table's XML element
        tbl = table._tbl
        
        # Create table properties
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        
        # Set table borders to white
        tblBorders = OxmlElement('w:tblBorders')
        
        # Define border properties (white color, single line)
        border_attrs = {
            qn('w:val'): 'single',
            qn('w:sz'): '2',  # Small border size
            qn('w:space'): '0',
            qn('w:color'): 'FFFFFF'  # White color
        }
        
        # Create all border elements
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            for attr, value in border_attrs.items():
                border.set(attr, value)
            tblBorders.append(border)
        
        tblPr.append(tblBorders)
    
    def _find_main_text_start_page(self, pages):
        """Find the page where main text content begins (after title/author pages)"""
        for page_data in pages:
            columns = page_data['columns']
            if not columns:
                continue
                
            # Check if this page contains substantial text content
            # (more than just title and author which typically occupy first 2 columns)
            if len(columns) > 2:
                # Check for actual paragraph content (column 3 and beyond)
                for col_num in sorted(columns.keys()):
                    if col_num >= 3:  # Main text starts from column 3
                        return page_data['page_num']
        
        # Fallback: if no clear main text found, start numbering from page 1
        return 1
        
    def save_document(self, output_path):
        """Save the document"""
        self.doc.save(output_path)


def read_japanese_text(file_path):
    """Read Japanese text from file with encoding auto-detection and normalization"""
    # Try chardet if available
    if chardet:
        with open(file_path, 'rb') as f:
            raw = f.read()
            result = chardet.detect(raw)
            encoding = result['encoding']
            try:
                text = raw.decode(encoding).strip()
            except Exception:
                text = raw.decode('utf-8', errors='replace').strip()
    else:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                text = f.read().strip()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='shift-jis') as f:
                text = f.read().strip()
    # Normalize text
    text = unicodedata.normalize('NFKC', text)
    return text


def main():
    parser = argparse.ArgumentParser(description='Convert Japanese text to genkō yōshi format. Requires Noto Sans JP font for best results.')
    parser.add_argument('input_file', help='Input Japanese text file (.txt)')
    parser.add_argument('-o', '--output', help='Output DOCX file')
    parser.add_argument('-a', '--author', help='Author name (if not in text)')
    parser.add_argument('--font', default='Noto Sans JP', help='Font for DOCX output (default: Noto Sans JP)')
    parser.add_argument('--columns', type=int, default=5, help='Number of columns per page (default: 5)')
    parser.add_argument('--squares', type=int, default=12, help='Number of squares per column (default: 12)')
    parser.add_argument('--paragraph-split', choices=['blank', 'single'], default='blank', help='Paragraph split mode: blank (default, split on blank lines) or single (split on every newline)')
    parser.add_argument('--chapter-pagebreak', action='store_true', help='Insert a page break before each chapter')
    args = parser.parse_args()
    input_path = Path(args.input_file)
    if not input_path.exists():
        print(f"Error: Input file '{input_path}' not found.", file=sys.stderr)
        sys.exit(1)
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_name(input_path.stem + '_genkou_yoshi.docx')
    try:
        text = read_japanese_text(input_path)
        if not text:
            print("Error: Input file is empty.", file=sys.stderr)
            sys.exit(1)
        processor = JapaneseTextProcessor()
        structure = processor.identify_text_structure(
            text,
            paragraph_split_mode=args.paragraph_split
        )
        console = Console(color_system="auto", force_terminal=True, force_interactive=True)
        console.print("[cyan]░▄▀▄░▀█▀░█▄█▒██▀▒█▀▄░▀█▀▒▄▀▄░█▒░▒██▀░▄▀▀[/cyan]")
        console.print("[cyan]░▀▄▀░▒█▒▒█▒█░█▄▄░█▀▄░▒█▒░█▀█▒█▄▄░█▄▄▒▄██[/cyan]")
        for line in __doc__.strip().splitlines():
            console.print(f"[bold yellow]{line}[/bold yellow]")
        console.print()
        console.print("[bold cyan]Text analysis:[/bold cyan]")
        console.print(f"  [bold]Novel title:[/bold] {structure['novel_title']}")
        console.print(f"  [bold]Subtitle (chapter):[/bold] {structure['subtitle']}")
        console.print(f"  [bold]Author:[/bold] {structure['author']}")
        if structure['subheadings']:
            total_paragraphs = sum(len(pars) for _, pars in structure['subheadings'])
            console.print(f"  [bold]Chapters detected:[/bold] {len(structure['subheadings'])}")
            for i, (chapter, paragraphs) in enumerate(structure['subheadings'], 1):
                display_chapter = re.sub(r'^第[^章]*章[:：]?\s*', '', chapter)
                console.print(
                    f"    [bold magenta]Chapter {i}:[/bold magenta] '{display_chapter}' ({len(paragraphs)} paragraphs)"
                )
            console.print(f"  [bold]Total paragraphs:[/bold] {total_paragraphs}")
            console.print(
                f"  [bold]Processing all {total_paragraphs} paragraphs for {args.columns}-column layout...[/bold]"
            )
        else:
            total_paragraphs = len(structure['body_paragraphs'])
            console.print(f"  [bold]Total paragraphs:[/bold] {total_paragraphs}")
            console.print(
                f"  [bold]Processing all {total_paragraphs} paragraphs for {args.columns}-column layout...[/bold]"
            )
        if structure['subheadings'] and total_paragraphs <= 1:
            console.print(
                "[bold yellow]Warning:[/bold yellow] "
                "Only one paragraph detected in chapters. Check your input or paragraph split mode."
            )
        if not structure['subheadings'] and len(structure['body_paragraphs']) <= 1:
            console.print(
                "[bold yellow]Warning:[/bold yellow] "
                "Only one paragraph detected in body. Check your input or paragraph split mode."
            )
        builder = GenkouYoshiDocumentBuilder(
            font_name=args.font,
            squares_per_column=args.squares,
            max_columns_per_page=args.columns,
            chapter_pagebreak=args.chapter_pagebreak
        )
        progress = Progress(
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            TimeRemainingColumn(),
            console=console,
            auto_refresh=False,
        )
        status_panel = Panel("", height=5, title="Status", expand=True)
        panel_group = Group(status_panel, progress)
        with Live(panel_group, refresh_per_second=4, auto_refresh=False,
                  console=console, transient=False) as live:
            task = progress.add_task("Processing paragraphs...", total=total_paragraphs)
            if structure['subheadings']:
                for idx, (chapter, paragraphs) in enumerate(structure['subheadings']):
                    if args.chapter_pagebreak and idx > 0:
                        builder.grid.finish_page()
                    status_panel.renderable = chapter
                    live.update(panel_group, refresh=True)
                    builder.place_subheading(chapter)
                    for i, paragraph in enumerate(paragraphs):
                        if paragraph == "":
                            builder.grid.advance_column(1)
                            progress.update(task, advance=1)
                            continue
                        if i > 0:
                            builder.grid.advance_column(2)
                        status_panel.renderable = paragraph
                        live.update(panel_group, refresh=True)
                        builder.place_paragraph(paragraph)
                        progress.update(task, advance=1)
            else:
                paragraphs = structure['body_paragraphs']
                for i, paragraph in enumerate(paragraphs):
                    if paragraph == "":
                        builder.grid.advance_column(1)
                        progress.update(task, advance=1)
                        continue
                    if i > 0:
                        builder.grid.advance_column(2)
                    status_panel.renderable = paragraph
                    live.update(panel_group, refresh=True)
                    builder.place_paragraph(paragraph)
                    progress.update(task, advance=1)
            builder.grid.finish_page()

        # Generate DOCX pages with progress bar
        pages = builder.grid.get_all_pages()
        with Progress(
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            TimeRemainingColumn(),
            console=console,
        ) as pg:
            page_task = pg.add_task("Generating pages...", total=None)
            builder.generate_docx_content(pg, page_task)
        # Save document
        builder.save_document(output_path)
        console.print()
        console.print(f"[bold green]DOCX file saved to:[/bold green] {output_path}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()