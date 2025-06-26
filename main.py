#!/usr/bin/env python3
# ░▄▀▄░▀█▀░█▄█▒██▀▒█▀▄░▀█▀▒▄▀▄░█▒░▒██▀░▄▀▀
# ░▀▄▀░▒█▒▒█▒█░█▄▄░█▀▄░▒█▒░█▀█▒█▄▄░█▄▄▒▄██
"""
Genkō Yōshi Tategaki Converter - Convert Japanese text to proper genkō yōshi format
Implements authentic Japanese manuscript paper rules with grid-based layout
"""

import re
import argparse
import logging
try:
    import chardet
except ImportError:
    print("Warning: chardet module not found, falling back to default encodings")
    chardet = None
from pathlib import Path
from docx import Document
from docx.shared import Mm, Pt
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
try:
    from rich.prompt import Prompt
except ImportError:
    Prompt = None

# Import the page size selector
try:
    from sizes import PageSizeSelector
except ImportError:
    PageSizeSelector = None

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')


class GenkouYoshiGrid:
    """Manages the genkō yōshi grid layout and positioning"""
    
    def __init__(self, squares_per_column=20, max_columns_per_page=10, page_format=None):
        # If page_format is provided, use its grid dimensions
        if page_format and 'grid' in page_format:
            self.squares_per_column = page_format['grid']['rows']
            self.max_columns_per_page = page_format['grid']['columns']
        else:
            self.squares_per_column = squares_per_column
            self.max_columns_per_page = max_columns_per_page
            
        self.current_column = 1
        self.current_square = 1
        self.columns = {}  # Store content for each column/square
        self.pages = []  # Store completed pages
        self.current_page = 1
        
        # Store margins and other format properties
        self.page_format = page_format
        
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
        # Use more comprehensive chapter pattern regex
        chapter_pattern = re.compile(
            r'^\s*(?:第[一二三四五六七八九十百千\d]+章|'
            r'Chapter\s*\d+|[0-9]+\.|[一二三四五六七八九十]+\.).*'
        )
        
        # Use a regex that treats lines with only whitespace (including full-width) as blank
        def blankline_split(txt):
            """Split text on blank lines including those with only whitespace"""
            return [p.strip() for p in re.split(r'(?:\n[\s\u3000]*\n)+', txt)]

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
    
    # Define prohibited characters for line breaking algorithm (禁則処理)
    PROHIBITED_COLUMN_START = {'。', '、', '」', '）', '］', '？', '！', '‼', '⁇', '⁈', '⁉', 
                              '︒', '︑', '﹂', '︶', '︼', '︖', '︕'}
    PROHIBITED_COLUMN_END = {'「', '（', '［', '﹁', '︵', '︻'}
    SMALL_KANA = {'っ', 'ゃ', 'ゅ', 'ょ', 'ァ', 'ィ', 'ゥ', 'ェ', 'ォ', 'ッ', 'ャ', 'ュ', 'ョ'}
        
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
            '—': '︱',  # Vertical em dash
            '－': '︲',  # Vertical en dash
            '…': '︙',  # Vertical ellipsis
        }
        return vertical_mappings.get(char, char)
    
    @staticmethod
    def is_punctuation(char):
        """Check if character is punctuation that follows special rules"""
        punctuation_chars = {'。', '、', '！', '？', '：', '；', '︒', '︑', '︕', '︖', '︓', '︔'}
        return char in punctuation_chars
        
    @staticmethod
    def is_small_kana(char):
        """Check if character is a small kana that needs special handling"""
        return char in JapaneseTextProcessor.SMALL_KANA
    
    @staticmethod
    def is_opening_bracket(char):
        """Check if character is an opening bracket/quote"""
        return char in JapaneseTextProcessor.PROHIBITED_COLUMN_END
    
    @staticmethod
    def is_closing_bracket(char):
        """Check if character is a closing bracket/quote"""
        return char in JapaneseTextProcessor.PROHIBITED_COLUMN_START
    
    @staticmethod
    def is_long_vowel_mark(char):
        """Check if character is a long vowel mark that needs vertical orientation"""
        return char == 'ー'
    
    @staticmethod
    def to_fullwidth(text):
        """Convert all half-width ASCII and katakana to full-width equivalents."""
        # Full-width ASCII range: FF01-FF5E
        def _fw(c):
            code = ord(c)
            if 0x21 <= code <= 0x7E:
                return chr(code + 0xFEE0)
            # Half-width katakana (U+FF61–U+FF9F)
            half_kana = '｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ'
            full_kana = '。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜'
            if c in half_kana:
                return full_kana[half_kana.index(c)]
            return c
        return ''.join(_fw(c) for c in text)
    
    @staticmethod
    def convert_symbols(text):
        text = JapaneseTextProcessor.to_fullwidth(text)
        # Handle paired quotes separately to avoid dictionary key conflicts
        text = text.replace('"', '「')  # Opening double quote
        text = text.replace('"', '」')  # Closing double quote
        text = text.replace("'", '「')  # Opening single quote
        text = text.replace("'", '」')  # Closing single quote
        
        # Other symbol replacements
        symbol_map = {
            '(': '（',
            ')': '）',
            '[': '［',
            ']': '］',
            '!': '！',
            '?': '？',
            '...': '…',
        }
        
        for western, japanese in symbol_map.items():
            text = text.replace(western, japanese)
            
        return text

    @staticmethod
    def tokenize_with_boundaries(text):
        """Tokenize Japanese text and return list of (surface, part-of-speech, is_proper_noun, is_compound) tuples."""
        try:
            from janome.tokenizer import Tokenizer
        except ImportError:
            raise ImportError("janome is required for tokenizer-based word boundary handling.")
        t = Tokenizer()
        tokens = []
        for token in t.tokenize(text):
            surface = token.surface
            pos = token.part_of_speech.split(',')[0]
            is_proper_noun = '固有名詞' in token.part_of_speech
            is_compound = pos in ['名詞', '動詞', '形容詞'] and len(surface) > 1
            tokens.append((surface, pos, is_proper_noun, is_compound))
        return tokens


class GenkouYoshiDocumentBuilder:
    """Build DOCX document following genkō yōshi rules"""
    
    # Default page format (Bunko format)
    DEFAULT_PAGE_FORMAT = {
        'name': 'Bunko',
        'width': 105, 
        'height': 148,
        'grid': {'columns': 17, 'rows': 24},
        'characters_per_page': 408,
        'margins': {'top': 15, 'bottom': 15, 'inner': 12, 'outer': 8},
        'description': 'Standard mass-market paperback fiction'
    }
    
    def __init__(self, font_name='Noto Sans JP', squares_per_column=None, max_columns_per_page=None, 
                 chapter_pagebreak=False, page_size=None, page_format=None):
        self.doc = Document()
        self.text_processor = JapaneseTextProcessor()
        self.font_name = font_name
        self.chapter_pagebreak = chapter_pagebreak
        
        # Use page_format if provided, otherwise convert page_size to format, or use default
        if page_format:
            self.page_format = page_format
        elif page_size:
            # Convert legacy page_size to new format with appropriate margins and grid
            self.page_format = self.convert_page_size_to_format(page_size)
        else:
            self.page_format = self.DEFAULT_PAGE_FORMAT
            
        # For backwards compatibility
        self.page_size = {
            'name': self.page_format['name'],
            'width': self.page_format['width'],
            'height': self.page_format['height'],
        }
        
        # Override grid dimensions if explicitly specified
        if squares_per_column:
            self.page_format['grid']['rows'] = squares_per_column
        if max_columns_per_page:
            self.page_format['grid']['columns'] = max_columns_per_page
        
        # Create grid with format
        self.grid = GenkouYoshiGrid(
            squares_per_column=self.page_format['grid']['rows'], 
            max_columns_per_page=self.page_format['grid']['columns'],
            page_format=self.page_format
        )
        
        self.setup_page_layout()
        
    def convert_page_size_to_format(self, page_size):
        """Convert legacy page size dict to full page format"""
        width = page_size.get('width', 105)
        height = page_size.get('height', 148)
        name = page_size.get('name', 'Custom')
        
        # Default margins based on size
        if width <= 120:  # Small formats like A6, Bunko
            margins = {'top': 12, 'bottom': 12, 'inner': 10, 'outer': 8}
            cell_size = 7
        elif width <= 160:  # Medium formats like B6
            margins = {'top': 15, 'bottom': 15, 'inner': 12, 'outer': 10}
            cell_size = 8
        else:  # Large formats like A5, B5, A4
            margins = {'top': 20, 'bottom': 20, 'inner': 15, 'outer': 12}
            cell_size = 9
            
        # Calculate grid dimensions
        if PageSizeSelector:
            # Use the algorithm from PageSizeSelector if available
            grid = PageSizeSelector.calculate_grid_dimensions(width, height, margins, cell_size)
            columns = grid['columns']
            rows = grid['rows']
            characters_per_page = grid['characters_per_page']
        else:
            # Simple fallback calculation
            text_width = width - margins['inner'] - margins['outer']
            text_height = height - margins['top'] - margins['bottom']
            columns = max(10, int(text_width / cell_size))
            rows = max(15, int(text_height / cell_size))
            characters_per_page = columns * rows
            
        return {
            'name': name,
            'width': width,
            'height': height,
            'grid': {'columns': columns, 'rows': rows},
            'characters_per_page': characters_per_page,
            'margins': margins,
            'description': f'Custom {width}×{height}mm format',
            'character_size': cell_size
        }
        
    @staticmethod
    def convert_to_fullwidth_number(num):
        """Convert Arabic numeral to full-width character"""
        fullwidth_map = {
            '0': '０', '1': '１', '2': '２', '3': '３', '4': '４',
            '5': '５', '6': '６', '7': '７', '8': '８', '9': '９'
        }
        return ''.join(fullwidth_map.get(char, char) for char in str(num))
        
    def _apply_number_rules(self, text: str) -> str:
        """Apply Japanese number writing rules: small numbers to kanji, 
        larger to full-width digits, dates in 年月日 format, 
        times with 時分 notation."""
        numbers_to_kanji = {
            1: '一', 2: '二', 3: '三', 4: '四', 5: '五',
            6: '六', 7: '七', 8: '八', 9: '九', 10: '十'
        }

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
        """Setup genkō yōshi page layout based on selected page size"""
        section = self.doc.sections[0]
        
        # Set document dimensions based on selected page size
        section.page_width = Mm(self.page_size["width"])
        section.page_height = Mm(self.page_size["height"])
        section.orientation = WD_ORIENTATION.PORTRAIT
        
        # Use margins from page format if available
        if self.page_format and 'margins' in self.page_format:
            margins = self.page_format['margins']
            section.top_margin = Mm(margins['top'])
            section.bottom_margin = Mm(margins['bottom'])
            section.left_margin = Mm(margins['inner'])
            section.right_margin = Mm(margins['outer'])
        else:
            # Fallback to basic margin calculation
            if self.page_size["width"] <= 120:  # Small formats
                margin_top_bottom = 15
                margin_left_right = 8
            elif self.page_size["width"] <= 160:  # Medium formats
                margin_top_bottom = 18
                margin_left_right = 10
            else:  # Large formats
                margin_top_bottom = 20
                margin_left_right = 15
                
            # Set margins
            section.top_margin = Mm(margin_top_bottom)
            section.bottom_margin = Mm(margin_top_bottom)
            section.left_margin = Mm(margin_left_right)
            section.right_margin = Mm(margin_left_right)
        
        # Set document vertical text direction
        self._set_document_vertical_text_direction(section)
    
    def create_title_page(self, title, subtitle=None, author=None):
        """Create a properly formatted title page following Japanese conventions"""
        grid_cols = self.page_format['grid']['columns']
        grid_rows = self.page_format['grid']['rows']
        
        # Calculate adaptive positioning based on grid
        title_start_row = int(grid_rows * 0.2)  # 20% down from top
        author_row = int(grid_rows * 0.6)       # 60% down from top
        
        # Center the title
        title_length = len(title)
        center_start = max(1, (grid_cols - min(title_length, grid_cols - 4)) // 2)
        
        # Place title with large spacing
        self.grid.move_to_column(center_start, title_start_row)
        
        for char in title:
            vertical_char = self.text_processor.get_vertical_char_equivalent(char)
            self.grid.place_character(vertical_char)
            
        # Add subtitle if present
        if subtitle:
            subtitle_row = title_start_row + len(title) + 2  # Add spacing
            subtitle_length = len(subtitle)
            subtitle_center = max(1, (grid_cols - min(subtitle_length, grid_cols - 4)) // 2)
            
            self.grid.move_to_column(subtitle_center, subtitle_row)
            
            for char in subtitle:
                vertical_char = self.text_processor.get_vertical_char_equivalent(char)
                self.grid.place_character(vertical_char)
                
        # Add author at the bottom area if present
        if author:
            author_length = len(author)
            author_center = max(1, (grid_cols - min(author_length, grid_cols - 4)) // 2)
            
            self.grid.move_to_column(author_center, author_row)
            
            for char in author:
                vertical_char = self.text_processor.get_vertical_char_equivalent(char)
                self.grid.place_character(vertical_char)
                
        # Finish the page
        self.grid.finish_page()
        
    def create_chapter_title_page(self, chapter_title):
        """Create a page with properly formatted chapter title"""
        grid_cols = self.page_format['grid']['columns']
        grid_rows = self.page_format['grid']['rows']
        
        # Calculate adaptive positioning - chapter title at 15% from top
        title_start_row = int(grid_rows * 0.15)
        
        # Clean up chapter title if needed (remove 第X章 prefix for display)
        display_title = re.sub(r'^第[一二三四五六七八九十\d]+章[:：]?\s*', '', chapter_title)
        
        # Center the title
        title_length = len(display_title)
        center_start = max(1, (grid_cols - min(title_length, grid_cols - 4)) // 2)
        
        self.grid.move_to_column(center_start, title_start_row)
        
        for char in display_title:
            vertical_char = self.text_processor.get_vertical_char_equivalent(char)
            self.grid.place_character(vertical_char)
            
        # Add empty space after chapter title
        spacing = 3  # 3 empty rows after chapter title
        self.grid.move_to_column(1, title_start_row + title_length + spacing)
        
        self.grid.finish_page()
        
    def create_part_page(self, part_title, part_number=None, series_title=None):
        """Create a dedicated part/series/volume page."""
        grid_cols = self.page_format['grid']['columns']
        grid_rows = self.page_format['grid']['rows']
        start_row = int(grid_rows * 0.25)
        if series_title:
            center = max(1, (grid_cols - min(len(series_title), grid_cols - 4)) // 2)
            self.grid.move_to_column(center, start_row)
            for char in series_title:
                self.grid.place_character(self.text_processor.get_vertical_char_equivalent(char))
            self.grid.advance_column(2)
        if part_number:
            part_str = f"第{part_number}部"
            center = max(1, (grid_cols - min(len(part_str), grid_cols - 4)) // 2)
            self.grid.move_to_column(center, start_row + 2)
            for char in part_str:
                self.grid.place_character(self.text_processor.get_vertical_char_equivalent(char))
            self.grid.advance_column(2)
        if part_title:
            center = max(1, (grid_cols - min(len(part_title), grid_cols - 4)) // 2)
            self.grid.move_to_column(center, start_row + 4)
            for char in part_title:
                self.grid.place_character(self.text_processor.get_vertical_char_equivalent(char))
        self.grid.finish_page()
        
    def _adaptive_paragraph_spacing(self, paragraph_count):
        """Determine adaptive paragraph spacing based on grid size/format and paragraph count."""
        # Example: For denser grids, less spacing; for larger, more. Tune as needed.
        rows = self.page_format['grid']['rows']
        if rows < 20:
            return 1  # Minimal spacing
        elif rows < 30:
            return 2
        else:
            return 3

    def create_genkou_yoshi_document(self, text):
        try:
            text = self.text_processor.convert_symbols(text)
            structure = self.text_processor.identify_text_structure(text)
            if structure['novel_title']:
                self.create_title_page(
                    structure['novel_title'],
                    subtitle=structure['subtitle'],
                    author=structure['author']
                )
            if structure['subheadings']:
                for chapter, paragraphs in structure['subheadings']:
                    if self.chapter_pagebreak:
                        self.grid.finish_page()
                    self.create_chapter_title_page(chapter)
                    spacing = self._adaptive_paragraph_spacing(len(paragraphs))
                    for i, paragraph in enumerate(paragraphs):
                        if paragraph == '':
                            self.grid.advance_column(1)
                            continue
                        if i > 0:
                            self.grid.advance_column(spacing)
                        try:
                            self.place_paragraph(paragraph)
                        except Exception as e:
                            logging.error(f"Error placing paragraph: {e}")
            else:
                if structure['novel_title']:
                    self.grid.move_to_column(3, 2)
                paragraphs_to_process = structure['body_paragraphs']
                spacing = self._adaptive_paragraph_spacing(len(paragraphs_to_process))
                for i, paragraph in enumerate(paragraphs_to_process):
                    if paragraph == '':
                        self.grid.advance_column(1)
                        continue
                    if i > 0:
                        self.grid.advance_column(spacing)
                    try:
                        self.place_paragraph(paragraph)
                    except Exception as e:
                        logging.error(f"Error placing paragraph: {e}")
            self.grid.finish_page()
            self._balance_columns_last_page()
            self.generate_docx_content()
        except Exception as e:
            logging.critical(f"Failed to generate document: {e}")
            raise
        
    def _split_nested_quotes(self, text, open_quote='「', close_quote='」', alt_open='『', alt_close='』'):
        """Recursively split text into segments by nested Japanese quotes/dialogue marks."""
        segments = []
        stack = []
        buf = ''
        i = 0
        while i < len(text):
            c = text[i]
            if c == open_quote or c == alt_open:
                if buf:
                    segments.append(('text', buf))
                    buf = ''
                stack.append(c)
                segments.append(('open', c))
            elif c == close_quote or c == alt_close:
                if buf:
                    segments.append(('text', buf))
                    buf = ''
                if stack:
                    stack.pop()
                segments.append(('close', c))
            else:
                buf += c
            i += 1
        if buf:
            segments.append(('text', buf))
        return segments

    def _place_recursive_quotes(self, segments, indent=True, level=0):
        """Recursively place segments with correct formatting for nested quotes/dialogue."""
        for typ, val in segments:
            if typ == 'text':
                # Place as normal paragraph or sentence
                if val.strip():
                    self.place_sentence(val.strip())
            elif typ == 'open':
                # For each new quote, advance column and indent less for deeper levels
                self.grid.advance_column(1)
                if level == 0 and indent:
                    self.grid.advance_square()  # Standard dialogue indent
                self.grid.place_character(val)
            elif typ == 'close':
                self.grid.place_character(val)
                # For dialogue: advance to end of column after closing quote at outermost level
                if level == 0:
                    remaining = self.grid.squares_per_column - (self.grid.current_square - 1)
                    for _ in range(remaining):
                        self.grid.advance_square()

    def place_paragraph(self, paragraph, indent=True):
        """Place paragraph text following punctuation rules, with optional indentation and recursive quote handling"""
        paragraph = self._apply_number_rules(paragraph)
        # Use recursive quote splitting
        segments = self._split_nested_quotes(paragraph)
        self._place_recursive_quotes(segments, indent=indent, level=0)
    
    def place_sentence(self, sentence, indent=True):
        """
        Place a single sentence or text segment into the grid.
        This is used by _place_recursive_quotes for non-nested text.
        """
        # Optionally indent the first line of the sentence
        if indent and self.grid.is_at_column_top():
            self.grid.advance_square()
        for char in sentence:
            vertical_char = self.text_processor.get_vertical_char_equivalent(char)
            self.grid.place_character(vertical_char)
            
    def _balance_columns_last_page(self):
        """Balance columns on the last page for optimal visual flow (e.g., avoid single short columns)."""
        if not self.grid.pages:
            return
        last_page = self.grid.pages[-1] if self.grid.pages else None
        if not last_page:
            return
        columns = last_page['columns']
        max_col = self.grid.max_columns_per_page
        # Find the last non-empty column
        last_filled = max([col for col in columns if columns[col]], default=1)
        # If only one or two columns are filled, redistribute content for balance
        if last_filled <= 2:
            all_chars = []
            for col in range(1, last_filled+1):
                for sq in range(1, self.grid.squares_per_column+1):
                    if sq in columns.get(col, {}):
                        all_chars.append(columns[col][sq])
            # Redistribute across more columns
            new_cols = {i+1: {} for i in range(max_col)}
            idx = 0
            for c in all_chars:
                col = (idx // self.grid.squares_per_column) + 1
                sq = (idx % self.grid.squares_per_column) + 1
                new_cols[col][sq] = c
                idx += 1
            last_page['columns'] = new_cols

    def export_grid_metadata_json(self, output_path=None):
        """Export grid/character metadata as JSON for QA/validation."""
        import json
        pages = self.grid.get_all_pages()
        metadata = []
        for page in pages:
            page_data = {
                'page_num': page['page_num'],
                'columns': {}
            }
            for col, squares in page['columns'].items():
                page_data['columns'][col] = {sq: char for sq, char in squares.items()}
            metadata.append(page_data)
        json_str = json.dumps(metadata, ensure_ascii=False, indent=2)
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(json_str)
        return json_str
    
    def validate_grid_state(self):
        """Validate grid/page state for consistency and report issues."""
        issues = []
        for page in self.grid.get_all_pages():
            cols = page['columns']
            for col in range(1, self.grid.max_columns_per_page + 1):
                if col not in cols:
                    issues.append(f"Page {page['page_num']} missing column {col}")
                else:
                    for sq in range(1, self.grid.squares_per_column + 1):
                        if sq not in cols[col]:
                            # Allow empty squares at end, but warn if in middle
                            if any(s > sq for s in cols[col]):
                                issues.append(f"Page {page['page_num']} column {col} missing square {sq}")
        if issues:
            import logging
            for issue in issues:
                logging.warning(issue)
        return issues

    def run_integrated_tests(self):
        """Run comprehensive integrated tests and print results."""
        import io
        import sys
        test_results = []
        # Test 1: Simple vertical writing
        try:
            self.grid = GenkouYoshiGrid(squares_per_column=10, max_columns_per_page=5)
            self.create_genkou_yoshi_document("テスト文章。改行テスト。")
            assert self.grid.pages, "No pages generated."
            test_results.append(("Simple vertical writing", True, ""))
        except Exception as e:
            test_results.append(("Simple vertical writing", False, str(e)))
        # Test 2: Nested dialogue/quotes
        try:
            self.grid = GenkouYoshiGrid(squares_per_column=10, max_columns_per_page=5)
            self.create_genkou_yoshi_document("「外側『内側』外側」")
            assert self.grid.pages, "No pages generated."
            test_results.append(("Nested dialogue/quotes", True, ""))
        except Exception as e:
            test_results.append(("Nested dialogue/quotes", False, str(e)))
        # Test 3: Word/compound boundary (janome)
        try:
            self.grid = GenkouYoshiGrid(squares_per_column=10, max_columns_per_page=5)
            self.create_genkou_yoshi_document("東京都に行った。")
            assert self.grid.pages, "No pages generated."
            test_results.append(("Word/compound boundary", True, ""))
        except Exception as e:
            test_results.append(("Word/compound boundary", False, str(e)))
        # Test 4: Adaptive paragraph spacing
        try:
            self.grid = GenkouYoshiGrid(squares_per_column=24, max_columns_per_page=17)
            self.create_genkou_yoshi_document("段落1。\n\n段落2。\n\n段落3。")
            assert self.grid.pages, "No pages generated."
            test_results.append(("Adaptive paragraph spacing", True, ""))
        except Exception as e:
            test_results.append(("Adaptive paragraph spacing", False, str(e)))
        # Test 5: Column balancing
        try:
            self.grid = GenkouYoshiGrid(squares_per_column=5, max_columns_per_page=2)
            self.create_genkou_yoshi_document("短文。")
            assert self.grid.pages, "No pages generated."
            test_results.append(("Column balancing", True, ""))
        except Exception as e:
            test_results.append(("Column balancing", False, str(e)))
        # Test 6: Error handling (malformed input)
        try:
            self.grid = GenkouYoshiGrid(squares_per_column=5, max_columns_per_page=2)
            self.create_genkou_yoshi_document(None)
            test_results.append(("Error handling (malformed input)", False, "No error on malformed input"))
        except Exception:
            test_results.append(("Error handling (malformed input)", True, ""))
        # Test 7: JSON metadata output
        try:
            self.grid = GenkouYoshiGrid(squares_per_column=5, max_columns_per_page=2)
            self.create_genkou_yoshi_document("JSONテスト。")
            json_str = self.export_grid_metadata_json()
            assert 'page_num' in json_str, "JSON output missing page_num."
            test_results.append(("JSON metadata output", True, ""))
        except Exception as e:
            test_results.append(("JSON metadata output", False, str(e)))
        # Test 8: Layout validation
        try:
            self.grid = GenkouYoshiGrid(squares_per_column=5, max_columns_per_page=2)
            self.create_genkou_yoshi_document("検証テスト。")
            issues = self.validate_grid_state()
            test_results.append(("Layout validation", True if not issues else False, ", ".join(issues)))
        except Exception as e:
            test_results.append(("Layout validation", False, str(e)))
        # Print results
        print("\nIntegrated Test Results:")
        for name, passed, msg in test_results:
            print(f"  {'[OK]' if passed else '[FAIL]'} {name}: {msg}")
        return test_results

    def generate_docx_content(self, progress_callback=None):
        """Render the current grid state into the DOCX document as vertical (tategaki) text. Optionally call progress_callback after each page."""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        # Remove all paragraphs except the first (python-docx always creates one)
        while len(self.doc.paragraphs) > 0:
            p = self.doc.paragraphs[0]
            p._element.getparent().remove(p._element)
        # For each page in the grid, create a table representing the grid
        for page in self.grid.get_all_pages():
            columns = self.grid.max_columns_per_page
            rows = self.grid.squares_per_column
            table = self.doc.add_table(rows=rows, cols=columns)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            # Set table style and cell size
            cell_size = self.page_format.get('character_size', 8)  # mm
            for col in range(columns):
                for row in range(rows):
                    cell = table.cell(row, col)
                    # Set cell width/height
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    w = OxmlElement('w:tcW')
                    w.set(qn('w:w'), str(int(cell_size * 56.7)))  # mm to twips
                    w.set(qn('w:type'), 'dxa')
                    tcPr.append(w)
                    cell.height = Mm(cell_size)
                    # Clear default paragraph
                    for p in cell.paragraphs:
                        p.clear()
            # Fill in characters
            for col in range(1, columns+1):
                for row in range(1, rows+1):
                    cell = table.cell(row-1, col-1)
                    char = page['columns'].get(col, {}).get(row, '')
                    if char:
                        p = cell.paragraphs[0]
                        run = p.add_run(char)
                        run.font.name = self.font_name
                        run.font.size = Pt(self.page_format.get('font_size', 12))
                        self._set_vertical_text_direction(p)
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Add a page break after each page except the last
            if page != self.grid.get_all_pages()[-1]:
                self.doc.add_page_break()
            if progress_callback:
                progress_callback()

    def _set_document_vertical_text_direction(self, section):
        """Set vertical text direction for the document section (Word)."""
        # Set section text direction to vertical (tategaki)
        sectPr = section._sectPr
        text_dir = sectPr.find(qn('w:textDirection'))
        if text_dir is None:
            text_dir = OxmlElement('w:textDirection')
            sectPr.append(text_dir)
        text_dir.set(qn('w:val'), 'tbRl')

    def _set_vertical_text_direction(self, paragraph):
        """Set vertical text direction for a paragraph (Word)."""
        pPr = paragraph._p.get_or_add_pPr()
        text_dir = pPr.find(qn('w:textDirection'))
        if text_dir is None:
            text_dir = OxmlElement('w:textDirection')
            pPr.append(text_dir)
        text_dir.set(qn('w:val'), 'tbRl')


def main():
    from rich.console import Console, Group
    from rich.progress import Progress, BarColumn, TextColumn, TimeRemainingColumn, TaskProgressColumn
    from rich.panel import Panel
    from rich.live import Live
    import re
    import sys
    parser = argparse.ArgumentParser(description="Japanese Tategaki DOCX Generator")
    parser.add_argument("input", nargs="?", help="Input text file (UTF-8)")
    parser.add_argument("-o", "--output", help="Output DOCX file")
    parser.add_argument("--test", action="store_true", help="Run integrated tests and exit")
    parser.add_argument("--json", help="Export grid/character metadata as JSON to file")
    parser.add_argument("--format", default="bunko", help="Page format (bunko, tankobon, shinsho, a5_standard, b6_standard, genkou_yoshi_20x20, custom)")
    args = parser.parse_args()

    console = Console(color_system="auto", force_terminal=True, force_interactive=True)
    ascii_art = (
        "[cyan]░▄▀▄░▀█▀░█▄█▒██▀▒█▀▄░▀█▀▒▄▀▄░█▒░▒██▀░▄▀▀[/cyan]\n"
        "[cyan]░▀▄▀░▒█▒▒█▒█░█▄▄░█▀▄░▒█▒░█▀█▒█▄▄░█▄▄▒▄██[/cyan]"
    )
    console.print(ascii_art)
    for line in __doc__.strip().splitlines():
        console.print(f"[bold yellow]{line}[/bold yellow]")
    console.print()

    if args.test:
        builder = GenkouYoshiDocumentBuilder()
        builder.run_integrated_tests()
        sys.exit(0)

    if not args.input:
        console.print("[bold red]No input file specified. Use --test to run tests.[/bold red]")
        sys.exit(1)

    input_path = Path(args.input)
    if not input_path.exists():
        console.print(f"[bold red]Error: Input file '{input_path}' not found.[/bold red]", style="red")
        sys.exit(1)
    with open(input_path, encoding="utf-8") as f:
        text = f.read().strip()
    if not text:
        console.print("[bold red]Error: Input file is empty.[/bold red]", style="red")
        sys.exit(1)

    # Select page format
    if args.format == "custom":
        page_format = None
    else:
        try:
            from sizes import PageSizeSelector
            page_format = PageSizeSelector.get_format(args.format)
        except Exception:
            page_format = None

    builder = GenkouYoshiDocumentBuilder(page_format=page_format)
    processor = builder.text_processor
    structure = processor.identify_text_structure(text)
    console.print("[bold cyan]Text analysis:[/bold cyan]")
    console.print(f"  [bold]Novel title:[/bold] {structure['novel_title']}")
    console.print(f"  [bold]Subtitle (chapter):[/bold] {structure['subtitle']}")
    console.print(f"  [bold]Author:[/bold] {structure['author']}")
    if structure['subheadings']:
        total_paragraphs = sum(len(pars) for _, pars in structure['subheadings'])
        console.print(f"  [bold]Chapters detected:[/bold] {len(structure['subheadings'])}")
        for i, (chapter, paragraphs) in enumerate(structure['subheadings'], 1):
            display_chapter = re.sub(r'^第[^章]*章[:：]?\\s*', '', chapter)
            console.print(f"    [bold magenta]Chapter {i}:[/bold magenta] '{display_chapter}' ({len(paragraphs)} paragraphs)")
        console.print(f"  [bold]Total paragraphs:[/bold] {total_paragraphs}")
        console.print(f"  [bold]Processing all {total_paragraphs} paragraphs...[/bold]")
    else:
        total_paragraphs = len(structure['body_paragraphs'])
        console.print(f"  [bold]Total paragraphs:[/bold] {total_paragraphs}")
        console.print(f"  [bold]Processing all {total_paragraphs} paragraphs...[/bold]")
    if structure['subheadings'] and total_paragraphs <= 1:
        console.print("[bold yellow]Warning:[/bold yellow] Only one paragraph detected in chapters. Check your input or paragraph split mode.")
    if not structure['subheadings'] and len(structure['body_paragraphs']) <= 1:
        console.print("[bold yellow]Warning:[/bold yellow] Only one paragraph detected in body. Check your input or paragraph split mode.")

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
    with Live(panel_group, refresh_per_second=4, auto_refresh=False, console=console, transient=False) as live:
        task = progress.add_task("Processing paragraphs...", total=total_paragraphs)
        if structure['subheadings']:
            for idx, (chapter, paragraphs) in enumerate(structure['subheadings']):
                status_panel.renderable = chapter
                live.update(panel_group, refresh=True)
                builder.create_chapter_title_page(chapter)
                for i, paragraph in enumerate(paragraphs):
                    if paragraph == "":
                        builder.grid.advance_column(1)
                        progress.update(task, advance=1)
                        continue
                    if i > 0:
                        builder.grid.advance_column(builder._adaptive_paragraph_spacing(len(paragraphs)))
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
                    builder.grid.advance_column(builder._adaptive_paragraph_spacing(len(paragraphs)))
                status_panel.renderable = paragraph
                live.update(panel_group, refresh=True)
                builder.place_paragraph(paragraph)
                progress.update(task, advance=1)
        builder.grid.finish_page()

    # Generate DOCX pages with real progress bar
    pages = builder.grid.get_all_pages()
    with Progress(
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        TimeRemainingColumn(),
        console=console,
    ) as pg:
        page_task = pg.add_task("Rendering DOCX pages...", total=len(pages))
        # Patch generate_docx_content to accept a progress callback
        def progress_callback():
            pg.update(page_task, advance=1)
        builder.generate_docx_content(progress_callback=progress_callback)
    output_path = Path(args.output) if args.output else input_path.with_name(input_path.stem + '_genkou_yoshi.docx')
    builder.doc.save(output_path)
    console.print()
    console.print(f"[bold green]DOCX file saved to:[/bold green] {output_path}")
    if args.json:
        builder.export_grid_metadata_json(args.json)
        console.print(f"[bold green]Grid metadata JSON saved to:[/bold green] {args.json}")


if __name__ == "__main__":
    main()