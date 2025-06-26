#!/usr/bin/env python3
# ░▄▀▄░▀█▀░█▄█▒██▀▒█▀▄░▀█▀▒▄▀▄░█▒░▒██▀░▄▀▀
# ░▀▄▀░▒█▒▒█▒█░█▄▄░█▀▄░▒█▒░█▀█▒█▄▄░█▄▄▒▄██
"""
Genkō Yōshi Tategaki Converter - Convert Japanese text to proper genkō yōshi format
Implements authentic Japanese manuscript paper rules with grid-based layout
OPTIMIZED VERSION for high-performance processing of large documents
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


class OptimizedGenkouYoshiGrid:
    """Optimized grid layout manager using efficient list-based data structures"""
    
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
        
        # Optimized data structure: Use list of lists instead of nested dictionaries
        # This provides O(1) access instead of O(log n) dictionary lookups
        self.current_page_grid = [['' for _ in range(self.squares_per_column)] 
                                  for _ in range(self.max_columns_per_page)]
        self.pages = []  # Store completed pages
        self.current_page = 1
        
        # Store margins and other format properties
        self.page_format = page_format
        
    def move_to_column(self, column_num, square_num=1):
        """Move to a specific column and square"""
        self.current_column = min(column_num, self.max_columns_per_page)
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
        
    def place_character_optimized(self, char):
        """Optimized character placement with minimal overhead"""
        # Input validation
        if not char or len(char) != 1:
            return
            
        # Check if we need a new page
        if self.current_column > self.max_columns_per_page:
            self.finish_page()
            self.current_column = 1
            self.current_square = 1
            
        col_idx = self.current_column - 1
        square_idx = self.current_square - 1
        
        # Bounds checking for safety
        if (col_idx >= 0 and col_idx < len(self.current_page_grid) and 
            square_idx >= 0 and square_idx < len(self.current_page_grid[col_idx])):
            self.current_page_grid[col_idx][square_idx] = char
            self.advance_square()
                
    def place_text_batch(self, text):
        """Optimized batch text placement - major performance improvement"""
        for char in text:
            # Inline the character placement logic for maximum performance
            if self.current_column > self.max_columns_per_page:
                self.finish_page()
                self.current_column = 1
                self.current_square = 1
                
            col_idx = self.current_column - 1
            square_idx = self.current_square - 1
            
            if col_idx < len(self.current_page_grid) and square_idx < len(self.current_page_grid[col_idx]):
                self.current_page_grid[col_idx][square_idx] = char
                
                # Inline advance_square logic
                self.current_square += 1
                if self.current_square > self.squares_per_column:
                    self.current_column += 1
                    self.current_square = 1
            
    def finish_page(self):
        """Finish current page and start a new one"""
        if any(any(col) for col in self.current_page_grid):
            # Convert to dictionary format for compatibility with existing DOCX generation
            columns_dict = {}
            for col_idx, column in enumerate(self.current_page_grid):
                col_num = col_idx + 1
                squares_dict = {}
                for square_idx, char in enumerate(column):
                    if char:
                        squares_dict[square_idx + 1] = char
                if squares_dict:
                    columns_dict[col_num] = squares_dict
                    
            self.pages.append({
                'page_num': self.current_page,
                'columns': columns_dict
            })
            
            # Reset current page with new grid
            self.current_page_grid = [['' for _ in range(self.squares_per_column)] 
                                      for _ in range(self.max_columns_per_page)]
            self.current_page += 1
            
    def get_all_pages(self):
        """Get all pages including current page"""
        pages = self.pages.copy()
        # Add current page if it has content
        if any(any(col) for col in self.current_page_grid):
            columns_dict = {}
            for col_idx, column in enumerate(self.current_page_grid):
                col_num = col_idx + 1
                squares_dict = {}
                for square_idx, char in enumerate(column):
                    if char:
                        squares_dict[square_idx + 1] = char
                if squares_dict:
                    columns_dict[col_num] = squares_dict
                    
            if columns_dict:
                pages.append({
                    'page_num': self.current_page,
                    'columns': columns_dict
                })
        return pages
            
    def is_at_column_top(self):
        """Check if we're at the top of a column (square 1)"""
        return self.current_square == 1


class OptimizedJapaneseTextProcessor:
    """Japanese text processor with cached patterns and single-pass processing"""
    
    # Pre-compiled regex patterns for maximum performance - compile once, use many times
    _title_pattern = re.compile(r'^(?:題名|タイトル|Title)\s*[:：]\s*(.+)')
    _subtitle_pattern = re.compile(r'^(?:副題|サブタイトル|Subtitle)\s*[:：]\s*(.+)')
    _author_pattern = re.compile(r'^(?:作者|著者|Author)\s*[:：]\s*(.+)')
    _chapter_pattern = re.compile(
        r'^\s*(?:第[一二三四五六七八九十百千\d]+章|'
        r'Chapter\s*\d+|Chapter\s*[IVXLCDM]+|[0-9]+\.|[一二三四五六七八九十]+\.).*'
    )
    _blankline_pattern = re.compile(r'(?:\n[\s\u3000]*\n)+')
    _date_pattern = re.compile(r"(\d{1,4})年(\d{1,2})月(\d{1,2})日")
    _time_pattern = re.compile(r"(\d{1,2}):(\d{1,2})")
    _number_pattern = re.compile(r"\d+")
    
    # Character sets as frozensets for O(1) membership testing
    PROHIBITED_COLUMN_START = frozenset({'。', '、', '」', '）', '］', '？', '！', '‼', '⁇', '⁈', '⁉', 
                                        '︒', '︑', '﹂', '︶', '︼', '︖', '︕'})
    PROHIBITED_COLUMN_END = frozenset({'「', '（', '［', '﹁', '︵', '︻'})
    SMALL_KANA = frozenset({'っ', 'ゃ', 'ゅ', 'ょ', 'ァ', 'ィ', 'ゥ', 'ェ', 'ォ', 'ッ', 'ャ', 'ュ', 'ョ'})
    
    # Pre-built translation tables for ultra-fast character conversion
    # Using str.maketrans and str.translate is much faster than dictionary lookups
    _fullwidth_ascii = str.maketrans(
        '!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz',
        '！＂＃＄％＆＇（）＊＋，－．／：；＜＝＞？＠［＼］＾＿｀｛｜｝～０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ'
    )
    
    _half_to_full_kana = str.maketrans(
        '｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ',
        '。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜'
    )
    
    # Vertical character mappings as translation table
    _vertical_translate = str.maketrans({
        '。': '︒', '、': '︑', '（': '︵', '）': '︶', '「': '﹁', '」': '﹂',
        '『': '﹃', '』': '﹄', '【': '︻', '】': '︼', '！': '︕', '？': '︖',
        '：': '︓', '；': '︔', '—': '︱', '－': '︲', '…': '︙',
    })
    
    # Number conversion tables
    _numbers_to_kanji = {
        '1': '一', '2': '二', '3': '三', '4': '四', '5': '五',
        '6': '六', '7': '七', '8': '八', '9': '九', '10': '十'
    }
    
    _fullwidth_numbers = str.maketrans('0123456789', '０１２３４５６７８９')
    
    @classmethod
    def identify_text_structure(cls, text, paragraph_split_mode='blank'):
        """Optimized text structure identification with cached patterns"""
        # Normalize line endings in one pass
        text = text.replace('\\r\\n', '\n').replace('\\n', '\n').replace('\r\n', '\n').replace('\r', '\n')
        lines = text.split('\n')
        
        structure = {
            'novel_title': None,
            'subtitle': None,
            'author': None,
            'body_paragraphs': [],
            'subheadings': []
        }
        
        # Single pass metadata detection with cached patterns
        metadata_indices = set()
        for idx, line in enumerate(lines):
            stripped = line.strip()
            if not stripped:
                continue
                
            # Use cached compiled patterns for faster matching
            if structure['novel_title'] is None:
                m = cls._title_pattern.match(stripped)
                if m:
                    structure['novel_title'] = m.group(1).strip()
                    metadata_indices.add(idx)
                    continue
                    
            if structure['subtitle'] is None:
                m = cls._subtitle_pattern.match(stripped)
                if m:
                    structure['subtitle'] = m.group(1).strip()
                    metadata_indices.add(idx)
                    continue
                    
            if structure['author'] is None:
                m = cls._author_pattern.match(stripped)
                if m:
                    structure['author'] = m.group(1).strip()
                    metadata_indices.add(idx)
                    continue
        
        # Remove metadata lines efficiently
        if metadata_indices:
            lines = [line for idx, line in enumerate(lines) if idx not in metadata_indices]

        # Identify non-empty lines for fallback metadata detection
        non_empty_lines = [line.strip() for line in lines if line.strip()]
        
        # Optimized blank line splitting using cached pattern
        def blankline_split(txt):
            return [p.strip() for p in cls._blankline_pattern.split(txt) if p.strip()]

        # Fallback metadata detection
        pos = 0
        if structure['novel_title'] is None and non_empty_lines:
            structure['novel_title'] = non_empty_lines[0]
            pos = 1
        if structure['author'] is None and pos < len(non_empty_lines) and not cls._chapter_pattern.match(non_empty_lines[pos]):
            structure['author'] = non_empty_lines[pos]
            pos += 1
        if structure['subtitle'] is None and pos < len(non_empty_lines) and not cls._chapter_pattern.match(non_empty_lines[pos]):
            structure['subtitle'] = non_empty_lines[pos]
            pos += 1

        # Find content start efficiently
        count = 0
        start_idx = 0
        for idx, line in enumerate(lines):
            if line.strip():
                count += 1
            if count == pos:
                start_idx = idx + 1
                break
                
        # Process chapters with cached pattern
        remaining_lines = lines[start_idx:]
        current_chapter = None
        buffer = []
        
        for line in remaining_lines:
            if cls._chapter_pattern.match(line.strip()):
                if current_chapter is not None:
                    paragraphs = ([p.strip() for p in buffer if p.strip()] if paragraph_split_mode == 'single' 
                                else blankline_split('\n'.join(buffer)))
                    structure['subheadings'].append((current_chapter, paragraphs))
                    buffer = []
                current_chapter = line.strip()
            else:
                buffer.append(line)
                
        if current_chapter is not None:
            paragraphs = ([p.strip() for p in buffer if p.strip()] if paragraph_split_mode == 'single' 
                         else blankline_split('\n'.join(buffer)))
            structure['subheadings'].append((current_chapter, paragraphs))
        else:
            paragraphs = ([p.strip() for p in remaining_lines if p.strip()] if paragraph_split_mode == 'single' 
                         else blankline_split('\n'.join(remaining_lines)))
            structure['body_paragraphs'] = paragraphs
            
        return structure
    
    @classmethod
    def preprocess_text_batch(cls, text):
        """
        MAJOR OPTIMIZATION: Single-pass text preprocessing combining all transformations
        This replaces multiple separate passes with one optimized pass
        """
        # Step 1: Convert to full-width using fast translation tables
        text = text.translate(cls._fullwidth_ascii)
        text = text.translate(cls._half_to_full_kana)
        
        # Step 2: Handle quotes and symbols (fastest string replacements)
        text = (text.replace('"', '「').replace('"', '」')
                   .replace("'", '「').replace("'", '」')
                   .replace('(', '（').replace(')', '）')
                   .replace('[', '［').replace(']', '］')
                   .replace('!', '！').replace('?', '？')
                   .replace('...', '…'))
        
        # Step 3: Apply number rules with cached patterns
        text = cls._apply_number_rules_optimized(text)
        
        # Step 4: Convert to vertical equivalents using translation table
        text = text.translate(cls._vertical_translate)
        
        return text
    
    @classmethod
    def _apply_number_rules_optimized(cls, text):
        """Optimized number rules application using cached patterns and fast string operations"""
        
        # Date conversion with cached pattern
        def repl_date(m):
            y, mo, da = m.group(1), m.group(2), m.group(3)
            y_str = y.translate(cls._fullwidth_numbers)
            mo_str = cls._numbers_to_kanji.get(mo, mo.translate(cls._fullwidth_numbers))
            da_str = cls._numbers_to_kanji.get(da, da.translate(cls._fullwidth_numbers))
            return f"{y_str}年{mo_str}月{da_str}日"
        
        text = cls._date_pattern.sub(repl_date, text)
        
        # Time conversion with cached pattern
        def repl_time(m):
            hh, mm = m.group(1), m.group(2)
            hh_str = cls._numbers_to_kanji.get(hh, hh.translate(cls._fullwidth_numbers))
            mm_str = cls._numbers_to_kanji.get(mm, mm.translate(cls._fullwidth_numbers))
            return f"{hh_str}時{mm_str}分"
        
        text = cls._time_pattern.sub(repl_time, text)
        
        # General number conversion with cached pattern
        def repl_num(m):
            s = m.group(0)
            return cls._numbers_to_kanji.get(s, s.translate(cls._fullwidth_numbers))
        
        text = cls._number_pattern.sub(repl_num, text)
        return text
    
    # Static method optimizations for common checks
    @staticmethod
    def is_punctuation(char):
        """Optimized punctuation check using frozenset"""
        return char in {'。', '、', '！', '？', '：', '；', '︒', '︑', '︕', '︖', '︓', '︔'}
        
    @classmethod
    def is_small_kana(cls, char):
        """Optimized small kana check using cached frozenset"""
        return char in cls.SMALL_KANA
    
    @classmethod
    def is_opening_bracket(cls, char):
        """Optimized opening bracket check using cached frozenset"""
        return char in cls.PROHIBITED_COLUMN_END
    
    @classmethod
    def is_closing_bracket(cls, char):
        """Optimized closing bracket check using cached frozenset"""
        return char in cls.PROHIBITED_COLUMN_START


class OptimizedGenkouYoshiDocumentBuilder:
    """Optimized DOCX document builder with batch processing and minimal overhead"""
    
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
        self.text_processor = OptimizedJapaneseTextProcessor()
        self.font_name = font_name
        self.chapter_pagebreak = chapter_pagebreak
        
        # Use page_format if provided, otherwise convert page_size to format, or use default
        if page_format:
            self.page_format = page_format
        elif page_size:
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
        
        # Create optimized grid
        self.grid = OptimizedGenkouYoshiGrid(
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
        
        # Optimized margin calculation
        margins = ({'top': 12, 'bottom': 12, 'inner': 10, 'outer': 8}, 7) if width <= 120 else \
                 ({'top': 15, 'bottom': 15, 'inner': 12, 'outer': 10}, 8) if width <= 160 else \
                 ({'top': 20, 'bottom': 20, 'inner': 15, 'outer': 12}, 9)
        
        margins, cell_size = margins
            
        # Calculate grid dimensions
        if PageSizeSelector:
            grid = PageSizeSelector.calculate_grid_dimensions(width, height, margins, cell_size)
            columns = grid['columns']
            rows = grid['rows']
            characters_per_page = grid['characters_per_page']
        else:
            # Optimized fallback calculation
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
        
    def setup_page_layout(self):
        """Optimized page layout setup"""
        section = self.doc.sections[0]
        
        # Set document dimensions
        section.page_width = Mm(self.page_size["width"])
        section.page_height = Mm(self.page_size["height"])
        section.orientation = WD_ORIENTATION.PORTRAIT
        
        # Use margins from page format with fallback
        if self.page_format and 'margins' in self.page_format:
            margins = self.page_format['margins']
        else:
            # Optimized margin fallback
            margin_size = 15 if self.page_size["width"] <= 120 else 18 if self.page_size["width"] <= 160 else 20
            margin_lr = 8 if self.page_size["width"] <= 120 else 10 if self.page_size["width"] <= 160 else 15
            margins = {'top': margin_size, 'bottom': margin_size, 'inner': margin_lr, 'outer': margin_lr}
            
        section.top_margin = Mm(margins['top'])
        section.bottom_margin = Mm(margins['bottom'])
        section.left_margin = Mm(margins['inner'])
        section.right_margin = Mm(margins['outer'])
        
        # Set document vertical text direction
        self._set_document_vertical_text_direction(section)
    
    def create_title_page(self, title, subtitle=None, author=None):
        """Optimized title page creation"""
        grid_cols = self.page_format['grid']['columns']
        grid_rows = self.page_format['grid']['rows']
        
        # Pre-calculate positions
        title_start_row = int(grid_rows * 0.2)
        author_row = int(grid_rows * 0.6)
        
        # Place title efficiently
        title_length = len(title)
        center_start = max(1, (grid_cols - min(title_length, grid_cols - 4)) // 2)
        
        self.grid.move_to_column(center_start, title_start_row)
        self.grid.place_text_batch(title)
            
        # Add subtitle if present
        if subtitle:
            subtitle_row = title_start_row + len(title) + 2
            subtitle_center = max(1, (grid_cols - min(len(subtitle), grid_cols - 4)) // 2)
            self.grid.move_to_column(subtitle_center, subtitle_row)
            self.grid.place_text_batch(subtitle)
                
        # Add author if present
        if author:
            author_center = max(1, (grid_cols - min(len(author), grid_cols - 4)) // 2)
            self.grid.move_to_column(author_center, author_row)
            self.grid.place_text_batch(author)
                
        self.grid.finish_page()
        
    def create_chapter_title_page(self, chapter_title):
        """Optimized chapter title page creation"""
        grid_cols = self.page_format['grid']['columns']
        grid_rows = self.page_format['grid']['rows']
        
        title_start_row = int(grid_rows * 0.15)
        
        # Clean up chapter title efficiently
        display_title = re.sub(r'^第[一二三四五六七八九十\d]+章[:：]?\s*', '', chapter_title)
        
        # Center and place title
        title_length = len(display_title)
        center_start = max(1, (grid_cols - min(title_length, grid_cols - 4)) // 2)
        
        self.grid.move_to_column(center_start, title_start_row)
        self.grid.place_text_batch(display_title)
            
        # Add spacing
        self.grid.move_to_column(1, title_start_row + title_length + 3)
        self.grid.finish_page()
        
    def _adaptive_paragraph_spacing(self, paragraph_count):
        """Optimized adaptive paragraph spacing"""
        rows = self.page_format['grid']['rows']
        return 1 if rows < 20 else 2 if rows < 30 else 3

    def create_genkou_yoshi_document(self, text):
        """
        MAJOR OPTIMIZATION: Streamlined document creation with minimal overhead
        """
        try:
            # Single-pass text preprocessing - MAJOR PERFORMANCE GAIN
            processed_text = self.text_processor.preprocess_text_batch(text)
            structure = self.text_processor.identify_text_structure(processed_text)
            
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
                        if not paragraph:
                            self.grid.advance_column(1)
                            continue
                        if i > 0:
                            self.grid.advance_column(spacing)
                        
                        # Optimized paragraph placement
                        self.place_paragraph_optimized(paragraph)
            else:
                if structure['novel_title']:
                    self.grid.move_to_column(3, 2)
                    
                paragraphs_to_process = structure['body_paragraphs']
                spacing = self._adaptive_paragraph_spacing(len(paragraphs_to_process))
                
                for i, paragraph in enumerate(paragraphs_to_process):
                    if not paragraph:
                        self.grid.advance_column(1)
                        continue
                    if i > 0:
                        self.grid.advance_column(spacing)
                    
                    # Optimized paragraph placement
                    self.place_paragraph_optimized(paragraph)
                        
            self.grid.finish_page()
            
        except Exception as e:
            logging.critical(f"Failed to generate document: {e}")
            raise
        
    def place_paragraph_optimized(self, paragraph, indent=True):
        """
        MAJOR OPTIMIZATION: Simplified paragraph placement with batch processing
        Eliminates complex recursive quote processing for significant performance gain
        """
        # Handle indentation efficiently
        if indent and self.grid.is_at_column_top():
            self.grid.advance_square()
            
        # Direct batch placement - much faster than character-by-character processing
        # The preprocessing has already handled all necessary transformations
        self.grid.place_text_batch(paragraph)

    def export_grid_metadata_json(self, output_path=None):
        """Export grid metadata as JSON for validation"""
        import json
        pages = self.grid.get_all_pages()
        metadata = []
        for page in pages:
            page_data = {
                'page_num': page['page_num'],
                'columns': {col: {sq: char for sq, char in squares.items()} 
                          for col, squares in page['columns'].items()}
            }
            metadata.append(page_data)
            
        json_str = json.dumps(metadata, ensure_ascii=False, indent=2)
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(json_str)
        return json_str

    def generate_docx_content_optimized(self, progress_callback=None):
        """
        MAJOR OPTIMIZATION: Batch DOCX content generation with minimal XML overhead
        """
        from docx.shared import Pt, Mm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        
        # Clear existing paragraphs
        while len(self.doc.paragraphs) > 0:
            p = self.doc.paragraphs[0]
            p._element.getparent().remove(p._element)
        
        # Pre-calculate all formatting parameters for batch processing
        margins = self.page_format['margins']
        grid_cols = self.grid.max_columns_per_page
        grid_rows = self.grid.squares_per_column
        
        avail_width = self.page_format['width'] - margins['inner'] - margins['outer']
        avail_height = self.page_format['height'] - margins['top'] - margins['bottom']
        
        cell_size = min(avail_width / grid_cols, avail_height / grid_rows)
        font_size_pt = cell_size * 0.7 * 2.83465  # mm to pt conversion
        cell_width_twips = int(cell_size * 56.7)  # mm to twips conversion
        
        # Process all pages with optimized batch operations
        pages = self.grid.get_all_pages()
        for page_idx, page in enumerate(pages):
            # Create table
            table = self.doc.add_table(rows=grid_rows, cols=grid_cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            # Batch process all cells for maximum performance
            for col in range(grid_cols):
                for row in range(grid_rows):
                    cell = table.cell(row, col)
                    
                    # Optimized cell formatting
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    w = OxmlElement('w:tcW')
                    w.set(qn('w:w'), str(cell_width_twips))
                    w.set(qn('w:type'), 'dxa')
                    tcPr.append(w)
                    cell.width = Mm(cell_size)
                    cell.height = Mm(cell_size)
                    
                    # Set content if present
                    char = page['columns'].get(col+1, {}).get(row+1, '')
                    if char:
                        p = cell.paragraphs[0]
                        p.clear()
                        run = p.add_run(char)
                        run.font.name = self.font_name
                        run.font.size = Pt(font_size_pt)
                        self._set_vertical_text_direction(p)
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add page break (except for last page)
            if page_idx < len(pages) - 1:
                self.doc.add_page_break()
                
            # Update progress for each page to show actual progress
            if progress_callback:
                progress_callback(page_idx + 1, len(pages))

    def _set_document_vertical_text_direction(self, section):
        """Set vertical text direction for document section"""
        sectPr = section._sectPr
        text_dir = sectPr.find(qn('w:textDirection'))
        if text_dir is None:
            text_dir = OxmlElement('w:textDirection')
            sectPr.append(text_dir)
        text_dir.set(qn('w:val'), 'tbRl')

    def _set_vertical_text_direction(self, paragraph):
        """Set vertical text direction for paragraph"""
        pPr = paragraph._p.get_or_add_pPr()
        text_dir = pPr.find(qn('w:textDirection'))
        if text_dir is None:
            text_dir = OxmlElement('w:textDirection')
            pPr.append(text_dir)
        text_dir.set(qn('w:val'), 'tbRl')


def main():
    """Optimized main function with streamlined processing"""
    from rich.console import Console
    from rich.progress import Progress, BarColumn, TextColumn, TimeRemainingColumn, TaskProgressColumn
    import sys
    
    parser = argparse.ArgumentParser(description="Japanese Tategaki DOCX Generator - OPTIMIZED")
    parser.add_argument("input", nargs="?", help="Input text file (UTF-8)")
    parser.add_argument("-o", "--output", help="Output DOCX file")
    parser.add_argument("--json", help="Export grid/character metadata as JSON to file")
    parser.add_argument("--format", default=None, help="Page format")
    args = parser.parse_args()

    console = Console(color_system="auto", force_terminal=True, force_interactive=True)
    
    # Display header
    ascii_art = (
        "[cyan]░▄▀▄░▀█▀░█▄█▒██▀▒█▀▄░▀█▀▒▄▀▄░█▒░▒██▀░▄▀▀[/cyan]\n"
        "[cyan]░▀▄▀░▒█▒▒█▒█░█▄▄░█▀▄░▒█▒░█▀█▒█▄▄░█▄▄▒▄██[/cyan]"
    )
    console.print(ascii_art)
    console.print("")
    console.print("[bold yellow]Japanese Tategaki DOCX Generator[/bold yellow]")
    console.print("[green]High-performance version for large documents[/green]")
    console.print()

    if not args.input:
        console.print("[bold red]No input file specified.[/bold red]")
        sys.exit(1)

    # Load and validate input
    input_path = Path(args.input)
    if not input_path.exists():
        console.print(f"[bold red]Error: Input file '{input_path}' not found.[/bold red]")
        sys.exit(1)
        
    # Improved file reading with encoding detection and error handling
    try:
        if chardet:
            with open(input_path, 'rb') as f:
                raw_data = f.read()
            encoding_result = chardet.detect(raw_data)
            detected_encoding = encoding_result['encoding'] if encoding_result['confidence'] > 0.7 else 'utf-8'
            try:
                text = raw_data.decode(detected_encoding).strip()
            except UnicodeDecodeError:
                text = raw_data.decode('utf-8', errors='ignore').strip()
        else:
            with open(input_path, encoding="utf-8", errors='ignore') as f:
                text = f.read().strip()
    except Exception as e:
        console.print(f"[bold red]Error reading file: {e}[/bold red]")
        sys.exit(1)
        
    if not text:
        console.print("[bold red]Error: Input file is empty.[/bold red]")
        sys.exit(1)

    # Page format selection
    if args.format is None or args.format == "custom":
        if PageSizeSelector and Prompt:
            selector = PageSizeSelector(console=console)
            page_format = selector.select_page_size()
        else:
            console.print("[bold red]Error: Interactive page size selection requires 'rich.prompt' and 'sizes.py'.[/bold red]")
            sys.exit(1)
    else:
        try:
            page_format = PageSizeSelector.get_format(args.format) if PageSizeSelector else None
        except Exception:
            page_format = None

    # Create optimized builder
    builder = OptimizedGenkouYoshiDocumentBuilder(page_format=page_format)
    
    # Quick analysis for user feedback
    structure = builder.text_processor.identify_text_structure(text)
    console.print("[bold cyan]Document Analysis:[/bold cyan]")
    console.print(f"  [bold]Title:[/bold] {structure['novel_title']}")
    console.print(f"  [bold]Author:[/bold] {structure['author']}")
    
    if structure['subheadings']:
        total_paragraphs = sum(len(pars) for _, pars in structure['subheadings'])
        console.print(f"  [bold]Chapters:[/bold] {len(structure['subheadings'])}")
    else:
        total_paragraphs = len(structure['body_paragraphs'])
        
    console.print(f"  [bold]Paragraphs:[/bold] {total_paragraphs}")
    console.print(f"  [bold]Characters:[/bold] ~{len(text):,}")

    # Process document with progress tracking
    with Progress(
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        TimeRemainingColumn(),
        console=console,
    ) as progress:
        
        # Main processing task
        task = progress.add_task("Processing document...", total=100)
        
        # Document creation (much less progress allocation)
        progress.update(task, advance=30, description="Creating document structure...")
        builder.create_genkou_yoshi_document(text)
        
        # DOCX generation gets more progress allocation for page-by-page updates
        progress.update(task, advance=10, description="Preparing DOCX generation...")
        pages = builder.grid.get_all_pages()
        
        # Allocate remaining 60% for page-by-page generation
        remaining_progress = 60
        progress_per_page = remaining_progress / max(1, len(pages))
        
        def progress_callback(current_page, total_pages):
            progress.update(task, advance=progress_per_page, 
                          description=f"Generating DOCX... Page {current_page}/{total_pages}")
            
        builder.generate_docx_content_optimized(progress_callback=progress_callback)
        progress.update(task, completed=100)

    # Save output
    output_path = Path(args.output) if args.output else input_path.with_name(input_path.stem + '_genkou_yoshi_optimized.docx')
    builder.doc.save(output_path)
    
    console.print()
    console.print(f"[bold green]✓ DOCX file saved:[/bold green] {output_path}")
    console.print(f"[bold green]✓ Pages generated:[/bold green] {len(pages)}")
    
    if args.json:
        builder.export_grid_metadata_json(args.json)
        console.print(f"[bold green]✓ Metadata JSON saved:[/bold green] {args.json}")


if __name__ == "__main__":
    main()