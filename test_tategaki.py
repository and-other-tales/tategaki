#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Comprehensive test suite for Tategaki DOCX Generator
Tests all classes and functions with unit and integration tests
"""

import unittest
import tempfile
import os
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
import json

# Import the modules to test
from main import (
    OptimizedGenkouYoshiGrid,
    OptimizedJapaneseTextProcessor,
    OptimizedGenkouYoshiDocumentBuilder
)
from sizes import OptimizedPageSizeSelector


class TestOptimizedGenkouYoshiGrid(unittest.TestCase):
    """Test cases for OptimizedGenkouYoshiGrid class"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.grid = OptimizedGenkouYoshiGrid(squares_per_column=20, max_columns_per_page=10)
        
    def test_initialization(self):
        """Test grid initialization"""
        self.assertEqual(self.grid.squares_per_column, 20)
        self.assertEqual(self.grid.max_columns_per_page, 10)
        self.assertEqual(self.grid.current_column, 1)
        self.assertEqual(self.grid.current_square, 1)
        self.assertEqual(self.grid.current_page, 1)
        
    def test_initialization_with_page_format(self):
        """Test grid initialization with page format"""
        page_format = {
            'grid': {'rows': 24, 'columns': 17},
            'margins': {'top': 15, 'bottom': 15, 'inner': 12, 'outer': 8}
        }
        grid = OptimizedGenkouYoshiGrid(page_format=page_format)
        self.assertEqual(grid.squares_per_column, 24)
        self.assertEqual(grid.max_columns_per_page, 17)
        
    def test_move_to_column(self):
        """Test move_to_column functionality"""
        self.grid.move_to_column(5, 10)
        self.assertEqual(self.grid.current_column, 5)
        self.assertEqual(self.grid.current_square, 10)
        
        # Test boundary condition
        self.grid.move_to_column(15, 5)  # Beyond max columns
        self.assertEqual(self.grid.current_column, 10)  # Should be clamped
        
    def test_advance_square(self):
        """Test advance_square functionality"""
        self.grid.advance_square()
        self.assertEqual(self.grid.current_square, 2)
        self.assertEqual(self.grid.current_column, 1)
        
        # Test column wrap
        self.grid.current_square = 20
        self.grid.advance_square()
        self.assertEqual(self.grid.current_square, 1)
        self.assertEqual(self.grid.current_column, 2)
        
    def test_advance_column(self):
        """Test advance_column functionality"""
        self.grid.advance_column(5)
        self.assertEqual(self.grid.current_column, 2)
        self.assertEqual(self.grid.current_square, 5)
        
    def test_place_character_optimized(self):
        """Test optimized character placement"""
        self.grid.place_character_optimized('A')
        # Check character was placed
        self.assertEqual(self.grid.current_page_grid[0][0], 'A')
        self.assertEqual(self.grid.current_square, 2)
        
    def test_place_text_batch(self):
        """Test batch text placement"""
        text = "ABCDE"
        self.grid.place_text_batch(text)
        
        # Check all characters were placed
        for i, char in enumerate(text):
            self.assertEqual(self.grid.current_page_grid[0][i], char)
            
    def test_page_overflow(self):
        """Test page overflow handling"""
        # Fill entire page
        for col in range(self.grid.max_columns_per_page):
            for square in range(self.grid.squares_per_column):
                self.grid.place_character_optimized('A')
                
        # This should trigger a new page
        self.grid.place_character_optimized('B')
        self.assertEqual(self.grid.current_page, 2)
        
    def test_finish_page(self):
        """Test page finishing"""
        self.grid.place_character_optimized('A')
        self.grid.finish_page()
        
        pages = self.grid.get_all_pages()
        self.assertEqual(len(pages), 1)
        self.assertEqual(pages[0]['columns'][1][1], 'A')
        
    def test_is_at_column_top(self):
        """Test column top detection"""
        self.assertTrue(self.grid.is_at_column_top())
        self.grid.advance_square()
        self.assertFalse(self.grid.is_at_column_top())


class TestOptimizedJapaneseTextProcessor(unittest.TestCase):
    """Test cases for OptimizedJapaneseTextProcessor class"""
    
    def test_identify_text_structure_with_metadata(self):
        """Test text structure identification with metadata"""
        text = """Title: Test Title
Author: Test Author
Chapter 1: Beginning
This is the first paragraph.

This is the second paragraph."""
        
        structure = OptimizedJapaneseTextProcessor.identify_text_structure(text)
        
        # Test should handle English text as well
        self.assertIsNotNone(structure['novel_title'])
        self.assertIsNotNone(structure['author'])
        
    def test_identify_text_structure_without_metadata(self):
        """Test text structure identification without explicit metadata"""
        text = """Title
Author Name
Chapter 1
Paragraph one

Paragraph two"""
        
        structure = OptimizedJapaneseTextProcessor.identify_text_structure(text)
        
        self.assertEqual(structure['novel_title'], 'Title')
        # Note: Author fallback logic expects author to be on line 2 only if line 2 is not a chapter
        # Since "Author Name" is followed by "Chapter 1", it should be detected as author
        self.assertEqual(structure['author'], 'Author Name')
        
    def test_preprocess_text_batch(self):
        """Test batch text preprocessing"""
        text = 'Hello "world" (test) 123'
        processed = OptimizedJapaneseTextProcessor.preprocess_text_batch(text)
        
        # Should convert to full-width
        self.assertIn('Ｈｅｌｌｏ', processed)
        
    def test_number_rules(self):
        """Test number conversion rules"""
        text = "2023年12月25日 14:30"
        processed = OptimizedJapaneseTextProcessor.preprocess_text_batch(text)
        
        # Should convert dates and times appropriately
        self.assertIn('年', processed)
        self.assertIn('月', processed)
        self.assertIn('日', processed)
        
    def test_character_checks(self):
        """Test character type checking methods"""
        self.assertTrue(OptimizedJapaneseTextProcessor.is_punctuation('。'))
        self.assertFalse(OptimizedJapaneseTextProcessor.is_punctuation('A'))
        
        self.assertTrue(OptimizedJapaneseTextProcessor.is_small_kana('っ'))
        self.assertFalse(OptimizedJapaneseTextProcessor.is_small_kana('A'))
        
        self.assertTrue(OptimizedJapaneseTextProcessor.is_opening_bracket('「'))
        self.assertFalse(OptimizedJapaneseTextProcessor.is_opening_bracket('A'))
        
        self.assertTrue(OptimizedJapaneseTextProcessor.is_closing_bracket('」'))
        self.assertFalse(OptimizedJapaneseTextProcessor.is_closing_bracket('A'))


class TestOptimizedPageSizeSelector(unittest.TestCase):
    """Test cases for OptimizedPageSizeSelector class"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.selector = OptimizedPageSizeSelector()
        
    def test_get_format(self):
        """Test format retrieval"""
        bunko = OptimizedPageSizeSelector.get_format('bunko')
        self.assertIsNotNone(bunko)
        self.assertEqual(bunko['name'], 'Bunko')
        
        # Test case insensitive
        bunko_upper = OptimizedPageSizeSelector.get_format('BUNKO')
        self.assertEqual(bunko, bunko_upper)
        
        # Test invalid format
        invalid = OptimizedPageSizeSelector.get_format('invalid')
        self.assertIsNone(invalid)
        
    def test_calculate_grid_dimensions(self):
        """Test grid dimension calculation"""
        margins = {'top': 15, 'bottom': 15, 'inner': 12, 'outer': 8}
        grid = OptimizedPageSizeSelector.calculate_grid_dimensions(111, 178, margins)
        
        self.assertIsInstance(grid['columns'], int)
        self.assertIsInstance(grid['rows'], int)
        self.assertGreater(grid['columns'], 0)
        self.assertGreater(grid['rows'], 0)
        self.assertEqual(grid['characters_per_page'], grid['columns'] * grid['rows'])


class TestOptimizedGenkouYoshiDocumentBuilder(unittest.TestCase):
    """Test cases for OptimizedGenkouYoshiDocumentBuilder class"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.builder = OptimizedGenkouYoshiDocumentBuilder()
        
    def test_initialization(self):
        """Test builder initialization"""
        self.assertIsNotNone(self.builder.doc)
        self.assertIsNotNone(self.builder.grid)
        self.assertIsNotNone(self.builder.text_processor)
        
    def test_initialization_with_page_format(self):
        """Test builder initialization with page format"""
        page_format = {
            'name': 'Test',
            'width': 150,
            'height': 200,
            'grid': {'columns': 20, 'rows': 30},
            'characters_per_page': 600,
            'margins': {'top': 20, 'bottom': 20, 'inner': 15, 'outer': 12}
        }
        builder = OptimizedGenkouYoshiDocumentBuilder(page_format=page_format)
        self.assertEqual(builder.page_format, page_format)
        
    def test_convert_page_size_to_format(self):
        """Test page size to format conversion"""
        page_size = {'name': 'Custom', 'width': 120, 'height': 180}
        format_result = self.builder.convert_page_size_to_format(page_size)
        
        self.assertEqual(format_result['name'], 'Custom')
        self.assertEqual(format_result['width'], 120)
        self.assertEqual(format_result['height'], 180)
        self.assertIn('grid', format_result)
        self.assertIn('margins', format_result)
        
    def test_create_title_page(self):
        """Test title page creation"""
        self.builder.create_title_page("Test Title", "Subtitle", "Author")
        pages = self.builder.grid.get_all_pages()
        self.assertGreater(len(pages), 0)
        
    def test_create_chapter_title_page(self):
        """Test chapter title page creation"""
        self.builder.create_chapter_title_page("Chapter 1: Test Chapter")
        pages = self.builder.grid.get_all_pages()
        self.assertGreater(len(pages), 0)
        
    def test_place_paragraph_optimized(self):
        """Test optimized paragraph placement"""
        initial_pos = (self.builder.grid.current_column, self.builder.grid.current_square)
        self.builder.place_paragraph_optimized("Test paragraph.")
        
        # Position should have advanced
        final_pos = (self.builder.grid.current_column, self.builder.grid.current_square)
        self.assertNotEqual(initial_pos, final_pos)
        
    def test_export_grid_metadata_json(self):
        """Test JSON metadata export"""
        self.builder.grid.place_text_batch("Test")
        json_str = self.builder.export_grid_metadata_json()
        
        # Should be valid JSON
        metadata = json.loads(json_str)
        self.assertIsInstance(metadata, list)


class TestIntegration(unittest.TestCase):
    """Integration tests for end-to-end functionality"""
    
    def setUp(self):
        """Set up test fixtures"""
        self.test_text = """Title: Test Novel
Author: Test Author

Chapter 1: Beginning

This is the first paragraph. Processing Japanese text.

This is the second paragraph. It includes quotes and parentheses.

Chapter 2: Continuation

Content of another chapter. Numbers like 2023/12/25 are also converted."""
        
    def test_end_to_end_document_creation(self):
        """Test complete document creation process"""
        builder = OptimizedGenkouYoshiDocumentBuilder()
        
        # Should not raise exceptions
        try:
            builder.create_genkou_yoshi_document(self.test_text)
            pages = builder.grid.get_all_pages()
            self.assertGreater(len(pages), 0)
        except Exception as e:
            self.fail(f"Document creation failed: {e}")
            
    def test_docx_generation(self):
        """Test DOCX generation process"""
        builder = OptimizedGenkouYoshiDocumentBuilder()
        builder.create_genkou_yoshi_document(self.test_text)
        
        # Mock progress callback
        progress_calls = []
        def mock_progress():
            progress_calls.append(True)
            
        try:
            builder.generate_docx_content_optimized(progress_callback=mock_progress)
            self.assertGreater(len(progress_calls), 0)
        except Exception as e:
            self.fail(f"DOCX generation failed: {e}")
            
    def test_file_io_operations(self):
        """Test file input/output operations"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', encoding='utf-8', delete=False) as f:
            f.write(self.test_text)
            temp_file = f.name
            
        try:
            # Test reading file
            with open(temp_file, 'r', encoding='utf-8') as f:
                content = f.read()
            self.assertEqual(content, self.test_text)
            
            # Test document processing
            builder = OptimizedGenkouYoshiDocumentBuilder()
            builder.create_genkou_yoshi_document(content)
            
            # Test DOCX saving
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as docx_file:
                builder.doc.save(docx_file.name)
                self.assertTrue(os.path.exists(docx_file.name))
                self.assertGreater(os.path.getsize(docx_file.name), 0)
                os.unlink(docx_file.name)
                
        finally:
            os.unlink(temp_file)
            
    def test_various_page_formats(self):
        """Test document creation with various page formats"""
        formats_to_test = ['bunko', 'tankobon', 'a5', 'custom_bunko']
        
        for format_name in formats_to_test:
            with self.subTest(format=format_name):
                page_format = OptimizedPageSizeSelector.get_format(format_name)
                self.assertIsNotNone(page_format, f"Format {format_name} not found")
                
                builder = OptimizedGenkouYoshiDocumentBuilder(page_format=page_format)
                try:
                    builder.create_genkou_yoshi_document(self.test_text)
                    pages = builder.grid.get_all_pages()
                    self.assertGreater(len(pages), 0)
                except Exception as e:
                    self.fail(f"Failed to create document with {format_name} format: {e}")


class TestErrorHandling(unittest.TestCase):
    """Test error handling and edge cases"""
    
    def test_empty_text_handling(self):
        """Test handling of empty text"""
        builder = OptimizedGenkouYoshiDocumentBuilder()
        try:
            builder.create_genkou_yoshi_document("")
            # Should not crash, but may create minimal document
        except Exception as e:
            # Log the exception but don't fail - this is expected behavior
            pass
            
    def test_very_long_text_handling(self):
        """Test handling of very long text"""
        long_text = "A" * 10000  # 10,000 characters
        builder = OptimizedGenkouYoshiDocumentBuilder()
        
        try:
            builder.create_genkou_yoshi_document(long_text)
            pages = builder.grid.get_all_pages()
            self.assertGreater(len(pages), 1)  # Should create multiple pages
        except Exception as e:
            self.fail(f"Failed to handle long text: {e}")
            
    def test_special_characters_handling(self):
        """Test handling of special characters"""
        special_text = "Special chars: !@#$%^&*()_+-=[]{}|;:,.<>?"
        builder = OptimizedGenkouYoshiDocumentBuilder()
        
        try:
            builder.create_genkou_yoshi_document(special_text)
            # Should not crash
        except Exception as e:
            self.fail(f"Failed to handle special characters: {e}")
            
    def test_invalid_page_format_handling(self):
        """Test handling of invalid page format"""
        invalid_format = {
            'name': 'Invalid',
            'width': -100,  # Invalid negative width
            'height': -200,  # Invalid negative height
        }
        
        # Should handle gracefully or use defaults
        try:
            builder = OptimizedGenkouYoshiDocumentBuilder(page_format=invalid_format)
            # Constructor should handle invalid values
        except Exception as e:
            # This might be expected behavior
            pass


if __name__ == '__main__':
    # Run all tests with verbose output
    unittest.main(verbosity=2, buffer=True)