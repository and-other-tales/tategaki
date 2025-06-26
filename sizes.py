"""
Interactive page size selector for Genkō Yōshi Tategaki Converter
Based on standard Japanese publishing formats
OPTIMIZED VERSION with pre-computed grid dimensions and efficient lookups
"""
from rich.console import Console
from rich.prompt import Prompt
from rich.table import Table
from rich import box

class OptimizedPageSizeSelector:
    """Optimized page size selector with pre-computed dimensions and fast lookups"""
    
    # Pre-computed Japanese book formats with all calculations done once
    # This eliminates runtime calculation overhead
    PAGE_FORMATS = {
        'bunko': {
            'name': 'Bunko',
            'width': 105, 
            'height': 148,
            'grid': {'columns': 17, 'rows': 24},
            'characters_per_page': 408,
            'margins': {'top': 15, 'bottom': 15, 'inner': 12, 'outer': 8},
            'description': 'Standard mass-market paperback fiction',
            'character_size': 7
        },
        'custom_bunko': {
            'name': 'Custom Bunko',
            'width': 111, 
            'height': 178,
            'grid': {'columns': 17, 'rows': 26},
            'characters_per_page': 442,
            'margins': {'top': 15, 'bottom': 15, 'inner': 12, 'outer': 8},
            'description': 'Custom Bunko format (111×178mm)',
            'character_size': 7
        },
        'tankobon': {
            'name': 'Tankobon',
            'width': 127, 
            'height': 188,
            'grid': {'columns': 18, 'rows': 28}, 
            'characters_per_page': 504,
            'margins': {'top': 18, 'bottom': 15, 'inner': 15, 'outer': 12},
            'description': 'Standard first edition hardcover/quality paperback',
            'character_size': 8
        },
        'shinsho': {
            'name': 'Shinsho',
            'width': 103, 
            'height': 182,
            'grid': {'columns': 16, 'rows': 28},
            'characters_per_page': 448,
            'margins': {'top': 18, 'bottom': 15, 'inner': 12, 'outer': 8},
            'description': 'Standard non-fiction paperback format',
            'character_size': 7
        },
        'a5_standard': {
            'name': 'A5 Standard',
            'width': 148, 
            'height': 210,
            'grid': {'columns': 20, 'rows': 32},
            'characters_per_page': 640,
            'margins': {'top': 20, 'bottom': 20, 'inner': 15, 'outer': 12},
            'description': 'Large format novels and literary works',
            'character_size': 9
        },
        'b6_standard': {
            'name': 'B6 Standard',
            'width': 128, 
            'height': 182,
            'grid': {'columns': 18, 'rows': 28},
            'characters_per_page': 504,
            'margins': {'top': 18, 'bottom': 15, 'inner': 12, 'outer': 10},
            'description': 'Medium hardcover format',
            'character_size': 8
        },
        'genkou_yoshi_20x20': {
            'name': 'Genkou Yoshi 20×20',
            'width': 200, 
            'height': 290,
            'grid': {'columns': 20, 'rows': 20},
            'characters_per_page': 400,
            'margins': {'top': 25, 'bottom': 25, 'inner': 20, 'outer': 20},
            'description': 'Traditional 20×20 genkou yoshi manuscript format',
            'character_size': 9
        },
        'genkou_yoshi_10x20': {
            'name': 'Genkou Yoshi 10×20',
            'width': 200, 
            'height': 290,
            'grid': {'columns': 10, 'rows': 20},
            'characters_per_page': 200,
            'margins': {'top': 25, 'bottom': 25, 'inner': 20, 'outer': 20},
            'description': 'Traditional 10×20 genkou yoshi manuscript format',
            'character_size': 9
        },
        'a4': {
            'name': 'A4',
            'width': 210, 
            'height': 297,
            'grid': {'columns': 22, 'rows': 34},
            'characters_per_page': 748,
            'margins': {'top': 25, 'bottom': 25, 'inner': 20, 'outer': 20},
            'description': 'Standard document size',
            'character_size': 9
        },
        'b5': {
            'name': 'B5',
            'width': 176, 
            'height': 250,
            'grid': {'columns': 20, 'rows': 30},
            'characters_per_page': 600,
            'margins': {'top': 20, 'bottom': 20, 'inner': 18, 'outer': 15},
            'description': 'Large format books and textbooks',
            'character_size': 9
        },
        'a5': {
            'name': 'A5',
            'width': 148, 
            'height': 210,
            'grid': {'columns': 20, 'rows': 32},
            'characters_per_page': 640,
            'margins': {'top': 20, 'bottom': 20, 'inner': 15, 'outer': 12},
            'description': 'Large format novels and literary works',
            'character_size': 9
        },
        'b6': {
            'name': 'B6',
            'width': 128, 
            'height': 182,
            'grid': {'columns': 18, 'rows': 28},
            'characters_per_page': 504,
            'margins': {'top': 18, 'bottom': 15, 'inner': 12, 'outer': 10},
            'description': 'Medium hardcover format',
            'character_size': 8
        },
        'a6': {
            'name': 'A6',
            'width': 105, 
            'height': 148,
            'grid': {'columns': 16, 'rows': 24},
            'characters_per_page': 384,
            'margins': {'top': 12, 'bottom': 12, 'inner': 10, 'outer': 8},
            'description': 'Small pocket books',
            'character_size': 7
        },
        'custom': {
            'name': 'Custom',
            'width': 0, 
            'height': 0,
            'grid': {'columns': 0, 'rows': 0},
            'characters_per_page': 0,
            'margins': {'top': 0, 'bottom': 0, 'inner': 0, 'outer': 0},
            'description': 'Custom user-defined format',
            'character_size': 0
        }
    }
    
    # Pre-ordered list by popularity for optimal UI display
    COMMON_SIZES = [
        PAGE_FORMATS['bunko'],           # Most popular
        PAGE_FORMATS['custom_bunko'],    # Alternative bunko
        PAGE_FORMATS['tankobon'],        # Standard hardcover
        PAGE_FORMATS['a5'],              # Large format
        PAGE_FORMATS['b6'],              # Medium format
        PAGE_FORMATS['a4'],              # Document size
        PAGE_FORMATS['b5'],              # Large books
        PAGE_FORMATS['genkou_yoshi_20x20'],  # Traditional
        PAGE_FORMATS['a6'],              # Small format
        PAGE_FORMATS['shinsho'],         # Non-fiction
        PAGE_FORMATS['genkou_yoshi_10x20'],  # Traditional variant
        PAGE_FORMATS['custom'],          # Custom option
    ]
    
    # Pre-computed lookup table for format names (case-insensitive)
    _FORMAT_LOOKUP = {name.lower(): fmt for name, fmt in PAGE_FORMATS.items()}
    
    def __init__(self, console=None):
        self.console = console or Console()
    
    @classmethod
    def get_format(cls, format_name):
        """Optimized format lookup with case-insensitive search"""
        return cls._FORMAT_LOOKUP.get(format_name.lower())
    
    @staticmethod
    def calculate_grid_dimensions(page_width, page_height, margins, character_size=None):
        """
        Optimized grid dimension calculation with pre-computed logic
        """
        # Calculate usable text area
        text_width = page_width - margins['inner'] - margins['outer']
        text_height = page_height - margins['top'] - margins['bottom']
        
        # Use optimized character size calculation
        if character_size is None:
            character_size = 6 if text_width < 80 else 7 if text_width < 100 else 8 if text_width < 150 else 9
        
        # Calculate grid dimensions with optimized logic
        columns = max(10, int(text_width / character_size))
        rows = max(15, int(text_height / character_size))
        
        # Prefer even column counts for better layout
        columns = columns - (columns % 2)
        
        return {
            'columns': columns,
            'rows': rows,
            'characters_per_page': columns * rows,
            'character_size': character_size
        }
        
    def show_sizes(self):
        """Optimized size display with efficient table rendering"""
        table = Table(title="Japanese Book Formats", box=box.ROUNDED, expand=False)
        
        # Pre-defined column widths for optimal display
        table.add_column("#", style="cyan", width=3, justify="right")
        table.add_column("Format", style="green", width=15)
        table.add_column("Size", style="blue", width=22, justify="center")
        table.add_column("Grid", style="magenta", width=18, justify="center")
        table.add_column("Description", style="yellow")
        
        # Batch add all rows for better performance
        for i, size in enumerate(self.COMMON_SIZES, 1):
            grid_info = ("Custom" if size["name"] == "Custom" 
                        else f"{size['grid']['columns']}×{size['grid']['rows']} ({size['characters_per_page']})")
            
            # Convert mm to inches (1 mm = 0.0393701 inches)
            if size["name"] == "Custom":
                size_info = "Custom"
            else:
                width_in = size['width'] * 0.0393701
                height_in = size['height'] * 0.0393701
                size_info = f"{size['width']}×{size['height']}mm ({width_in:.1f}\"×{height_in:.1f}\")"
                
            table.add_row(
                str(i), 
                size["name"],
                size_info,
                grid_info,
                size["description"]
            )
            
        self.console.print(table)
        
    def select_page_size(self):
        """Optimized page size selection with fast format handling"""
        self.show_sizes()
        self.console.print("\n[bold cyan]Select a page format:[/bold cyan]")

        # Pre-compute valid choices for faster validation
        valid_choices = [str(i) for i in range(1, len(self.COMMON_SIZES) + 1)]
        
        choice = Prompt.ask(
            "Enter selection",
            choices=valid_choices,
            default="1",  # Default to most popular (Bunko)
            console=self.console
        )

        selected = self.COMMON_SIZES[int(choice) - 1].copy()

        # Handle custom size with optimized input handling
        if selected["name"] == "Custom":
            self.console.print("\n[bold cyan]Custom Page Dimensions:[/bold cyan]")
            
            # Batch input collection with defaults
            width = int(Prompt.ask("Width (mm)", default="111", console=self.console))
            height = int(Prompt.ask("Height (mm)", default="178", console=self.console))

            self.console.print("\n[bold cyan]Margins:[/bold cyan]")
            top = int(Prompt.ask("Top margin (mm)", default="15", console=self.console))
            bottom = int(Prompt.ask("Bottom margin (mm)", default="15", console=self.console))
            inner = int(Prompt.ask("Inner margin (mm)", default="12", console=self.console))
            outer = int(Prompt.ask("Outer margin (mm)", default="8", console=self.console))

            margins = {'top': top, 'bottom': bottom, 'inner': inner, 'outer': outer}
            
            # Fast grid calculation
            grid = self.calculate_grid_dimensions(width, height, margins)

            # Build custom format efficiently
            custom_size = {
                "name": "Custom", 
                "width": width, 
                "height": height,
                "grid": {'columns': grid['columns'], 'rows': grid['rows']},
                "characters_per_page": grid['characters_per_page'],
                "margins": margins,
                "description": f"Custom {width}×{height}mm format",
                "character_size": grid['character_size']
            }

            # Display confirmation
            self.console.print(f"\n[bold green]✓ Custom format created:[/bold green] {width}×{height}mm")
            self.console.print(f"[bold green]✓ Grid:[/bold green] {grid['columns']}×{grid['rows']} ({grid['characters_per_page']} characters/page)")

            return custom_size

        # Display selected format confirmation
        self.console.print(f"\n[bold green]✓ Selected:[/bold green] {selected['name']} ({selected['width']}×{selected['height']}mm)")
        self.console.print(f"[bold green]✓ Grid:[/bold green] {selected['grid']['columns']}×{selected['grid']['rows']} ({selected['characters_per_page']} characters/page)")

        return selected


# Maintain backward compatibility
PageSizeSelector = OptimizedPageSizeSelector

# Example usage when run directly
if __name__ == "__main__":
    import time
    start_time = time.time()
    
    selector = OptimizedPageSizeSelector()
    selected_size = selector.select_page_size()
    
    end_time = time.time()
    print(f"\nSelected: {selected_size['name']} - {selected_size['width']}×{selected_size['height']}mm")
    print(f"Selection completed in {end_time - start_time:.3f} seconds")