"""
Interactive page size selector for Genkō Yōshi Tategaki Converter
Based on standard Japanese publishing formats
"""
from rich.console import Console
from rich.prompt import Prompt
from rich.table import Table
from rich import box

class PageSizeSelector:
    """Interactive page size selector using rich library"""
    
    # Detailed Japanese book formats with grid specifications
    # Based on standard Japanese publishing conventions
    PAGE_FORMATS = {
        'bunko': {          # 文庫本 (Mass market paperback) - Most common
            'name': 'Bunko',
            'width': 105, 
            'height': 148,  # A6 size
            'grid': {'columns': 17, 'rows': 24},
            'characters_per_page': 408,
            'margins': {'top': 15, 'bottom': 15, 'inner': 12, 'outer': 8},
            'description': 'Standard mass-market paperback fiction'
        },
        'custom_bunko': {   # Custom Bunko format 111×178mm
            'name': 'Custom Bunko',
            'width': 111, 
            'height': 178,
            'grid': {'columns': 17, 'rows': 26},
            'characters_per_page': 442,
            'margins': {'top': 15, 'bottom': 15, 'inner': 12, 'outer': 8},
            'description': 'Custom Bunko format (111×178mm)'
        },
        'tankobon': {       # 単行本 (First edition hardcover/softcover)
            'name': 'Tankobon',
            'width': 127, 
            'height': 188,  # B6 variant
            'grid': {'columns': 18, 'rows': 28}, 
            'characters_per_page': 504,
            'margins': {'top': 18, 'bottom': 15, 'inner': 15, 'outer': 12},
            'description': 'Standard first edition hardcover/quality paperback'
        },
        'shinsho': {        # 新書 (Non-fiction paperback)
            'name': 'Shinsho',
            'width': 103, 
            'height': 182,
            'grid': {'columns': 16, 'rows': 28},
            'characters_per_page': 448,
            'margins': {'top': 18, 'bottom': 15, 'inner': 12, 'outer': 8},
            'description': 'Standard non-fiction paperback format'
        },
        'a5_standard': {    # A5 (Large format novels)
            'name': 'A5 Standard',
            'width': 148, 
            'height': 210,
            'grid': {'columns': 20, 'rows': 32},
            'characters_per_page': 640,
            'margins': {'top': 20, 'bottom': 20, 'inner': 15, 'outer': 12},
            'description': 'Large format novels and literary works'
        },
        'b6_standard': {    # B6 (Medium hardcover)
            'name': 'B6 Standard',
            'width': 128, 
            'height': 182,
            'grid': {'columns': 18, 'rows': 28},
            'characters_per_page': 504,
            'margins': {'top': 18, 'bottom': 15, 'inner': 12, 'outer': 10},
            'description': 'Medium hardcover format'
        },
        'genkou_yoshi_20x20': {  # Traditional manuscript paper
            'name': 'Genkou Yoshi 20×20',
            'width': 200, 
            'height': 290,  # B4-based traditional size
            'grid': {'columns': 20, 'rows': 20},
            'characters_per_page': 400,
            'margins': {'top': 25, 'bottom': 25, 'inner': 20, 'outer': 20},
            'description': 'Traditional 20×20 genkou yoshi manuscript format'
        },
        'genkou_yoshi_10x20': {  # Traditional manuscript paper variant
            'name': 'Genkou Yoshi 10×20',
            'width': 200, 
            'height': 290,
            'grid': {'columns': 10, 'rows': 20},
            'characters_per_page': 200,
            'margins': {'top': 25, 'bottom': 25, 'inner': 20, 'outer': 20},
            'description': 'Traditional 10×20 genkou yoshi manuscript format'
        },
        'a4': {           # A4 (Standard document size)
            'name': 'A4',
            'width': 210, 
            'height': 297,
            'grid': {'columns': 22, 'rows': 34},
            'characters_per_page': 748,
            'margins': {'top': 25, 'bottom': 25, 'inner': 20, 'outer': 20},
            'description': 'Standard document size'
        },
        'b5': {           # B5 (Large format books)
            'name': 'B5',
            'width': 176, 
            'height': 250,
            'grid': {'columns': 20, 'rows': 30},
            'characters_per_page': 600,
            'margins': {'top': 20, 'bottom': 20, 'inner': 18, 'outer': 15},
            'description': 'Large format books and textbooks'
        },
        'a5': {    # Alternative name for a5_standard
            'name': 'A5',
            'width': 148, 
            'height': 210,
            'grid': {'columns': 20, 'rows': 32},
            'characters_per_page': 640,
            'margins': {'top': 20, 'bottom': 20, 'inner': 15, 'outer': 12},
            'description': 'Large format novels and literary works'
        },
        'b6': {    # Alternative name for b6_standard
            'name': 'B6',
            'width': 128, 
            'height': 182,
            'grid': {'columns': 18, 'rows': 28},
            'characters_per_page': 504,
            'margins': {'top': 18, 'bottom': 15, 'inner': 12, 'outer': 10},
            'description': 'Medium hardcover format'
        },
        'a6': {           # A6 (Small pocket books)
            'name': 'A6',
            'width': 105, 
            'height': 148,
            'grid': {'columns': 16, 'rows': 24},
            'characters_per_page': 384,
            'margins': {'top': 12, 'bottom': 12, 'inner': 10, 'outer': 8},
            'description': 'Small pocket books'
        },
        'custom': {       # Custom user-defined
            'name': 'Custom',
            'width': 0, 
            'height': 0,
            'grid': {'columns': 0, 'rows': 0},
            'characters_per_page': 0,
            'margins': {'top': 0, 'bottom': 0, 'inner': 0, 'outer': 0},
            'description': 'Custom user-defined format'
        }
    }
    
    # Simple list for UI display
    COMMON_SIZES = [
        PAGE_FORMATS['bunko'],
        PAGE_FORMATS['custom_bunko'],
        PAGE_FORMATS['tankobon'],
        PAGE_FORMATS['a5'],
        PAGE_FORMATS['b6'],
        PAGE_FORMATS['a4'],
        PAGE_FORMATS['b5'],
        PAGE_FORMATS['genkou_yoshi_20x20'],
        PAGE_FORMATS['a6'],
        PAGE_FORMATS['shinsho'],
        PAGE_FORMATS['genkou_yoshi_10x20'],
        PAGE_FORMATS['custom'],
    ]
    
    def __init__(self, console=None):
        self.console = console or Console()
    
    @staticmethod
    def calculate_grid_dimensions(page_width, page_height, margins, character_size=None):
        """
        Calculate optimal grid dimensions for given page format
        
        Args:
            page_width (int): Width of page in mm
            page_height (int): Height of page in mm
            margins (dict): Dictionary with 'top', 'bottom', 'inner', 'outer' margins in mm
            character_size (float, optional): Size of each character cell in mm
        
        Returns:
            dict: Grid dimensions and other calculated values
        """
        # Calculate usable text area
        text_width = page_width - margins['inner'] - margins['outer']
        text_height = page_height - margins['top'] - margins['bottom']
        
        # If character size not specified, calculate optimal size
        if character_size is None:
            # Japanese publishing standard is typically square cells
            # Aim for cells between 6-9mm depending on page size
            if text_width < 80:  # Very small format
                character_size = 6
            elif text_width < 100:  # Small format
                character_size = 7
            elif text_width < 150:  # Medium format
                character_size = 8
            else:  # Large format
                character_size = 9
        
        # Calculate grid dimensions
        columns = int(text_width / character_size)
        rows = int(text_height / character_size)
        
        # Apply Japanese publishing conventions
        if columns % 2 == 0:  # Prefer even column counts
            pass  # Keep as is
        else:
            columns = columns - 1  # Reduce to even number
            
        # Ensure minimum dimensions
        columns = max(10, columns)
        rows = max(15, rows)
        
        return {
            'columns': columns,
            'rows': rows,
            'characters_per_page': columns * rows,
            'character_size': character_size
        }
        
    def show_sizes(self):
        """Display the available page sizes in a rich table"""
        table = Table(title="Common Japanese Book Formats", box=box.ROUNDED)
        table.add_column("#", style="cyan")
        table.add_column("Format", style="green")
        table.add_column("Size (W×H)", style="blue")
        table.add_column("Grid", style="magenta")
        table.add_column("Description", style="yellow")
        
        for i, size in enumerate(self.COMMON_SIZES, 1):
            if size["name"] == "Custom":
                grid_info = "Custom"
            else:
                grid_info = f"{size['grid']['columns']}×{size['grid']['rows']} ({size['characters_per_page']} chars)"
                
            table.add_row(
                str(i), 
                size["name"],
                f"{size['width']}×{size['height']}mm", 
                grid_info,
                size["description"]
            )
            
        self.console.print(table)
        
    def select_page_size(self):
        """Prompt user to select a page size and return the dimensions with grid information"""
        self.show_sizes()
        self.console.print("\n[bold cyan]Select a page size format:[/bold cyan]")

        # Always use the provided console for Prompt.ask to ensure interactive input
        choice = Prompt.ask(
            "Enter selection",
            choices=[str(i) for i in range(1, len(self.COMMON_SIZES) + 1)],
            default="1",  # Default to Bunko format
            console=self.console  # Force use of this console for input
        )

        selected = self.COMMON_SIZES[int(choice) - 1].copy()

        # Handle custom size
        if selected["name"] == "Custom":
            self.console.print("\n[bold cyan]Enter custom page dimensions:[/bold cyan]")
            width = int(Prompt.ask("Width (mm)", default="111", console=self.console))
            height = int(Prompt.ask("Height (mm)", default="178", console=self.console))

            # Get margins
            self.console.print("\n[bold cyan]Enter margins:[/bold cyan]")
            top = int(Prompt.ask("Top margin (mm)", default="15", console=self.console))
            bottom = int(Prompt.ask("Bottom margin (mm)", default="15", console=self.console))
            inner = int(Prompt.ask("Inner margin (mm)", default="12", console=self.console))
            outer = int(Prompt.ask("Outer margin (mm)", default="8", console=self.console))

            margins = {'top': top, 'bottom': bottom, 'inner': inner, 'outer': outer}

            # Calculate optimal grid dimensions
            grid = self.calculate_grid_dimensions(width, height, margins)

            custom_size = {
                "name": "Custom", 
                "width": width, 
                "height": height,
                "grid": {'columns': grid['columns'], 'rows': grid['rows']},
                "characters_per_page": grid['characters_per_page'],
                "margins": margins,
                "description": "Custom user-defined format",
                "character_size": grid['character_size']
            }

            self.console.print(f"\n[bold green]Custom format:[/bold green] {width}×{height}mm with {grid['columns']}×{grid['rows']} grid")
            self.console.print(f"[bold green]Cell size:[/bold green] {grid['character_size']}mm, [bold green]Characters per page:[/bold green] {grid['characters_per_page']}")

            return custom_size

        self.console.print(f"\n[bold green]Selected:[/bold green] {selected['name']} ({selected['width']}×{selected['height']}mm)")
        self.console.print(f"[bold green]Grid:[/bold green] {selected['grid']['columns']}×{selected['grid']['rows']}, [bold green]Characters per page:[/bold green] {selected['characters_per_page']}")

        return selected


# Example usage when run directly
if __name__ == "__main__":
    selector = PageSizeSelector()
    selected_size = selector.select_page_size()
    print(f"Selected: {selected_size['name']} - {selected_size['width']}×{selected_size['height']}mm")
