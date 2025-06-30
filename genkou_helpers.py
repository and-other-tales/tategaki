"""
Helper methods for genkou yoshi table configuration
"""

def configure_genkou_table(table, cell_width, cell_height):
    """
    Configure table to look like traditional genkou yoshi manuscript paper
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    
    # Set table properties
    tbl = table._tbl
    tblPr = tbl.tblPr
    
    # Set table layout to fixed for consistent cell sizes
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)
    
    # Set table borders for grid appearance
    tblBorders = OxmlElement('w:tblBorders')
    border_style = 'single'
    border_size = '4'  # Thin border
    border_color = '000000'  # Black
    
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), border_style)
        border.set(qn('w:sz'), border_size)
        border.set(qn('w:color'), border_color)
        tblBorders.append(border)
    
    tblPr.append(tblBorders)
    
    # Set column widths to be uniform
    for col in table.columns:
        col.width = cell_width

def configure_genkou_cell(cell, character, font_size_points, font_name):
    """
    Configure individual cell for genkou yoshi character placement
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    
    # Clear existing cell content
    cell.text = ''
    
    # Configure cell properties
    tcPr = cell._tc.tcPr
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        cell._tc.insert(0, tcPr)
    
    # Set cell vertical alignment to center
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), 'center')
    tcPr.append(vAlign)
    
    # Set text direction to vertical for the cell
    textDirection = OxmlElement('w:textDirection')
    textDirection.set(qn('w:val'), 'tbRl')
    tcPr.append(textDirection)
    
    # Add character to cell if present
    if character and character.strip():
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        # Configure paragraph for vertical text
        pPr = paragraph._p.pPr
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            paragraph._p.insert(0, pPr)
        
        # Set paragraph text direction
        textDir = OxmlElement('w:textDirection')
        textDir.set(qn('w:val'), 'tbRl')
        pPr.append(textDir)
        
        # Set paragraph alignment
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'center')
        pPr.append(jc)
        
        # Add the character with proper formatting
        run = paragraph.add_run(character)
        run.font.name = font_name
        run.font.size = Pt(font_size_points)
    
    # Set cell margins to minimal for grid appearance
    tcMar = OxmlElement('w:tcMar')
    for margin in ['top', 'left', 'bottom', 'right']:
        mar = OxmlElement(f'w:{margin}')
        mar.set(qn('w:w'), '36')  # 2pt margin
        mar.set(qn('w:type'), 'dxa')
        tcMar.append(mar)
    tcPr.append(tcMar)