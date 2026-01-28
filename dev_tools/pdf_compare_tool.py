"""
PDF Formatting Analyzer & Comparison Tool
Extracts detailed formatting information from PDFs including:
- Fonts, sizes, colors
- Images and their positions
- Tables and borders
- Layout structure
- Colors and backgrounds
"""

import fitz  # PyMuPDF
import json
from collections import defaultdict
from pathlib import Path


class PDFAnalyzer:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.analysis = {
            'filename': Path(pdf_path).name,
            'page_count': len(self.doc),
            'pages': []
        }
    
    def rgb_to_hex(self, color):
        """Convert RGB color (int or tuple) to hex color code"""
        if color is None:
            return None
            
        if isinstance(color, (list, tuple)):
            # PyMuPDF draws often return (r,g,b) as 0-1 floats
            if all(0 <= c <= 1 for c in color):
                r, g, b = [int(max(0, min(255, c * 255))) for c in color[:3]]
            else:
                r, g, b = [int(max(0, min(255, c))) for c in color[:3]]
            return f"#{r:02X}{g:02X}{b:02X}"
            
        # PyMuPDF text spans often store color as integer
        try:
            rgb_int = int(color)
            r = (rgb_int >> 16) & 0xFF
            g = (rgb_int >> 8) & 0xFF
            b = rgb_int & 0xFF
            return f"#{r:02X}{g:02X}{b:02X}"
        except (ValueError, TypeError):
            return str(color)
    
    def analyze_fonts(self, page):
        """Extract all font information from a page"""
        fonts = {}
        text_dict = page.get_text("dict")
        
        for block in text_dict.get("blocks", []):
            if block.get("type") == 0:  # Text block
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        font_key = f"{span['font']}_{span['size']}"
                        if font_key not in fonts:
                            fonts[font_key] = {
                                'font': span['font'],
                                'size': round(span['size'], 2),
                                'color': self.rgb_to_hex(span.get('color', 0)),
                                'flags': span.get('flags', 0),
                                'is_bold': bool(span.get('flags', 0) & 2**4),
                                'is_italic': bool(span.get('flags', 0) & 2**1),
                                'sample_text': span['text'][:50]
                            }
        return fonts
    
    def analyze_images(self, page):
        """Extract all images and their properties"""
        images = []
        image_list = page.get_images(full=True)
        
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = self.doc.extract_image(xref)
            
            # Get image position
            img_rects = page.get_image_rects(xref)
            
            image_info = {
                'xref': xref,
                'width': base_image['width'],
                'height': base_image['height'],
                'colorspace': base_image['colorspace'],
                'bpc': base_image['bpc'],  # bits per component
                'size_bytes': len(base_image['image']),
                'ext': base_image['ext'],
                'positions': []
            }
            
            for rect in img_rects:
                image_info['positions'].append({
                    'x0': round(rect.x0, 2),
                    'y0': round(rect.y0, 2),
                    'x1': round(rect.x1, 2),
                    'y1': round(rect.y1, 2),
                    'width': round(rect.width, 2),
                    'height': round(rect.height, 2)
                })
            
            images.append(image_info)
        
        return images
    
    def analyze_text_blocks(self, page):
        """Analyze text blocks and their properties"""
        blocks = []
        text_dict = page.get_text("dict")
        
        for block in text_dict.get("blocks", []):
            if block.get("type") == 0:  # Text block
                block_info = {
                    'bbox': {
                        'x0': round(block['bbox'][0], 2),
                        'y0': round(block['bbox'][1], 2),
                        'x1': round(block['bbox'][2], 2),
                        'y1': round(block['bbox'][3], 2)
                    },
                    'lines': []
                }
                
                for line in block.get("lines", []):
                    line_info = {
                        'bbox': {
                            'x0': round(line['bbox'][0], 2),
                            'y0': round(line['bbox'][1], 2),
                            'x1': round(line['bbox'][2], 2),
                            'y1': round(line['bbox'][3], 2)
                        },
                        'spans': []
                    }
                    
                    for span in line.get("spans", []):
                        span_info = {
                            'text': span['text'],
                            'font': span['font'],
                            'size': round(span['size'], 2),
                            'color': self.rgb_to_hex(span.get('color', 0)),
                            'is_bold': bool(span.get('flags', 0) & 2**4),
                            'is_italic': bool(span.get('flags', 0) & 2**1),
                        }
                        line_info['spans'].append(span_info)
                    
                    block_info['lines'].append(line_info)
                
                blocks.append(block_info)
        
        return blocks
    
    def analyze_drawings(self, page):
        """Extract drawing objects (lines, rectangles, fills)"""
        drawings = []
        
        # Get drawings as paths
        paths = page.get_drawings()
        
        for path in paths:
            drawing_info = {
                'type': path.get('type', 'unknown'),
                'rect': {
                    'x0': round(path['rect'].x0, 2),
                    'y0': round(path['rect'].y0, 2),
                    'x1': round(path['rect'].x1, 2),
                    'y1': round(path['rect'].y1, 2)
                },
                'color': self.rgb_to_hex(path.get('color')),
                'fill': self.rgb_to_hex(path.get('fill')),
                'width': path.get('width', 0),
            }
            drawings.append(drawing_info)
        
        return drawings
    
    def detect_tables(self, page):
        """Attempt to detect table structures"""
        # This is a simple heuristic-based detection
        drawings = page.get_drawings()
        
        # Look for rectangular shapes that might be table cells
        horizontal_lines = []
        vertical_lines = []
        rectangles = []
        
        for drawing in drawings:
            rect = drawing['rect']
            width = rect.x1 - rect.x0
            height = rect.y1 - rect.y0
            
            # Horizontal lines (height very small)
            if height < 2 and width > 20:
                horizontal_lines.append({
                    'y': round(rect.y0, 2),
                    'x0': round(rect.x0, 2),
                    'x1': round(rect.x1, 2),
                    'color': self.rgb_to_hex(drawing.get('color'))
                })
            
            # Vertical lines (width very small)
            elif width < 2 and height > 20:
                vertical_lines.append({
                    'x': round(rect.x0, 2),
                    'y0': round(rect.y0, 2),
                    'y1': round(rect.y1, 2),
                    'color': self.rgb_to_hex(drawing.get('color'))
                })
            
            # Rectangles (potential cells or backgrounds)
            elif width > 10 and height > 10:
                rectangles.append({
                    'x0': round(rect.x0, 2),
                    'y0': round(rect.y0, 2),
                    'x1': round(rect.x1, 2),
                    'y1': round(rect.y1, 2),
                    'border_color': self.rgb_to_hex(drawing.get('color')),
                    'fill_color': self.rgb_to_hex(drawing.get('fill')),
                    'width': drawing.get('width', 0)
                })
        
        return {
            'horizontal_lines': horizontal_lines,
            'vertical_lines': vertical_lines,
            'rectangles': rectangles,
            'potential_table': len(horizontal_lines) > 2 and len(vertical_lines) > 2
        }
    
    def analyze_page(self, page_num):
        """Comprehensive analysis of a single page"""
        page = self.doc[page_num]
        
        page_analysis = {
            'page_number': page_num + 1,
            'dimensions': {
                'width': round(page.rect.width, 2),
                'height': round(page.rect.height, 2)
            },
            'fonts': self.analyze_fonts(page),
            'images': self.analyze_images(page),
            'text_blocks': self.analyze_text_blocks(page),
            'drawings': self.analyze_drawings(page),
            'tables': self.detect_tables(page)
        }
        
        return page_analysis
    
    def analyze_all(self):
        """Analyze all pages in the PDF"""
        for page_num in range(len(self.doc)):
            page_analysis = self.analyze_page(page_num)
            self.analysis['pages'].append(page_analysis)
        
        return self.analysis
    
    def save_analysis(self, output_path):
        """Save analysis to JSON file"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.analysis, f, indent=2, ensure_ascii=False)
        print(f"Analysis saved to: {output_path}")
    
    def print_summary(self):
        """Print a human-readable summary"""
        print(f"\n{'='*80}")
        print(f"PDF ANALYSIS SUMMARY: {self.analysis['filename']}")
        print(f"{'='*80}\n")
        
        for page in self.analysis['pages']:
            print(f"\n--- PAGE {page['page_number']} ---")
            print(f"Dimensions: {page['dimensions']['width']} x {page['dimensions']['height']} pts")
            
            # Fonts
            print(f"\nFonts Used ({len(page['fonts'])}):")
            for font_key, font_info in page['fonts'].items():
                bold = " [BOLD]" if font_info['is_bold'] else ""
                italic = " [ITALIC]" if font_info['is_italic'] else ""
                print(f"  • {font_info['font']} - {font_info['size']}pt - {font_info['color']}{bold}{italic}")
                print(f"    Sample: '{font_info['sample_text']}'")
            
            # Images
            if page['images']:
                print(f"\nImages ({len(page['images'])}):")
                for idx, img in enumerate(page['images']):
                    print(f"  • Image {idx+1}: {img['width']}x{img['height']}px, {img['ext']}, {img['size_bytes']} bytes")
                    if img['positions']:
                        pos = img['positions'][0]
                        print(f"    Position: ({pos['x0']}, {pos['y0']}) - ({pos['x1']}, {pos['y1']})")
            
            # Tables
            if page['tables']['potential_table']:
                print(f"\nTable Structure Detected:")
                print(f"  • Horizontal lines: {len(page['tables']['horizontal_lines'])}")
                print(f"  • Vertical lines: {len(page['tables']['vertical_lines'])}")
                print(f"  • Rectangles/Cells: {len(page['tables']['rectangles'])}")
                
                # Show rectangle fills (backgrounds)
                filled_rects = [r for r in page['tables']['rectangles'] if r['fill_color']]
                if filled_rects:
                    print(f"\n  Background Colors Found:")
                    unique_fills = set(r['fill_color'] for r in filled_rects)
                    for color in unique_fills:
                        print(f"    • {color}")
            
            print("\n" + "-"*80)


class PDFComparator:
    def __init__(self, pdf1_path, pdf2_path):
        self.analyzer1 = PDFAnalyzer(pdf1_path)
        self.analyzer2 = PDFAnalyzer(pdf2_path)
        
        self.analysis1 = self.analyzer1.analyze_all()
        self.analysis2 = self.analyzer2.analyze_all()
    
    def compare_fonts(self):
        """Compare fonts between two PDFs"""
        print(f"\n{'='*80}")
        print("FONT COMPARISON")
        print(f"{'='*80}\n")
        
        for page_num in range(min(len(self.analysis1['pages']), len(self.analysis2['pages']))):
            page1 = self.analysis1['pages'][page_num]
            page2 = self.analysis2['pages'][page_num]
            
            fonts1 = set(f"{f['font']} {f['size']}pt" for f in page1['fonts'].values())
            fonts2 = set(f"{f['font']} {f['size']}pt" for f in page2['fonts'].values())
            
            print(f"Page {page_num + 1}:")
            print(f"  Only in {self.analysis1['filename']}: {fonts1 - fonts2}")
            print(f"  Only in {self.analysis2['filename']}: {fonts2 - fonts1}")
            print()
    
    def compare_images(self):
        """Compare images between two PDFs"""
        print(f"\n{'='*80}")
        print("IMAGE COMPARISON")
        print(f"{'='*80}\n")
        
        for page_num in range(min(len(self.analysis1['pages']), len(self.analysis2['pages']))):
            page1 = self.analysis1['pages'][page_num]
            page2 = self.analysis2['pages'][page_num]
            
            print(f"Page {page_num + 1}:")
            print(f"  {self.analysis1['filename']}: {len(page1['images'])} images")
            print(f"  {self.analysis2['filename']}: {len(page2['images'])} images")
            
            if len(page1['images']) != len(page2['images']):
                print(f"  ⚠️  DIFFERENCE: Image count mismatch!")
            print()
    
    def compare_colors(self):
        """Compare color usage between two PDFs"""
        print(f"\n{'='*80}")
        print("COLOR COMPARISON")
        print(f"{'='*80}\n")
        
        for page_num in range(min(len(self.analysis1['pages']), len(self.analysis2['pages']))):
            page1 = self.analysis1['pages'][page_num]
            page2 = self.analysis2['pages'][page_num]
            
            # Collect all colors used
            colors1 = set()
            colors2 = set()
            
            # From text
            for font in page1['fonts'].values():
                if font['color']:
                    colors1.add(font['color'])
            for font in page2['fonts'].values():
                if font['color']:
                    colors2.add(font['color'])
            
            # From drawings/fills
            for drawing in page1['drawings']:
                if drawing['color']:
                    colors1.add(drawing['color'])
                if drawing['fill']:
                    colors1.add(drawing['fill'])
            for drawing in page2['drawings']:
                if drawing['color']:
                    colors2.add(drawing['color'])
                if drawing['fill']:
                    colors2.add(drawing['fill'])
            
            print(f"Page {page_num + 1}:")
            print(f"  Colors in {self.analysis1['filename']}: {sorted(colors1)}")
            print(f"  Colors in {self.analysis2['filename']}: {sorted(colors2)}")
            print(f"  Only in {self.analysis1['filename']}: {sorted(colors1 - colors2)}")
            print(f"  Only in {self.analysis2['filename']}: {sorted(colors2 - colors1)}")
            print()
    
    def compare_all(self):
        """Run all comparisons"""
        self.compare_fonts()
        self.compare_images()
        self.compare_colors()


if __name__ == "__main__":
    source_pdf = "Priya - PGT-A report_withlogo.pdf"
    rendered_pdf = "test_report.pdf"
    
    import sys
    import os
    
    if not os.path.exists(source_pdf):
        print(f"Error: Source PDF not found at {source_pdf}")
        sys.exit(1)
        
    if not os.path.exists(rendered_pdf):
        print(f"Error: Rendered PDF not found at {rendered_pdf}")
        sys.exit(1)
        
    comparator = PDFComparator(source_pdf, rendered_pdf)
    comparator.compare_all()
