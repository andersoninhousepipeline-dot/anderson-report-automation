import os
import io
from docx import Document
from lxml import etree
from PIL import Image as PILImage

class TemplateParser:
    """Parses DOCX template to extract structure, styles, and assets."""
    
    def __init__(self, template_path, assets_dir="extracted_assets"):
        self.template_path = template_path
        self.assets_dir = assets_dir
        self.doc = Document(template_path)
        
        if not os.path.exists(self.assets_dir):
            os.makedirs(self.assets_dir)
            
        self.logo_path = None
        self.colors = {
            'patient_info_bg': 'F1F1F7',
            'results_header_bg': 'FABF8F',
            'chromosome_table_bg': 'F2F2F2',
            'notes_bg': 'F2F2F2'
        }
        
    def extract_assets(self):
        """Extract images from the docx template."""
        image_count = 0
        for rel in self.doc.part.rels.values():
            if "image" in rel.target_ref:
                image_data = rel.target_part.blob
                image_ext = os.path.splitext(rel.target_ref)[1]
                image_name = f"template_image_{image_count}{image_ext}"
                
                # Check if it's likely the logo (usually first image or specific size)
                # For now, we'll save it as chrominst_logo if it's the first one
                if image_count == 0:
                    image_name = "chrominst_logo.png"
                    self.logo_path = os.path.join(self.assets_dir, image_name)
                
                output_path = os.path.join(self.assets_dir, image_name)
                with open(output_path, "wb") as f:
                    f.write(image_data)
                
                image_count += 1
        
        return image_count

    def get_table_info(self):
        """Extract background colors and structure from tables."""
        table_info = []
        for i, table in enumerate(self.doc.tables):
            shading_colors = []
            for row in table.rows:
                row_colors = []
                for cell in row.cells:
                    fill = None
                    try:
                        shading = cell._element.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
                        if shading is not None:
                            fill = shading.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
                    except:
                        pass
                    row_colors.append(fill)
                shading_colors.append(row_colors)
            
            table_info.append({
                'index': i,
                'rows': len(table.rows),
                'cols': len(table.columns),
                'shading': shading_colors
            })
        return table_info

if __name__ == "__main__":
    template = "/data/Sethu/PGTA-Report/Template/Chrominst- Surfseq -template - PGT-A report_withlogo.docx"
    parser = TemplateParser(template)
    count = parser.extract_assets()
    print(f"Extracted {count} images.")
    info = parser.get_table_info()
    for t in info:
        print(f"Table {t['index']}: {t['rows']}x{t['cols']}")
