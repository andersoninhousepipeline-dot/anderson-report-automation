import pdfplumber
import os

pdf_path = "/data/Sethu/PGTA-Report/Priya - PGT-A report_withlogo.pdf"
output_dir = "/data/Sethu/PGTA-Report/debug_extract"
os.makedirs(output_dir, exist_ok=True)

with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages):
        print(f"Processing page {i+1}...")
        for j, image in enumerate(page.images):
            # Extract image bytes
            img_name = f"page{i+1}_img{j+1}.png"
            img_path = os.path.join(output_dir, img_name)
            
            # Simple extraction via pdfplumber's crop and save
            try:
                # Get the bbox of the image
                bbox = (image['x0'], pdf.pages[i].height - image['y1'], image['x1'], pdf.pages[i].height - image['y0'])
                page.crop(bbox).to_image(resolution=300).save(img_path)
                print(f"  Saved {img_name}")
            except Exception as e:
                print(f"  Failed {img_name}: {e}")
