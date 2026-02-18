import os
from docx import Document
from docx.shared import Inches
from pathlib import Path

def replace_header_assets(doc, logo_path, doc_prefix="DG"):
    for section in doc.sections:
        header = section.header
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Update Doc No Prefix
                    if "Doc No" in cell.text:
                        for paragraph in cell.paragraphs:
                            if "Doc No" in paragraph.text:
                                parts = paragraph.text.split('-')
                                if len(parts) > 1:
                                    paragraph.text = f"Doc No: {doc_prefix}-" + "-".join(parts[1:])
                    
                    # Replace Logo
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if run._element.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"):
                                run.clear()
                                run.add_picture(logo_path, width=Inches(1.5))

def replace_footer_assets(doc, footer_path):
    for section in doc.sections:
        footer = section.footer
        for paragraph in footer.paragraphs:
            paragraph.clear()
        footer.paragraphs[0].add_run().add_picture(footer_path, width=Inches(6.0))

def replace_body_text(doc):
    replacements = {
        "Souring Summits Group Pty (Ltd)": "Dynatron Group (Pty) Ltd",
        "Scope of Work: The provision of engineering consulting, project and maintenance management.": 
        "Scope of Works: Provision of Engineering, Construction and Project Management Services"
    }
    for paragraph in doc.paragraphs:
        for old, new in replacements.items():
            if old in paragraph.text:
                paragraph.text = paragraph.text.replace(old, new)

def process_system():
    # Define exact paths
    base_path = Path(r'C:\Users\Admin\Documents\Work\Mntambo Services\Document Processing\Clients\01\Dynatron Quality Management System QMS ISO 9001 2015')
    assets_path = Path(r'C:\Users\Admin\Documents\Work\Mntambo Services\Document Processing\Clients\01\assets')
    logo_path = assets_path / 'logo' / '01.png'
    footer_path = assets_path / 'footer' / '01.png'
    output_base = Path(r'C:\Users\Admin\Documents\Work\Mntambo Services\Document Processing\Clients\01\outcome')

    # rglob("*.docx") finds all Word docs in ALL subfolders
    for doc_path in base_path.rglob("*.docx"):
        if doc_path.name.startswith("~$"): continue
        
        print(f"Processing: {doc_path.relative_to(base_path)}")
        
        # Determine output path (keeping subfolder structure)
        relative_path = doc_path.relative_to(base_path)
        new_path = output_base / relative_path
        new_path.parent.mkdir(parents=True, exist_ok=True) # Creates subfolders in 'outcome'

        try:
            doc = Document(str(doc_path))
            replace_header_assets(doc, str(logo_path))
            replace_footer_assets(doc, str(footer_path))
            replace_body_text(doc)
            doc.save(str(new_path))
        except Exception as e:
            print(f"!!! Error on {doc_path.name}: {e}")

if __name__ == "__main__":
    process_system()
    print("\n--- Finished! Check the 'outcome' folder. ---")