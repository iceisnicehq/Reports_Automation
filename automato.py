import os
import glob
import sys
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE

def process_images(directory):
    image_files = sorted(
        [f for f in glob.glob(os.path.join(directory, "*.*")) 
         if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))],
        key=lambda x: os.path.basename(x).lower()
    )
    
    if not image_files:
        print("\n❌ Error: No image files found!")
        print("Supported formats: PNG, JPG/JPEG, GIF")
        return None
    
    figure_data = []
    
    for idx, img_path in enumerate(image_files, 1):
        try:
            with Image.open(img_path) as img:
                # Show image using default viewer
                img.show()
                user_input = input(f"\nIMAGE {idx}/{len(image_files)}\n"
                                 f"File: {os.path.basename(img_path)}\n"
                                 "Enter title (press Enter for default): ").strip()
                
                if user_input:
                    stripped = user_input.rstrip('.')
                    if stripped:
                        processed = stripped[0].upper() + stripped[1:]
                    else:
                        processed = f"Рисунок {idx}"
                    title = f"{processed}."
                else:
                    title = f"Рисунок {idx}."
                figure_data.append({
                    'path': img_path,
                    'title': title
                })
        except Exception as e:
            print(f"\n⚠️ Error processing {os.path.basename(img_path)}: {str(e)}")
            continue
    return figure_data

def find_word_template(directory):
    templates = glob.glob(os.path.join(directory, "*.doc*"))
    
    if not templates:
        print("\n❌ Error: No Word document found in directory!")
        print("Please add a Word template file (.doc or .docx)")
        return None
    
    if len(templates) == 1:
        return templates[0]
    
    print("\nMultiple Word documents found:")
    for i, path in enumerate(templates, 1):
        print(f"{i}. {os.path.basename(path)}")
    
    while True:
        try:
            choice = int(input("\nEnter template number: "))
            if 1 <= choice <= len(templates):
                return templates[choice-1]
            print("Invalid number! Try again.")
        except ValueError:
            print("Please enter a valid number!")
            
def apply_document_styles(doc):
    styles = doc.styles

    normal_style = styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(14)
    normal_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    normal_style.paragraph_format.space_after = Pt(0)

    for heading_level in [1, 2]:
        style_name = f'Heading {heading_level}'
        heading_style = styles[style_name]
        heading_style.font.name = 'Times New Roman'
        heading_style.font.size = Pt(14)
        heading_style.font.bold = False
        heading_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        
def get_output_filename():
    while True:
        name = input("\nEnter report name (without extension): ").strip()
        if not name:
            print("Name cannot be empty!")
            continue
        
        invalid_chars = r'\/:*?"<>|'
        for char in invalid_chars:
            name = name.replace(char, '_')
        
        # Limit length to 50 characters
        name = name[:50]
        return f"{name}.docx"

def get_report_metadata():
    print("\n" + "="*40)
    print(" Report Metadata ".center(40, "="))
    print("="*40)
    
    while True:
        report_num = input("\nEnter report number: ").strip()
        if report_num:
            break
        print("❌ Report number cannot be empty!")
    
    while True:
        report_name = input("Enter report name: ").strip()
        if report_name:
            report_name = report_name[0].upper() + report_name[1:]
            break
        print("❌ Report name cannot be empty!")
    
    while True:
        print("\nSelect discipline:")
        print("1. Администрирование сетей передачи информации")
        print("2. Администрирование операционных систем")
        print("3. Другое (ручной ввод)")
        
        discipline_choice = input("Enter choice (1-3): ").strip()
        
        if discipline_choice in ("1", "2", "3"):
            break
        print("❌ Invalid choice! Please enter 1, 2 or 3")
    
    if discipline_choice == "1":
        discipline = "Администрирование сетей передачи информации"
    elif discipline_choice == "2":
        discipline = "Администрирование операционных систем"
    else:
        while True:
            discipline = input("Enter custom discipline: ").strip()
            if discipline:
                break
            print("❌ Discipline cannot be empty!")
    
    return {
        "num": report_num,
        "name": report_name,
        "discipline": discipline
    }

def replace_placeholders(doc, replacements):
    """Replace placeholders in Word document"""
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    
    # Handle tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, value)

def generate_report(directory, template_path, figure_data, output_filename, metadata):
    """Generate formatted report document with dynamic placeholders"""
    try:
        doc = Document(template_path)
        
        # Replace placeholders first
        replacements = {
            "%DISCIPLINE%": metadata["discipline"],
            "%NUM%": metadata["num"],
            "%REPORT_NAME%": metadata["name"]
        }
        replace_placeholders(doc, replacements)
        
        apply_document_styles(doc)

        doc.add_section()

        for i, item in enumerate(figure_data, 1):
            para = doc.add_paragraph(style='Normal')
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            try:
                run = para.add_run()
                run.add_picture(item['path'], width=Inches(6))
            except Exception as e:
                print(f"⚠️ Couldn't insert image: {os.path.basename(item['path'])}")
                continue
            
            caption = doc.add_paragraph(
                f"Рисунок {i} — {item['title']}", 
                style='Normal'
            )
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph()

        output_path = os.path.join(directory, output_filename)
        doc.save(output_path)
        return output_path
        
    except Exception as e:
        print(f"\n❌ Critical error generating report: {str(e)}")
        return None

def main():
    metadata = get_report_metadata()
    print("\n" + "="*40)
    print(" Lab Report Generator".ljust(39) + "=")
    print("="*40 + "\n")
    
    while True:
        directory = input("Enter full path to working directory: ").strip()
        if os.path.isdir(directory):
            break
        print("Invalid directory! Try again.")
    
    template_path = find_word_template(directory)
    if not template_path:
        sys.exit(1)
    
    figure_data = process_images(directory)
    if not figure_data:
        sys.exit(1)
    
    output_filename = get_output_filename()
    
    report_path = generate_report(directory, template_path, figure_data, output_filename, metadata)    
    
    if report_path:
        print("\n" + "="*40)
        print(f"✅ Report successfully generated at:\n{report_path}")
        print("="*40 + "\n")
    else:
        print("\n❌ Failed to generate report")

if __name__ == "__main__":
    main()
