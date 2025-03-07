import os
import glob
import sys
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE

def process_images(directory):
    """Process images with proper resource handling"""
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
                
                # Get title with smart formatting
                title = input(f"\nIMAGE {idx}/{len(image_files)}\n"
                            f"File: {os.path.basename(img_path)}\n"
                            "Enter title (press Enter for default): ").strip()
                
                # Format title with single trailing dot
                if not title:
                    title = f"Figure {idx}"
                title = title.rstrip('.') + '.'
                
                figure_data.append({
                    'path': img_path,
                    'title': title
                })
                
        except Exception as e:
            print(f"\n⚠️ Error processing {os.path.basename(img_path)}: {str(e)}")
            continue
            
    return figure_data

def find_word_template(directory):
    """Find and select Word template in directory"""
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
    """Set document-wide formatting styles"""
    styles = doc.styles

    # Set Normal style for body text
    normal_style = styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(14)
    normal_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    normal_style.paragraph_format.space_after = Pt(0)

    # Configure heading styles
    for heading_level in [1, 2]:
        style_name = f'Heading {heading_level}'
        heading_style = styles[style_name]
        heading_style.font.name = 'Times New Roman'
        heading_style.font.size = Pt(14)
        heading_style.font.bold = False
        heading_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        
def generate_report(directory, template_path, figure_data):
    """Generate formatted report document"""
    try:
        doc = Document(template_path)
        apply_document_styles(doc)

        # Add new section for figures
        doc.add_section()

        # Add figures with captions
        for i, item in enumerate(figure_data, 1):
            # Add image
            para = doc.add_paragraph(style='Normal')
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            try:
                run = para.add_run()
                run.add_picture(item['path'], width=Inches(6))
            except Exception as e:
                print(f"⚠️ Couldn't insert image: {os.path.basename(item['path'])}")
                continue
            
            # Add caption
            caption = doc.add_paragraph(
                f"Figure {i} — {item['title']}", 
                style='Normal'
            )
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Add spacing between figures
            doc.add_paragraph()

        # Save document
        output_path = os.path.join(directory, "Generated_Report.docx")
        doc.save(output_path)
        return output_path
        
    except Exception as e:
        print(f"\n❌ Critical error generating report: {str(e)}")
        return None

def main():
    print("\n" + "="*40)
    print(" Lab Report Generator".ljust(39) + "=")
    print("="*40 + "\n")
    
    # Get directory path
    while True:
        directory = input("Enter full path to working directory: ").strip()
        if os.path.isdir(directory):
            break
        print("Invalid directory! Try again.")
    
    # Find template
    template_path = find_word_template(directory)
    if not template_path:
        sys.exit(1)
    
    # Process images
    figure_data = process_images(directory)
    if not figure_data:
        sys.exit(1)
    
    # Generate report
    report_path = generate_report(directory, template_path, figure_data)
    
    if report_path:
        print("\n" + "="*40)
        print(f"✅ Report successfully generated at:\n{report_path}")
        print("="*40 + "\n")
    else:
        print("\n❌ Failed to generate report")

if __name__ == "__main__":
    main()
