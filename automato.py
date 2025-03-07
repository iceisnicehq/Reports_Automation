import os
import glob
import csv
from datetime import datetime
from PIL import Image
from docx import Document
from docx.shared import Inches

# --------------------------
# Core Functions
# --------------------------

def get_creation_time(file_path):
    stat = os.stat(file_path)
    try:
        return stat.st_birthtime  # macOS
    except AttributeError:
        return stat.st_ctime  # Windows

def select_directory():
    print("\n🖼️  Welcome to Lab Report Automation!")
    print("1. Please navigate to your experiment directory in terminal")
    print("2. This directory should contain all your screenshots")
    print("3. Press Enter to use current directory or type path\n")
    
    while True:
        default_dir = os.getcwd()
        user_input = input(f"Enter directory [Default: '{default_dir}']: ").strip()
        target_dir = user_input or default_dir
        
        if not os.path.exists(target_dir):
            print(f"❌ Error: Directory '{target_dir}' doesn't exist!")
            continue
            
        if not glob.glob(os.path.join(target_dir, "*.*")):
            print(f"❌ Error: Directory is empty!")
            continue
            
        return target_dir

def process_screenshots(target_dir):
    image_files = sorted(
        [f for f in glob.glob(os.path.join(target_dir, "*.*")) 
         if f.lower().endswith(('.png', '.jpg', '.jpeg'))],
        key=lambda x: get_creation_time(x)
    )
    
    experiments = []
    
    for idx, file_path in enumerate(image_files, 1):
        print(f"\n📄 Processing file {idx}/{len(image_files)}")
        print(f"📍 Location: {file_path}")
        
        # Show image to user
        Image.open(file_path).show()
        
        # Get user inputs
        new_name = input("📝 New filename (e.g., exp1.png): ").strip()
        title = input("🔬 Experiment title: ").strip()
        objective = input("🎯 Objective: ").strip()
        procedure = input("📋 Procedure (steps separated by |): ").strip()
        description = input("📷 Screenshot description: ").strip()
        
        # Rename file
        new_path = os.path.join(target_dir, new_name)
        os.rename(file_path, new_path)
        
        experiments.append({
            "title": title,
            "new_name": new_name,
            "objective": objective,
            "procedure": procedure.replace('|', '\n'),
            "description": description,
            "timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })
    
    return experiments

def create_csv(target_dir, experiments):
    csv_path = os.path.join(target_dir, "experiments_metadata.csv")
    with open(csv_path, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=experiments[0].keys())
        writer.writeheader()
        writer.writerows(experiments)
    return csv_path

def generate_reports(target_dir, csv_path):
    with open(csv_path, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            doc = Document()
            
            # Add report content
            doc.add_heading(row['title'], level=1)
            
            doc.add_heading('Objective', level=2)
            doc.add_paragraph(row['objective'])
            
            doc.add_heading('Procedure', level=2)
            doc.add_paragraph(row['procedure'])
            
            doc.add_heading('Screenshot', level=2)
            img_path = os.path.join(target_dir, row['new_name'])
            doc.add_picture(img_path, width=Inches(5.5))
            
            doc.add_paragraph(row['description'])
            
            # Save document
            safe_title = ''.join(c if c.isalnum() else '_' for c in row['title'])
            doc.save(os.path.join(target_dir, f"Report_{safe_title}.docx"))

# --------------------------
# Main Execution
# --------------------------

def main():
    target_dir = select_directory()
    os.chdir(target_dir)
    
    print("\n🔧 Starting screenshot processing...")
    experiments = process_screenshots(target_dir)
    
    print("\n📊 Creating CSV metadata file...")
    csv_path = create_csv(target_dir, experiments)
    
    print("\n📄 Generating DOCX reports...")
    generate_reports(target_dir, csv_path)
    
    print(f"\n✅ All done! Results saved in: {target_dir}")

if __name__ == "__main__":
    main()
