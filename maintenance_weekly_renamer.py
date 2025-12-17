import os
import re
from PyPDF2 import PdfReader
from docx import Document

# Directory path
directory = r"D:\Sepher_Pasargad\works\Maintenace\Maintenance Report\All_Extracted\weekly"

# Function to extract info from PDF
def extract_from_pdf(file_path):
    with open(file_path, 'rb') as f:
        reader = PdfReader(f)
        first_page = reader.pages[0]
        text = first_page.extract_text()
        
        # Parse sequence number (assuming pattern like 'Sequence Number 00xx')
        seq_match = re.search(r'Sequence Number\s*(\d+)', text, re.IGNORECASE)
        sequence = seq_match.group(1) if seq_match else None
        
        # Parse revision (like 'Revision G00')
        rev_match = re.search(r'Revision\s*(G\d+)', text, re.IGNORECASE)
        revision = rev_match.group(1) if rev_match else None
        
        # Parse date from approval table (like 'Date xx-xxx-xxxx')
        date_match = re.search(r'Date\s*(\d{2}-\w{3}-\d{4})', text, re.IGNORECASE)
        date = date_match.group(1) if date_match else None
        
        return sequence, revision, date

# Function to extract info from Word
def extract_from_docx(file_path):
    doc = Document(file_path)
    text = '\n'.join([para.text for para in doc.paragraphs])
    
    # Similar regex as above
    seq_match = re.search(r'Sequence Number\s*(\d+)', text, re.IGNORECASE)
    sequence = seq_match.group(1) if seq_match else None
    
    rev_match = re.search(r'Revision\s*(G\d+)', text, re.IGNORECASE)
    revision = rev_match.group(1) if rev_match else None
    
    date_match = re.search(r'Date\s*(\d{2}-\w{3}-\d{4})', text, re.IGNORECASE)
    date = date_match.group(1) if date_match else None
    
    return sequence, revision, date

# List to hold report
report = []

# Loop through files
for filename in os.listdir(directory):
    if filename.lower().endswith(('.pdf', '.docx')):
        file_path = os.path.join(directory, filename)
        sequence, revision, date = None, None, None
        
        if filename.lower().endswith('.pdf'):
            sequence, revision, date = extract_from_pdf(file_path)
        elif filename.lower().endswith('.docx'):
            sequence, revision, date = extract_from_docx(file_path)
        
        if sequence and revision and date:
            # Assuming discipline is 'PDME' based on images (user said MADR, might be typo)
            new_name = f"SJSC-GGNRSP-MADR-REWK-{sequence}-{revision}"
            ext = os.path.splitext(filename)[1]
            new_file_path = os.path.join(directory, new_name + ext)
            
            # Rename file
            os.rename(file_path, new_file_path)
            
            # Add to report
            report.append({
                'old_name': filename,
                'new_name': new_name + ext,
                'date': date
            })
        else:
            print(f"Could not extract info from {filename}")

# Output report
print("Renaming Report:")
for item in report:
    print(f"Old: {item['old_name']} -> New: {item['new_name']} | Date: {item['date']}")