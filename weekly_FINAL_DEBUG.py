#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

try:
    import pdfplumber
except:
    pdfplumber = None

try:
    from docx import Document
except:
    Document = None
    print("CRITICAL: python-docx not installed!")
    print("Run: pip install python-docx")
    exit(1)


def extract_from_docx(docx_path):
    """Extract Sequence Number and Revision from Word document"""
    try:
        doc = Document(docx_path)
        
        # Extract ALL text from tables and paragraphs
        all_text = []
        
        # Get text from tables
        for table_idx, table in enumerate(doc.tables):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    if cell_text:
                        all_text.append(cell_text)
        
        # Get text from paragraphs
        for para in doc.paragraphs:
            para_text = para.text.strip()
            if para_text:
                all_text.append(para_text)
        
        # Combine all text
        full_text = " ".join(all_text)
        
        if not full_text:
            print(f"    ✗ No text extracted from Word file")
            return None, None, None
        
        print(f"    ✓ Extracted {len(full_text)} characters")
        
        # Show first 300 chars for debugging
        print(f"    Preview: {full_text[:300]}...")
        
        # Extract Sequence Number
        seq_num = None
        
        # Try different patterns
        patterns = [
            r'Sequence\s+Number\s*(\d+)',           # "Sequence Number 0005"
            r'SequenceNumber\s*(\d+)',              # "SequenceNumber0005"
            r'Number\s+(\d+)',                      # "Number 0005"
            r'REWK\s+(\d+)',                        # "REWK 0005"
            r'(\d{4})',                             # Any 4 digits
        ]
        
        for i, pattern in enumerate(patterns):
            matches = re.findall(pattern, full_text, re.I)
            if matches:
                for match in matches:
                    # Skip years
                    if match not in ['2024', '2025', '2026', '2023']:
                        seq_num = match.zfill(4)
                        print(f"    ✓ Found Sequence: {seq_num} (pattern {i+1})")
                        break
                if seq_num:
                    break
        
        if not seq_num:
            print(f"    ✗ Sequence Number NOT FOUND")
        
        # Extract Revision
        revision = None
        
        patterns = [
            r'Revision\s+(G\d+)',                   # "Revision G00"
            r'(G\d{2,3})',                          # "G00" or "G000"
        ]
        
        for i, pattern in enumerate(patterns):
            match = re.search(pattern, full_text, re.I)
            if match:
                revision = match.group(1).upper()
                print(f"    ✓ Found Revision: {revision} (pattern {i+1})")
                break
        
        if not revision:
            print(f"    ✗ Revision NOT FOUND")
        
        # Extract Date
        date_str = None
        date_match = re.search(r'(\d{1,2}[-/][A-Za-z]{3}[-/]\d{4})', full_text)
        if date_match:
            date_str = date_match.group(1)
            print(f"    ✓ Found Date: {date_str}")
        
        return seq_num, revision, date_str
        
    except Exception as e:
        print(f"    ✗ ERROR: {e}")
        import traceback
        traceback.print_exc()
        return None, None, None


def extract_from_pdf(pdf_path):
    """Extract Sequence Number and Revision from PDF"""
    if not pdfplumber:
        return None, None, None
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text()
            
            if not text:
                return None, None, None
            
            # Sequence Number
            seq = re.search(r'(\d{4})', text)
            seq_num = seq.group(1).zfill(4) if seq else None
            
            # Revision
            rev = re.search(r'(G\d{2,3})', text, re.I)
            revision = rev.group(1).upper() if rev else None
            
            # Date
            date = re.search(r'(\d{1,2}[-/][A-Za-z]{3}[-/]\d{4})', text)
            date_str = date.group(1) if date else None
            
            return seq_num, revision, date_str
    except:
        return None, None, None


def process_directory(directory_path):
    """Process all Word and PDF files in directory"""
    directory = Path(directory_path)
    
    if not directory.exists():
        print(f"\n✗ ERROR: Directory not found!")
        print(f"  Path: {directory_path}")
        return []
    
    # Get all files
    word_files = list(directory.glob("*.docx")) + list(directory.glob("*.doc"))
    pdf_files = list(directory.glob("*.pdf"))
    all_files = word_files + pdf_files
    
    print(f"\n{'='*70}")
    print(f"FOUND FILES:")
    print(f"  Word files: {len(word_files)}")
    print(f"  PDF files:  {len(pdf_files)}")
    print(f"  Total:      {len(all_files)}")
    print(f"{'='*70}\n")
    
    results = []
    failed = []
    
    for idx, file_path in enumerate(all_files, 1):
        # Skip temp files
        if file_path.name.startswith('~$'):
            print(f"{idx}. SKIP (temp): {file_path.name}\n")
            continue
        
        print(f"{idx}. {file_path.name}")
        print(f"    Type: {file_path.suffix}")
        
        # Extract data
        if file_path.suffix.lower() in ['.docx', '.doc']:
            seq, rev, date = extract_from_docx(file_path)
        elif file_path.suffix.lower() == '.pdf':
            seq, rev, date = extract_from_pdf(file_path)
        else:
            print(f"    ✗ Unknown file type\n")
            continue
        
        # Check results
        if seq and rev:
            new_name = f"SJSC-GGNRSP-EPWC-REWK-{seq}-{rev}{file_path.suffix}"
            results.append({
                'original_name': file_path.name,
                'new_name': new_name,
                'date': date or 'N/A',
                'old_path': file_path
            })
            print(f"    ✓ SUCCESS: {new_name}\n")
        else:
            failed.append(file_path.name)
            print(f"    ✗ FAILED: Missing Seq={seq} or Rev={rev}\n")
    
    print(f"{'='*70}")
    print(f"RESULTS: {len(results)} successful, {len(failed)} failed")
    print(f"{'='*70}\n")
    
    return results


def create_excel_report(results, output_path):
    """Create Excel report with renamed files list"""
    print(f"Creating Excel report...")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Weekly Reports"
    
    # Headers
    ws.append(['Original Filename', 'New Filename', 'Date'])
    
    # Style headers
    for cell in ws[1]:
        cell.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        cell.font = Font(bold=True, color='FFFFFF', size=12)
        cell.alignment = Alignment(horizontal='center')
    
    # Add data
    for item in results:
        ws.append([item['original_name'], item['new_name'], item['date']])
    
    # Column widths
    ws.column_dimensions['A'].width = 60
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 20
    
    # Save
    wb.save(output_path)
    print(f"✓ Excel saved: {output_path}\n")


def rename_files(results):
    """Rename all files"""
    print(f"{'='*70}")
    print(f"RENAMING FILES")
    print(f"{'='*70}\n")
    
    success = 0
    
    for item in results:
        old_path = item['old_path']
        new_path = old_path.parent / item['new_name']
        
        if new_path.exists():
            print(f"✗ SKIP (exists): {item['new_name']}")
            continue
        
        try:
            old_path.rename(new_path)
            print(f"✓ {item['original_name']}")
            print(f"  → {item['new_name']}")
            success += 1
        except Exception as e:
            print(f"✗ FAILED: {item['original_name']}")
            print(f"  Error: {e}")
    
    print(f"\n{'='*70}")
    print(f"✓ Successfully renamed {success}/{len(results)} files")
    print(f"{'='*70}\n")


def main():
    """Main function"""
    
    # PATHS - CHANGE THESE
    DIRECTORY = r"D:\Sepher_Pasargad\works\Maintenace\Maintenance Report\All_Extracted\weekly"
    EXCEL_OUTPUT = r"D:\Sepher_Pasargad\works\Maintenace\Maintenance Report\All_Extracted\weekly_reports.xlsx"
    
    print("="*70)
    print("WEEKLY REPORTS RENAMING TOOL")
    print("Output Format: SJSC-GGNRSP-EPWC-REWK-####-G##")
    print("="*70)
    
    # Check if python-docx is available
    if not Document:
        print("\n✗ CRITICAL ERROR: python-docx not installed!")
        print("  Run: pip install python-docx")
        return
    
    # Process files
    results = process_directory(DIRECTORY)
    
    if not results:
        print("\n✗ No files were successfully processed!")
        print("\nPossible reasons:")
        print("1. Files don't contain Sequence Number in header table")
        print("2. Files don't contain Revision (G00, G01, etc.)")
        print("3. Directory path is wrong")
        return
    
    # Create Excel report
    create_excel_report(results, EXCEL_OUTPUT)
    
    # Ask to rename
    print(f"{len(results)} files are ready to rename")
    answer = input("\nRename files now? (yes/no): ").strip().lower()
    
    if answer in ['yes', 'y']:
        rename_files(results)
        print("✓ DONE!")
    else:
        print("\n✓ Files NOT renamed. Check Excel report first.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n✗ Interrupted by user")
    except Exception as e:
        print(f"\n\n✗ UNEXPECTED ERROR: {e}")
        import traceback
        traceback.print_exc()
