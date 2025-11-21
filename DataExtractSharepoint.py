import pandas as pd
import os
from pathlib import Path
import re
from datetime import datetime
import PyPDF2
import pdfplumber

def extract_text_from_pdf(pdf_path):
    """
    استخراج متن از فایل PDF با دو روش
    """
    text = ""
    
    # روش 1: استفاده از pdfplumber (بهترین روش)
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:3]:  # فقط 3 صفحه اول برای سرعت
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        
        if text.strip():
            return text
    except Exception as e:
        print(f"   ⚠️ خطا در pdfplumber: {str(e)}")
    
    # روش 2: استفاده از PyPDF2 (روش جایگزین)
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(min(3, len(pdf_reader.pages))):
                page = pdf_reader.pages[page_num]
                text += page.extract_text() + "\n"
    except Exception as e:
        print(f"   ⚠️ خطا در PyPDF2: {str(e)}")
    
    return text

def detect_report_type(pdf_text, file_name):
    """
    تشخیص نوع گزارش از محتوای PDF
    """
    pdf_text_lower = pdf_text.lower()
    file_name_lower = file_name.lower()
    
    # بررسی کلیدواژه‌ها
    if "weekly" in pdf_text_lower or "weekly" in file_name_lower:
        if "production engineering" in pdf_text_lower:
            return "Weekly Production Engineering Report"
        else:
            return "Weekly Production Report"
    
    elif "daily" in pdf_text_lower or "daily" in file_name_lower:
        if "light crude" in pdf_text_lower:
            return "Daily - Light Crude Wells Production"
        elif "heavy crude" in pdf_text_lower:
            return "Daily - Heavy Crude Wells Production"
        elif "crude" in pdf_text_lower:
            return "Daily - Crude Production Report"
        else:
            return "Daily Production Report"
    
    elif "production engineering" in pdf_text_lower:
        return "Production Engineering Report"
    
    elif "production report" in pdf_text_lower:
        return "Production Report"
    
    else:
        # اگر نتوانستیم تشخیص دهیم، از نام فایل استفاده می‌کنیم
        if "daily" in file_name_lower:
            return "Daily Production Report"
        else:
            return "Production Report"

def extract_date_from_text(pdf_text, file_name):
    """
    استخراج تاریخ از متن PDF یا نام فایل
    """
    # روش 1: از نام فایل (فرمت: YYYYMMDD-Daily Production Report.pdf)
    date_match = re.search(r'(\d{4})(\d{2})(\d{2})', file_name)
    if date_match:
        year, month, day = date_match.groups()
        try:
            date_obj = datetime(int(year), int(month), int(day))
            return date_obj.strftime('%m/%d/%Y')
        except:
            pass
    
    # روش 2: جستجو در متن PDF
    # فرمت‌های متداول تاریخ
    date_patterns = [
        r'(\d{1,2})/(\d{1,2})/(\d{4})',  # MM/DD/YYYY
        r'(\d{4})-(\d{2})-(\d{2})',      # YYYY-MM-DD
        r'(\d{2})\.(\d{2})\.(\d{4})',    # DD.MM.YYYY
    ]
    
    for pattern in date_patterns:
        match = re.search(pattern, pdf_text)
        if match:
            try:
                if '/' in pattern:
                    return match.group(0)
                else:
                    # تبدیل به فرمت MM/DD/YYYY
                    groups = match.groups()
                    if len(groups[0]) == 4:  # YYYY-MM-DD
                        date_obj = datetime(int(groups[0]), int(groups[1]), int(groups[2]))
                    else:  # DD.MM.YYYY
                        date_obj = datetime(int(groups[2]), int(groups[1]), int(groups[0]))
                    return date_obj.strftime('%m/%d/%Y')
            except:
                continue
    
    return "N/A"

def extract_refno_from_file(file_name, date_str):
    """
    استخراج RefNo از نام فایل یا تاریخ
    """
    # روش 1: از نام فایل
    ref_match = re.search(r'(\d{8})', file_name)
    if ref_match:
        return ref_match.group(1)
    
    # روش 2: از تاریخ (تبدیل MM/DD/YYYY به YYYYMMDD)
    if date_str != "N/A":
        try:
            date_obj = datetime.strptime(date_str, '%m/%d/%Y')
            return date_obj.strftime('%Y%m%d')
        except:
            pass
    
    # روش 3: از نام فایل بدون پسوند
    return file_name.replace('.pdf', '').replace(' ', '_')

def extract_production_reports(pdf_folder):
    """
    استخراج اطلاعات از تمام فایل‌های PDF در پوشه ProductionReport
    """
    results = []
    
    # پیدا کردن تمام فایل‌های PDF
    pdf_files = list(Path(pdf_folder).glob('*.pdf'))
    
    if not pdf_files:
        print("⚠️ هیچ فایل PDF پیدا نشد!")
        print(f"📂 در مسیر: {pdf_folder}")
        return results
    
    print(f"📁 {len(pdf_files)} فایل PDF پیدا شد\n")
    
    for idx, pdf_path in enumerate(sorted(pdf_files), start=1):
        file_name = pdf_path.name
        print(f"🔄 پردازش ({idx}/{len(pdf_files)}): {file_name}")
        
        try:
            # استخراج متن از PDF
            pdf_text = extract_text_from_pdf(pdf_path)
            
            if not pdf_text.strip():
                print(f"   ⚠️ نتوانستیم متن را از PDF استخراج کنیم")
                pdf_text = ""
            
            # استخراج تاریخ
            date_str = extract_date_from_text(pdf_text, file_name)
            
            # استخراج RefNo
            ref_no = extract_refno_from_file(file_name, date_str)
            
            # تشخیص نوع گزارش
            report_type = detect_report_type(pdf_text, file_name)
            
            results.append({
                'Row': idx,
                'RefNo.': ref_no,
                'Date': date_str,
                'TypeofReport': report_type
            })
            
            print(f"   ✅ RefNo: {ref_no} | Date: {date_str} | Type: {report_type}")
            
        except Exception as e:
            print(f"   ❌ خطا: {str(e)}")
            results.append({
                'Row': idx,
                'RefNo.': file_name.replace('.pdf', ''),
                'Date': 'Error',
                'TypeofReport': 'Error Processing File'
            })
    
    return results

def main():
    """
    تابع اصلی
    """
    # ⚠️ مسیر پوشه حاوی فایل‌های PDF دانلود شده از ProductionReport
    PDF_FOLDER = r"D:\Sepher_Pasargad\works\Production\LightCrude"  # 👈 مسیر خود را اینجا وارد 
    print("🚀 استخراج اطلاعات از گزارش‌های تولید (ProductionReport - PDF)")
    print("="*80)
    print(f"📂 مسیر پوشه: {PDF_FOLDER}\n")
    
    # بررسی وجود پوشه
    if not os.path.exists(PDF_FOLDER):
        print(f"❌ خطا: پوشه پیدا نشد!")
        print(f"لطفاً مسیر را بررسی کنید: {PDF_FOLDER}")
        print("\n💡 مراحل:")
        print("1. به SharePoint بروید:")
        print("   https://extranet.pedc.ir/pogp/PRD/ProductionReport")
        print("2. فایل‌های PDF را انتخاب کنید (یا همه را با Ctrl+A)")
        print("3. Download کنید")
        print("4. در یک پوشه ذخیره کنید")
        print("5. مسیر آن پوشه را در کد بالا وارد کنید")
        return
    
    # استخراج اطلاعات
    results = extract_production_reports(PDF_FOLDER)
    
    if not results:
        print("\n❌ هیچ داده‌ای استخراج نشد!")
        return
    
    # ایجاد DataFrame
    df_output = pd.DataFrame(results)
    
    # مرتب‌سازی بر اساس RefNo
    df_output = df_output.sort_values('RefNo.')
    df_output['Row'] = range(1, len(df_output) + 1)
    
    # ذخیره در Excel
    output_file = os.path.join(PDF_FOLDER, 'ProductionReport_Summary.xlsx')
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_output.to_excel(writer, sheet_name='Summary', index=False)
        
        # تنظیم عرض ستون‌ها
        worksheet = writer.sheets['Summary']
        worksheet.column_dimensions['A'].width = 8   # Row
        worksheet.column_dimensions['B'].width = 15  # RefNo.
        worksheet.column_dimensions['C'].width = 15  # Date
        worksheet.column_dimensions['D'].width = 50  # TypeofReport
    
    print("\n" + "="*80)
    print("✅ موفقیت! فایل خروجی ایجاد شد")
    print("="*80)
    print(f"📄 نام فایل: ProductionReport_Summary.xlsx")
    print(f"📂 مسیر کامل: {output_file}")
    print(f"📊 تعداد گزارش‌ها: {len(results)}")
    print("="*80)
    
    # نمایش نمونه داده‌ها
    print("\n📋 نمونه 10 رکورد اول:")
    print("-"*80)
    print(df_output.head(10).to_string(index=False))
    
    # ذخیره گزارش جزئیات برای بررسی
    details_file = os.path.join(PDF_FOLDER, 'extraction_details.txt')
    with open(details_file, 'w', encoding='utf-8') as f:
        f.write("جزئیات استخراج:\n")
        f.write("="*80 + "\n\n")
        for _, row in df_output.iterrows():
            f.write(f"Row: {row['Row']}\n")
            f.write(f"RefNo: {row['RefNo.']}\n")
            f.write(f"Date: {row['Date']}\n")
            f.write(f"Type: {row['TypeofReport']}\n")
            f.write("-"*80 + "\n")
    
    print(f"\n📝 فایل جزئیات نیز ذخیره شد: extraction_details.txt")
    print("\n✨ کار تمام شد!")

if __name__ == "__main__":
    main()
