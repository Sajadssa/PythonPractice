import pandas as pd
import os
from pathlib import Path
import re
from datetime import datetime
import shutil
import PyPDF2

def extract_info_from_pdf(pdf_path):
    """
    استخراج Date, Ref No و Title از فایل PDF
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            # خواندن صفحه اول
            first_page = pdf_reader.pages[0]
            text = first_page.extract_text()
            
            # استخراج Date
            date_match = re.search(r'Date\s*[:：]?\s*(\d{1,2}[-/]\w{3}[-/]\d{4})', text, re.IGNORECASE)
            if not date_match:
                date_match = re.search(r'(\d{1,2}[-/]\w{3}[-/]\d{4})', text)
            
            date_str = date_match.group(1) if date_match else None
            
            # تبدیل تاریخ به فرمت استاندارد
            date_obj = None
            if date_str:
                try:
                    # تبدیل فرمت مثل "4-Oct-2023" به datetime
                    date_obj = datetime.strptime(date_str, '%d-%b-%Y')
                except:
                    try:
                        date_obj = datetime.strptime(date_str, '%d/%b/%Y')
                    except:
                        pass
            
            # استخراج Ref No
            ref_match = re.search(r'Ref\s*No\.?\s*[:：]?\s*(SJSC-[A-Z]+-[A-Z]+-[A-Z]+-\d+-G\d+)', text, re.IGNORECASE)
            ref_no = ref_match.group(1) if ref_match else None
            
            # استخراج Title
            title = None
            if 'Light Crude Wells Production Performance' in text:
                title = 'Light Crude Wells Production Performance'
            elif 'Production Engineering' in text:
                title = 'Production Engineering Report'
            else:
                # جستجوی عنوان بین Date و جدول اول
                title_match = re.search(r'DAILY PRODUCTION REPORT\s*\n\s*(.+?)(?:\n|Production Parameters)', text, re.IGNORECASE | re.DOTALL)
                if title_match:
                    title = title_match.group(1).strip()
            
            return {
                'date': date_obj,
                'date_str': date_obj.strftime('%m/%d/%Y') if date_obj else date_str,
                'ref_no': ref_no,
                'title': title
            }
    
    except Exception as e:
        print(f"   ⚠️ خطا در خواندن PDF: {str(e)}")
        return None

def process_pdf_files(pdf_folder):
    """
    پردازش فایل‌های PDF و استخراج اطلاعات
    """
    results = []
    
    # پیدا کردن تمام فایل‌های PDF
    pdf_files = list(Path(pdf_folder).glob('*.pdf'))
    
    if not pdf_files:
        print("⚠️ هیچ فایل PDF پیدا نشد!")
        return results
    
    print(f"📁 {len(pdf_files)} فایل PDF پیدا شد\n")
    
    # استخراج اطلاعات از هر فایل
    files_with_info = []
    for pdf_path in pdf_files:
        print(f"🔄 بررسی: {pdf_path.name}")
        
        info = extract_info_from_pdf(pdf_path)
        
        if info and info['date']:
            files_with_info.append((pdf_path, info))
            print(f"   ✅ Date: {info['date_str']} | Ref: {info['ref_no']}")
            print(f"      Title: {info['title']}")
        else:
            print(f"   ⚠️ نتوانستیم اطلاعات را استخراج کنیم")
    
    if not files_with_info:
        print("❌ نتوانستیم اطلاعات هیچ فایلی را استخراج کنیم!")
        return results
    
    # مرتب‌سازی بر اساس تاریخ (صعودی - از قدیمی به جدید)
    files_with_info.sort(key=lambda x: x[1]['date'])
    
    print("\n" + "="*80)
    print("🔄 تغییر نام فایل‌ها و استخراج اطلاعات...")
    print("="*80)
    
    # ایجاد پوشه برای فایل‌های تغییر نام داده شده
    renamed_folder = Path(pdf_folder) / "Renamed_Files"
    renamed_folder.mkdir(exist_ok=True)
    
    # پردازش فایل‌ها
    for idx, (old_path, info) in enumerate(files_with_info, start=1):
        # نام جدید
        new_name = f"SJSC-GGNRSP-MOCD-REDA-{idx:04d}-G00.pdf"
        new_path = renamed_folder / new_name
        
        # کپی فایل با نام جدید
        try:
            shutil.copy2(old_path, new_path)
            
            results.append({
                'Row': idx,
                'DATE': info['date_str'],
                'Ref no.': info['ref_no'] or 'N/A',
                'Title': info['title'] or 'N/A',
                'New_RefNo': f"SJSC-GGNRSP-MOCD-REDA-{idx:04d}-G00",
                'Original_File': old_path.name,
                'New_File': new_name
            })
            
            print(f"✅ [{idx:04d}] {info['date_str']}")
            print(f"         {old_path.name}")
            print(f"         ➜ {new_name}")
        
        except Exception as e:
            print(f"❌ [{idx:04d}] خطا: {str(e)}")
    
    return results

def main():
    """
    تابع اصلی
    """
    # مسیر پوشه حاوی فایل‌های PDF
    PDF_FOLDER = r"D:\Sepher_Pasargad\works\Production\DailyProductionReport-2023\2] Nov-2023"
    
    print("="*80)
    print("🚀 استخراج اطلاعات از گزارش‌های روزانه تولید")
    print("="*80)
    print(f"📂 مسیر پوشه: {PDF_FOLDER}\n")
    
    # بررسی وجود پوشه
    if not os.path.exists(PDF_FOLDER):
        print(f"❌ خطا: پوشه پیدا نشد!")
        print(f"مسیر: {PDF_FOLDER}")
        return
    
    # پردازش فایل‌ها
    results = process_pdf_files(PDF_FOLDER)
    
    if not results:
        print("\n❌ هیچ داده‌ای استخراج نشد!")
        return
    
    # ایجاد DataFrame
    df_output = pd.DataFrame(results)
    
    # فقط ستون‌های مورد نیاز برای خروجی نهایی
    df_final = df_output[['Row', 'DATE', 'Ref no.', 'Title']].copy()
    
    # ذخیره در Excel
    output_file = os.path.join(PDF_FOLDER, 'Production_Reports_Summary.xlsx')
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_final.to_excel(writer, sheet_name='Summary', index=False)
        
        # تنظیم عرض ستون‌ها
        worksheet = writer.sheets['Summary']
        worksheet.column_dimensions['A'].width = 8   # Row
        worksheet.column_dimensions['B'].width = 15  # DATE
        worksheet.column_dimensions['C'].width = 40  # Ref no.
        worksheet.column_dimensions['D'].width = 50  # Title
    
    # ذخیره جزئیات کامل
    details_file = os.path.join(PDF_FOLDER, 'Production_Reports_Details.xlsx')
    with pd.ExcelWriter(details_file, engine='openpyxl') as writer:
        df_output.to_excel(writer, sheet_name='Details', index=False)
        
        worksheet = writer.sheets['Details']
        for col_num, column in enumerate(df_output.columns, 1):
            worksheet.column_dimensions[chr(64 + col_num)].width = 40
    
    print("\n" + "="*80)
    print("✅ موفقیت! فایل‌های خروجی ایجاد شدند")
    print("="*80)
    print(f"📄 فایل خلاصه: Production_Reports_Summary.xlsx")
    print(f"📄 فایل جزئیات: Production_Reports_Details.xlsx")
    print(f"📂 فایل‌های تغییر نام داده شده: Renamed_Files/")
    print(f"📊 تعداد گزارش‌ها: {len(results)}")
    print("="*80)
    
    # نمایش نمونه داده‌ها
    print("\n📋 نمونه 10 رکورد اول:")
    print("-"*80)
    print(df_final.head(10).to_string(index=False))
    print("\n✨ کار تمام شد!")

if __name__ == "__main__":
    main()
