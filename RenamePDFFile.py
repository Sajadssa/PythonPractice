import os
from pathlib import Path
import re
from datetime import datetime

def extract_date_from_filename(file_name):
    """
    استخراج تاریخ از نام فایل
    """
    # فرمت: YYYYMMDD-Daily Production Report.pdf
    date_match = re.search(r'(\d{4})(\d{2})(\d{2})', file_name)
    if date_match:
        year, month, day = date_match.groups()
        try:
            return datetime(int(year), int(month), int(day))
        except:
            pass
    return None

def rename_pdf_files(pdf_folder):
    """
    تغییر نام فایل‌های PDF به فرمت SJSC-GGNRSP-MOWP-REDA-XXXX-G00
    """
    print("="*80)
    print("🔄 تغییر نام فایل‌های PDF")
    print("="*80)
    print(f"📂 مسیر پوشه: {pdf_folder}\n")
    
    # پیدا کردن تمام فایل‌های PDF
    pdf_files = list(Path(pdf_folder).glob('*.pdf'))
    
    if not pdf_files:
        print("❌ هیچ فایل PDF پیدا نشد!")
        return
    
    print(f"📁 {len(pdf_files)} فایل PDF پیدا شد\n")
    
    # استخراج تاریخ و مرتب‌سازی
    files_with_dates = []
    for pdf_path in pdf_files:
        date_obj = extract_date_from_filename(pdf_path.name)
        if date_obj:
            files_with_dates.append((pdf_path, date_obj))
        else:
            print(f"⚠️ نمی‌توانیم تاریخ را از این فایل استخراج کنیم: {pdf_path.name}")
    
    if not files_with_dates:
        print("❌ نتوانستیم تاریخ هیچ فایلی را استخراج کنیم!")
        return
    
    # مرتب‌سازی بر اساس تاریخ (صعودی - از قدیمی به جدید)
    files_with_dates.sort(key=lambda x: x[1])
    
    print("🔄 شروع تغییر نام فایل‌ها...")
    print("-"*80)
    
    # تغییر نام فایل‌ها
    renamed_count = 0
    for idx, (old_path, date_obj) in enumerate(files_with_dates, start=1):
        # نام جدید
        new_name = f"SJSC-GGNRSP-MOWP-REDA-{idx:04d}-G00.pdf"
        new_path = old_path.parent / new_name
        
        # بررسی اینکه فایل با نام جدید وجود ندارد
        if new_path.exists() and new_path != old_path:
            print(f"⚠️ [{idx:04d}] فایل با این نام قبلاً وجود دارد: {new_name}")
            continue
        
        try:
            old_path.rename(new_path)
            renamed_count += 1
            print(f"✅ [{idx:04d}] {date_obj.strftime('%Y/%m/%d')} | {old_path.name}")
            print(f"         ➜ {new_name}")
        except Exception as e:
            print(f"❌ [{idx:04d}] خطا در تغییر نام: {str(e)}")
    
    print("-"*80)
    print(f"\n✅ تعداد {renamed_count} فایل با موفقیت تغییر نام داده شد!")
    print("="*80)
    
    # نمایش لیست نهایی
    print("\n📋 لیست نهایی فایل‌ها:")
    print("-"*80)
    final_files = sorted(Path(pdf_folder).glob('SJSC-GGNRSP-MOWP-REDA-*.pdf'))
    for idx, file_path in enumerate(final_files, start=1):
        print(f"{idx:3d}. {file_path.name}")
    print("="*80)

def main():
    """
    تابع اصلی
    """
    # ⚠️ مسیر پوشه حاوی فایل‌های PDF
    PDF_FOLDER = r"D:\Sepher_Pasargad\works\DCC\ProductionReport"  # 👈 مسیر خود را اینجا وارد کنید
    
    # بررسی وجود پوشه
    if not os.path.exists(PDF_FOLDER):
        print(f"❌ خطا: پوشه پیدا نشد!")
        print(f"مسیر: {PDF_FOLDER}")
        return
    
    # تأییدیه از کاربر
    print("\n⚠️ هشدار: این عملیات نام تمام فایل‌های PDF را تغییر می‌دهد!")
    print("آیا مطمئن هستید؟ (y/n): ", end='')
    
    # برای اجرای خودکار، این خط را کامنت کنید
    # confirmation = input().lower()
    # if confirmation != 'y':
    #     print("❌ عملیات لغو شد.")
    #     return
    
    # تغییر نام فایل‌ها
    rename_pdf_files(PDF_FOLDER)
    
    print("\n✨ کار تمام شد!")

if __name__ == "__main__":
    main()

