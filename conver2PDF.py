import os
import shutil
from pathlib import Path

def extract_all_files(source_dir, destination_dir):
    """
    تمام فایل‌های موجود در پوشه مبدا و زیرپوشه‌های آن را استخراج کرده
    و در یک پوشه مقصد قرار می‌دهد (با مدیریت فایل‌های تکراری)
    """
    
    # ایجاد پوشه مقصد اگر وجود نداشته باشد
    Path(destination_dir).mkdir(parents=True, exist_ok=True)
    
    # شمارنده برای آمار
    copied_files = 0
    duplicate_files = 0
    
    # پیمایش تمام فایل‌ها در پوشه مبدا و زیرپوشه‌ها
    for root, dirs, files in os.walk(source_dir):
        for filename in files:
            source_file = os.path.join(root, filename)
            destination_file = os.path.join(destination_dir, filename)
            
            # بررسی وجود فایل در مقصد
            if os.path.exists(destination_file):
                duplicate_files += 1
                # افزودن شماره به نام فایل برای جلوگیری از بازنویسی
                name, ext = os.path.splitext(filename)
                counter = 1
                while os.path.exists(destination_file):
                    new_filename = f"{name}_{counter}{ext}"
                    destination_file = os.path.join(destination_dir, new_filename)
                    counter += 1
            
            try:
                # کپی فایل به مقصد
                shutil.copy2(source_file, destination_file)
                copied_files += 1
                print(f"کپی شد: {filename}")
            except Exception as e:
                print(f"خطا در کپی {filename}: {str(e)}")
    
    # نمایش گزارش نهایی
    print("\n" + "="*50)
    print(f"تعداد فایل‌های کپی شده: {copied_files}")
    print(f"تعداد فایل‌های تکراری (با نام جدید): {duplicate_files}")
    print("="*50)

def convert_excel_to_pdf(source_dir, pdf_output_dir):
    """
    فایل‌های اکسلی که نسخه PDF ندارند را به PDF تبدیل می‌کند
    """
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("\n⚠️ خطا: کتابخانه pywin32 نصب نیست!")
        print("لطفاً با دستور زیر نصب کنید:")
        print("pip install pywin32")
        return
    
    # ایجاد پوشه خروجی PDF
    Path(pdf_output_dir).mkdir(parents=True, exist_ok=True)
    
    # پسوندهای اکسل
    excel_extensions = {'.xlsx', '.xls', '.xlsm', '.xlsb'}
    
    # پیدا کردن تمام فایل‌های اکسل
    excel_files = []
    for file in os.listdir(source_dir):
        file_path = os.path.join(source_dir, file)
        if os.path.isfile(file_path):
            ext = os.path.splitext(file)[1].lower()
            if ext in excel_extensions:
                excel_files.append(file)
    
    # بررسی کدام فایل‌های اکسل نسخه PDF ندارند
    files_to_convert = []
    for excel_file in excel_files:
        name_without_ext = os.path.splitext(excel_file)[0]
        pdf_exists = False
        
        # بررسی وجود PDF با همان نام
        for file in os.listdir(source_dir):
            if os.path.splitext(file)[0] == name_without_ext and file.lower().endswith('.pdf'):
                pdf_exists = True
                break
        
        if not pdf_exists:
            files_to_convert.append(excel_file)
    
    if not files_to_convert:
        print("\n✓ همه فایل‌های اکسل دارای نسخه PDF هستند!")
        return
    
    print(f"\n📄 تعداد {len(files_to_convert)} فایل اکسل برای تبدیل به PDF یافت شد...")
    
    # تبدیل فایل‌ها
    pythoncom.CoInitialize()
    excel = None
    converted_count = 0
    failed_count = 0
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        for excel_file in files_to_convert:
            try:
                excel_path = os.path.join(source_dir, excel_file)
                name_without_ext = os.path.splitext(excel_file)[0]
                pdf_path = os.path.join(pdf_output_dir, f"{name_without_ext}.pdf")
                
                print(f"در حال تبدیل: {excel_file}")
                
                # باز کردن فایل اکسل
                wb = excel.Workbooks.Open(excel_path)
                
                # تبدیل به PDF (0 = xlTypePDF)
                wb.ExportAsFixedFormat(0, pdf_path)
                
                # بستن فایل
                wb.Close(False)
                
                converted_count += 1
                print(f"✓ تبدیل شد: {name_without_ext}.pdf")
                
            except Exception as e:
                failed_count += 1
                print(f"✗ خطا در تبدیل {excel_file}: {str(e)}")
        
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()
    
    # گزارش نهایی
    print("\n" + "="*50)
    print(f"تعداد فایل‌های تبدیل شده: {converted_count}")
    print(f"تعداد فایل‌های با خطا: {failed_count}")
    print("="*50)

# تنظیمات
source_directory = r"D:\Sepher_Pasargad\works\Production\Daily_Acceptance"
extracted_files_dir = r"D:\Sepher_Pasargad\works\Production\Daily_Acceptance"
pdf_output_dir = r"D:\Sepher_Pasargad\works\Production\Converted_Excel_to_PDF"

# اجرای برنامه
if __name__ == "__main__":
    print("=" * 60)
    print("مرحله 1: استخراج فایل‌ها از زیرپوشه‌ها")
    print("=" * 60)
    print(f"مسیر مبدا: {source_directory}")
    print(f"مسیر مقصد: {extracted_files_dir}\n")
    
    if os.path.exists(source_directory):
        extract_all_files(source_directory, extracted_files_dir)
        print("\n✓ مرحله 1 با موفقیت انجام شد!")
    else:
        print(f"خطا: مسیر مبدا وجود ندارد: {source_directory}")
    
    print("\n" + "=" * 60)
    print("مرحله 2: تبدیل فایل‌های اکسل به PDF")
    print("=" * 60)
    print(f"مسیر مبدا: {extracted_files_dir}")
    print(f"مسیر خروجی PDF: {pdf_output_dir}\n")
    
    if os.path.exists(extracted_files_dir):
        convert_excel_to_pdf(extracted_files_dir, pdf_output_dir)
        print("\n✓ مرحله 2 با موفقیت انجام شد!")
    else:
        print(f"خطا: پوشه {extracted_files_dir} وجود ندارد!")
    
    print("\n" + "=" * 60)
    print("✓ تمام عملیات‌ها با موفقیت انجام شد!")
    print("=" * 60)