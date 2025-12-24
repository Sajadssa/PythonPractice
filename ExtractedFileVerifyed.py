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

# تنظیمات
source_directory = r"D:\Sepher_Pasargad\works\reports"
destination_directory = r"D:\Sepher_Pasargad\works\reports\All_exteact"

# اجرای برنامه
if __name__ == "__main__":
    print("شروع استخراج فایل‌ها...")
    print(f"مسیر مبدا: {source_directory}")
    print(f"مسیر مقصد: {destination_directory}\n")
    
    if os.path.exists(source_directory):
        extract_all_files(source_directory, destination_directory)
        print("\nعملیات با موفقیت انجام شد!")
    else:
        print(f"خطا: مسیر مبدا وجود ندارد: {source_directory}")