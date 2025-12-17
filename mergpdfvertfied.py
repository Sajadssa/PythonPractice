import os
from PyPDF2 import PdfMerger, PdfReader
import pandas as pd
from datetime import datetime

def merge_pdf_files(directory_path):
    """
    ادغام فایل‌های PDF که پسوند _1 دارند با فایل اصلی بدون _1
    
    Args:
        directory_path: مسیر پوشه حاوی فایل‌های PDF
    """
    
    # بررسی وجود پوشه
    if not os.path.exists(directory_path):
        print(f"❌ پوشه '{directory_path}' یافت نشد!")
        return
    
    # دریافت لیست فایل‌های PDF با پسوند _1
    pdf_files = [f for f in os.listdir(directory_path) if f.endswith('_1.pdf')]
    
    if not pdf_files:
        print("⚠️  هیچ فایلی با پسوند '_1.pdf' یافت نشد!")
        return
    
    print(f"📁 تعداد {len(pdf_files)} فایل برای ادغام یافت شد.\n")
    
    # لیست برای ذخیره نتایج
    results = []
    merged_count = 0
    skipped_count = 0
    
    for idx, pdf_file in enumerate(pdf_files, 1):
        # نام فایل اصلی (بدون _1)
        original_file = pdf_file.replace('_1.pdf', '.pdf')
        
        # مسیر کامل فایل‌ها
        file_with_1 = os.path.join(directory_path, pdf_file)
        original_file_path = os.path.join(directory_path, original_file)
        
        # بررسی وجود فایل اصلی
        if not os.path.exists(original_file_path):
            print(f"[{idx}/{len(pdf_files)}] ⚠️  '{original_file}' یافت نشد - رد شد")
            skipped_count += 1
            
            results.append({
                'ردیف': idx,
                'نام فایل _1': pdf_file,
                'نام فایل اصلی': original_file,
                'وضعیت': 'ناموفق',
                'دلیل': 'فایل اصلی یافت نشد',
                'تعداد صفحات فایل _1': 'نامشخص',
                'تعداد صفحات فایل اصلی': 'نامشخص',
                'تعداد صفحات نهایی': 'نامشخص',
                'فایل _1 حذف شد': 'خیر'
            })
            continue
        
        try:
            # خواندن تعداد صفحات فایل‌ها قبل از ادغام
            try:
                reader_original = PdfReader(original_file_path)
                pages_original = len(reader_original.pages)
            except:
                pages_original = 'نامشخص'
            
            try:
                reader_1 = PdfReader(file_with_1)
                pages_1 = len(reader_1.pages)
            except:
                pages_1 = 'نامشخص'
            
            # ایجاد یک PdfMerger
            merger = PdfMerger()
            
            # اضافه کردن فایل اصلی
            merger.append(original_file_path)
            
            # اضافه کردن فایل با پسوند _1
            merger.append(file_with_1)
            
            # ذخیره فایل ادغام شده با نام موقت
            temp_file = os.path.join(directory_path, f"temp_{original_file}")
            merger.write(temp_file)
            merger.close()
            
            # خواندن تعداد صفحات فایل نهایی
            try:
                reader_final = PdfReader(temp_file)
                pages_final = len(reader_final.pages)
            except:
                pages_final = 'نامشخص'
            
            # جایگزینی فایل اصلی با فایل ادغام شده
            os.remove(original_file_path)
            os.rename(temp_file, original_file_path)
            
            # حذف فایل _1
            os.remove(file_with_1)
            
            print(f"[{idx}/{len(pdf_files)}] ✅ '{pdf_file}' به '{original_file}' اضافه شد (صفحات: {pages_1} + {pages_original} = {pages_final})")
            merged_count += 1
            
            results.append({
                'ردیف': idx,
                'نام فایل _1': pdf_file,
                'نام فایل اصلی': original_file,
                'وضعیت': 'موفق',
                'دلیل': 'ادغام و حذف انجام شد',
                'تعداد صفحات فایل _1': pages_1,
                'تعداد صفحات فایل اصلی': pages_original,
                'تعداد صفحات نهایی': pages_final,
                'فایل _1 حذف شد': 'بله'
            })
            
        except Exception as e:
            print(f"[{idx}/{len(pdf_files)}] ❌ خطا در پردازش '{pdf_file}': {str(e)}")
            skipped_count += 1
            
            results.append({
                'ردیف': idx,
                'نام فایل _1': pdf_file,
                'نام فایل اصلی': original_file,
                'وضعیت': 'ناموفق',
                'دلیل': f'خطا: {str(e)}',
                'تعداد صفحات فایل _1': 'نامشخص',
                'تعداد صفحات فایل اصلی': 'نامشخص',
                'تعداد صفحات نهایی': 'نامشخص',
                'فایل _1 حذف شد': 'خیر'
            })
    
    # ایجاد DataFrame و ذخیره در اکسل
    df = pd.DataFrame(results)
    
    # نام فایل اکسل با تاریخ و زمان
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"Merge_Report_{timestamp}.xlsx"
    excel_path = os.path.join(directory_path, excel_filename)
    
    # ذخیره با فرمت‌بندی
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='گزارش ادغام', index=False)
        
        # تنظیم عرض ستون‌ها
        worksheet = writer.sheets['گزارش ادغام']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 3, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # افزودن یک sheet برای خلاصه
        summary_data = {
            'شرح': [
                'تعداد کل فایل‌های _1',
                'تعداد ادغام موفق',
                'تعداد ناموفق',
                'درصد موفقیت',
                'تاریخ و زمان گزارش'
            ],
            'مقدار': [
                len(pdf_files),
                merged_count,
                skipped_count,
                f"{(merged_count/len(pdf_files)*100):.1f}%" if len(pdf_files) > 0 else "0%",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ]
        }
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='خلاصه', index=False)
        
        # تنظیم عرض ستون‌های sheet خلاصه
        worksheet_summary = writer.sheets['خلاصه']
        worksheet_summary.column_dimensions['A'].width = 30
        worksheet_summary.column_dimensions['B'].width = 25
    
    # چاپ خلاصه نتایج
    print(f"\n{'='*70}")
    print(f"📊 خلاصه نتایج:")
    print(f"  📁 تعداد کل فایل‌های _1: {len(pdf_files)}")
    print(f"  ✅ تعداد ادغام موفق: {merged_count}")
    print(f"  ❌ تعداد ناموفق: {skipped_count}")
    print(f"  📈 درصد موفقیت: {(merged_count/len(pdf_files)*100):.1f}%")
    print(f"\n  📄 گزارش اکسل ذخیره شد در:")
    print(f"     {excel_path}")
    print(f"{'='*70}")


if __name__ == "__main__":
    # مسیر پوشه
    folder_path = r"D:\Sepher_Pasargad\works\Production\Converted_Excel_to_PDF"
    
    print("🔄 شروع فرآیند ادغام فایل‌های PDF...\n")
    merge_pdf_files(folder_path)
    print("\n✨ فرآیند به پایان رسید!")