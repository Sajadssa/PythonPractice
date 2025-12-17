import os
import re
import pandas as pd
from datetime import datetime
from PyPDF2 import PdfMerger, PdfReader
import shutil

class PdfMergerProcessor:
    def __init__(self, directory_path):
        self.directory_path = directory_path
        self.results = []
    
    def extract_info_from_pdf_content(self, pdf_path):
        """
        استخراج Doc No و Date از محتوای داخل فایل PDF
        """
        doc_no = None
        date = None
        number = None
        rev = None
        
        try:
            # خواندن محتوای PDF
            reader = PdfReader(pdf_path)
            text = ""
            
            # خواندن چند صفحه اول (معمولاً Doc No در صفحه اول است)
            for page_num in range(min(3, len(reader.pages))):
                page = reader.pages[page_num]
                text += page.extract_text() + " "
            
            print(f"  🔍 متن استخراج شده: {text[:200]}...")
            
            # جستجوی Doc No با فرمت: SJSC-GGNRSP-PDPE-REDH/REDL-XXXX-GXX
            doc_pattern = r'(SJSC-GGNRSP-[A-Z]+-[A-Z]+-(\d{4})-(G\d{2}))'
            match = re.search(doc_pattern, text, re.IGNORECASE)
            
            if match:
                doc_no = match.group(1)
                number = match.group(2)  # 4 رقم وسط (مثل 0388)
                rev = match.group(3)      # Gxx (مثل G00)
                print(f"  ✅ Doc No پیدا شد: {doc_no}")
                print(f"  ✅ Number: {number}, Rev: {rev}")
            else:
                # تلاش با الگوی دیگر
                alt_pattern = r'Doc\s*No\.?\s*:?\s*([A-Z0-9\-]+)'
                match2 = re.search(alt_pattern, text, re.IGNORECASE)
                if match2:
                    doc_no = match2.group(1)
                    # استخراج number و rev از doc_no
                    parts = doc_no.split('-')
                    for i, part in enumerate(parts):
                        if re.match(r'\d{4}', part):
                            number = part
                            if i + 1 < len(parts):
                                rev_part = parts[i + 1]
                                if re.match(r'G?\d{2}', rev_part, re.IGNORECASE):
                                    rev = 'G' + re.sub(r'[^0-9]', '', rev_part).zfill(2)
                            break
            
            # جستجوی Date
            date_patterns = [
                r'Date\s*:?\s*(\d{1,2}-[A-Za-z]{3}-\d{4})',
                r'Date\s*:?\s*(\d{4}-\d{2}-\d{2})',
                r'(\d{4}-\d{2}-\d{2})',
                r'(\d{1,2}/\d{1,2}/\d{4})',
            ]
            
            for pattern in date_patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    date = match.group(1).strip()
                    print(f"  ✅ Date پیدا شد: {date}")
                    break
            
        except Exception as e:
            print(f"  ⚠️  خطا در خواندن PDF: {e}")
        
        return {
            'doc_no': doc_no,
            'number': number,
            'rev': rev,
            'date': date
        }
    
    def identify_file_type(self, filename):
        """
        شناسایی نوع فایل (Heavy یا Light)
        """
        filename_lower = filename.lower()
        if 'heavy' in filename_lower:
            return 'Heavy'
        elif 'light' in filename_lower:
            return 'Light'
        else:
            return 'Unknown'
    
    def process_pdf_files(self):
        """
        پردازش و ادغام تمام فایل‌های PDF
        """
        # دریافت لیست فایل‌های PDF
        pdf_files = [f for f in os.listdir(self.directory_path) 
                     if f.lower().endswith('.pdf')
                     and not f.startswith('SJSC-GGNRSP-MOWP-REDA')]  # فایل‌های ادغام شده قبلی
        
        if not pdf_files:
            print("❌ هیچ فایل PDF یافت نشد!")
            return
        
        print(f"📁 تعداد {len(pdf_files)} فایل PDF یافت شد.\n")
        print("="*80)
        
        # مرحله 1: خواندن و استخراج اطلاعات از تمام فایل‌های PDF
        print("\n🔄 مرحله 1: خواندن محتوای PDF و استخراج اطلاعات...\n")
        
        file_info_dict = {}  # {number: {'heavy': file_info, 'light': file_info}}
        
        for idx, pdf_file in enumerate(pdf_files, 1):
            print(f"[{idx}/{len(pdf_files)}] 📄 {pdf_file}")
            
            pdf_path = os.path.join(self.directory_path, pdf_file)
            
            # شناسایی نوع فایل از نام
            file_type = self.identify_file_type(pdf_file)
            print(f"  📋 نوع فایل: {file_type}")
            
            # استخراج اطلاعات از محتوای PDF
            info = self.extract_info_from_pdf_content(pdf_path)
            
            if not info['number'] or not info['rev']:
                print(f"  ⚠️  اطلاعات ناقص - Number: {info['number']}, Rev: {info['rev']}\n")
                self.results.append({
                    'نام فایل PDF': pdf_file,
                    'نوع گزارش': file_type,
                    'Doc No': info['doc_no'] or 'نامشخص',
                    'Number': info['number'] or 'نامشخص',
                    'Rev': info['rev'] or 'نامشخص',
                    'تاریخ': info['date'] or 'نامشخص',
                    'نام فایل نهایی': 'پردازش نشد',
                    'وضعیت': 'ناموفق - اطلاعات ناقص'
                })
                continue
            
            # دسته‌بندی بر اساس Number
            number = info['number']
            
            if number not in file_info_dict:
                file_info_dict[number] = {
                    'heavy': None,
                    'light': None,
                    'rev': info['rev'],
                    'date': info['date']
                }
            
            # ذخیره اطلاعات فایل
            file_data = {
                'original_name': pdf_file,
                'pdf_path': pdf_path,
                'info': info
            }
            
            if file_type == 'Heavy':
                if file_info_dict[number]['heavy'] is not None:
                    print(f"  ⚠️  هشدار: قبلاً یک فایل Heavy با Number={number} وجود دارد!")
                file_info_dict[number]['heavy'] = file_data
            elif file_type == 'Light':
                if file_info_dict[number]['light'] is not None:
                    print(f"  ⚠️  هشدار: قبلاً یک فایل Light با Number={number} وجود دارد!")
                file_info_dict[number]['light'] = file_data
            
            print(f"  ✅ اطلاعات استخراج شد و دسته‌بندی شد (Number: {number})\n")
        
        # مرحله 2: ادغام فایل‌های Heavy و Light با Number یکسان
        print("\n" + "="*80)
        print("🔄 مرحله 2: ادغام فایل‌های Heavy و Light با Number یکسان...\n")
        
        for number, group in file_info_dict.items():
            print(f"📊 Number: {number}")
            print(f"   Rev: {group['rev']}")
            print(f"   تاریخ: {group['date'] or 'نامشخص'}")
            
            has_heavy = group['heavy'] is not None
            has_light = group['light'] is not None
            
            if has_heavy:
                print(f"   ✅ Heavy: {group['heavy']['original_name']}")
            else:
                print(f"   ❌ Heavy: وجود ندارد")
            
            if has_light:
                print(f"   ✅ Light: {group['light']['original_name']}")
            else:
                print(f"   ❌ Light: وجود ندارد")
            
            # نام فایل نهایی
            final_pdf_name = f"SJSC-GGNRSP-MOWP-REDA-{number}-{group['rev']}.pdf"
            final_pdf_path = os.path.join(self.directory_path, final_pdf_name)
            
            try:
                if has_heavy and has_light:
                    # ادغام Heavy + Light
                    print(f"   🔗 ادغام Heavy + Light...")
                    merger = PdfMerger()
                    
                    # ترتیب: Heavy اول، سپس Light
                    merger.append(group['heavy']['pdf_path'])
                    merger.append(group['light']['pdf_path'])
                    
                    merger.write(final_pdf_path)
                    merger.close()
                    
                    print(f"   ✅ ادغام موفق: {final_pdf_name}\n")
                    
                    # ثبت نتیجه
                    self.results.append({
                        'نام فایل PDF': f"{group['heavy']['original_name']} + {group['light']['original_name']}",
                        'نوع گزارش': 'Heavy + Light',
                        'Doc No': group['heavy']['info']['doc_no'],
                        'Number': number,
                        'Rev': group['rev'],
                        'تاریخ': group['date'] or 'نامشخص',
                        'نام فایل نهایی': final_pdf_name,
                        'وضعیت': 'موفق - ادغام شده'
                    })
                    
                elif has_heavy or has_light:
                    # فقط یکی از دو فایل
                    source_data = group['heavy'] if has_heavy else group['light']
                    file_type_name = 'Heavy' if has_heavy else 'Light'
                    
                    print(f"   📋 فقط {file_type_name} موجود است - کپی...")
                    shutil.copy2(source_data['pdf_path'], final_pdf_path)
                    
                    print(f"   ✅ کپی موفق: {final_pdf_name}\n")
                    
                    self.results.append({
                        'نام فایل PDF': source_data['original_name'],
                        'نوع گزارش': file_type_name,
                        'Doc No': source_data['info']['doc_no'],
                        'Number': number,
                        'Rev': group['rev'],
                        'تاریخ': group['date'] or 'نامشخص',
                        'نام فایل نهایی': final_pdf_name,
                        'وضعیت': f'موفق - فقط {file_type_name}'
                    })
                    
            except Exception as e:
                print(f"   ❌ خطا در پردازش: {e}\n")
                
                if has_heavy:
                    self.results.append({
                        'نام فایل PDF': group['heavy']['original_name'],
                        'نوع گزارش': 'Heavy',
                        'Doc No': group['heavy']['info']['doc_no'],
                        'Number': number,
                        'Rev': group['rev'],
                        'تاریخ': group['date'] or 'نامشخص',
                        'نام فایل نهایی': 'ناموفق',
                        'وضعیت': f'ناموفق - خطا: {str(e)[:50]}'
                    })
                
                if has_light:
                    self.results.append({
                        'نام فایل PDF': group['light']['original_name'],
                        'نوع گزارش': 'Light',
                        'Doc No': group['light']['info']['doc_no'],
                        'Number': number,
                        'Rev': group['rev'],
                        'تاریخ': group['date'] or 'نامشخص',
                        'نام فایل نهایی': 'ناموفق',
                        'وضعیت': f'ناموفق - خطا: {str(e)[:50]}'
                    })
    
    def save_report(self):
        """
        ذخیره گزارش در Excel
        """
        if not self.results:
            print("\n⚠️  هیچ نتیجه‌ای برای ذخیره وجود ندارد!")
            return
        
        df = pd.DataFrame(self.results)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"PDF_Merge_Report_{timestamp}.xlsx"
        excel_path = os.path.join(self.directory_path, excel_filename)
        
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
                adjusted_width = min(max_length + 3, 60)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # خلاصه
            successful = len([r for r in self.results if 'موفق' in r['وضعیت']])
            failed = len([r for r in self.results if 'ناموفق' in r['وضعیت']])
            merged_count = len([r for r in self.results if 'ادغام شده' in r['وضعیت']])
            
            summary_data = {
                'شرح': [
                    'تعداد کل فایل‌های پردازش شده',
                    'تعداد ادغام موفق (Heavy + Light)',
                    'تعداد فایل واحد',
                    'تعداد ناموفق',
                    'درصد موفقیت',
                    'تاریخ و زمان'
                ],
                'مقدار': [
                    len(self.results),
                    merged_count,
                    successful - merged_count,
                    failed,
                    f"{(successful/len(self.results)*100):.1f}%" if self.results else "0%",
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ]
            }
            
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='خلاصه', index=False)
            
            worksheet_summary = writer.sheets['خلاصه']
            worksheet_summary.column_dimensions['A'].width = 40
            worksheet_summary.column_dimensions['B'].width = 35
        
        print(f"\n📊 گزارش Excel ذخیره شد:")
        print(f"   {excel_path}")
        print(f"\n📈 خلاصه نتایج:")
        print(f"   ✅ کل موفق: {successful}")
        print(f"   🔗 ادغام شده (Heavy + Light): {merged_count}")
        print(f"   ❌ ناموفق: {failed}")
        
        return excel_path


def main():
    folder_path = r"D:\Sepher_Pasargad\works\Production\Daily_Acceptance"
    
    print("🚀 شروع پردازش و ادغام فایل‌های PDF...")
    print("="*80)
    print("📋 توضیحات:")
    print("   این برنامه محتوای فایل‌های PDF را می‌خواند")
    print("   و فایل‌های Heavy و Light با شماره سریال یکسان را")
    print("   در یک فایل ادغام می‌کند.")
    print("   مثال:")
    print("   - Heavy Daily Production Report-NIOC-2024-10-01 (Doc No: ...REDH-0388-G00)")
    print("   - Light Daily Production Report-NIOC-2024-10-01 (Doc No: ...REDL-0388-G00)")
    print("   → SJSC-GGNRSP-MOWP-REDA-0388-G00.pdf")
    print("="*80)
    print("\n📋 مراحل:")
    print("   1️⃣  خواندن محتوای تمام فایل‌های PDF")
    print("   2️⃣  استخراج Doc No و Number از داخل PDF")
    print("   3️⃣  دسته‌بندی Heavy و Light بر اساس Number یکسان")
    print("   4️⃣  ادغام فایل‌ها (Heavy اول، سپس Light)")
    print("   5️⃣  ذخیره با نام استاندارد")
    print("   6️⃣  ایجاد گزارش Excel")
    print("="*80)
    
    processor = PdfMergerProcessor(folder_path)
    
    # پردازش فایل‌ها
    processor.process_pdf_files()
    
    # ذخیره گزارش
    print("\n" + "="*80)
    processor.save_report()
    
    print("\n✨ پردازش کامل شد!")
    print("="*80)


if __name__ == "__main__":
    main()
