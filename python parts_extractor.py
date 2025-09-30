import pytesseract
from PIL import Image
import pandas as pd
import re
import cv2
import numpy as np

# اگر Tesseract در مسیر دیگری نصب شده، مسیر را مشخص کنید
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def preprocess_image(image_path):
    """پیش‌پردازش تصویر برای بهبود OCR"""
    img = cv2.imread(image_path)
    
    # تبدیل به grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # اعمال threshold برای بهبود کیفیت
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
    
    # حذف نویز
    denoised = cv2.fastNlMeansDenoising(thresh, None, 10, 7, 21)
    
    return denoised

def extract_table_data(image_path):
    """استخراج داده‌های جدول از تصویر"""
    
    # پیش‌پردازش تصویر
    processed_img = preprocess_image(image_path)
    
    # استخراج متن با OCR
    text = pytesseract.image_to_string(processed_img, lang='eng', config='--psm 6')
    
    # پردازش متن و استخراج اطلاعات
    lines = text.split('\n')
    
    data = []
    row_num = 1
    
    for line in lines:
        # فیلتر کردن خطوط خالی
        if not line.strip():
            continue
            
        # الگوی استخراج اطلاعات هر سطر
        # شماره | واحد | تعداد | شرح | شماره نقشه | جنس | توصیه سازنده | توصیه پیمانکار
        
        # جستجوی الگوی شماره و PCS
        if 'PCS' in line or 'pcs' in line.lower():
            parts = line.split()
            
            try:
                # استخراج اطلاعات اولیه
                unit = 'PCS.'
                
                # پیدا کردن اعداد در خط
                numbers = re.findall(r'\d+', line)
                
                if len(numbers) >= 2:
                    no = int(numbers[0])
                    qty = int(numbers[1])
                    
                    # استخراج شرح قطعات
                    desc_match = re.search(r'PCS\.\s*\d+\s+(.+?)(?:SIEC-|$)', line)
                    description = desc_match.group(1).strip() if desc_match else ''
                    
                    # استخراج شماره نقشه
                    ref_match = re.search(r'(SIEC-[A-Z0-9-]+)', line)
                    drawing_ref = ref_match.group(1) if ref_match else 'SIEC-DGNRAC-ENEL-DIWI-0006'
                    
                    # افزودن به دیتا
                    data.append({
                        'ردیف': row_num,
                        'شماره': no,
                        'واحد': unit,
                        'تعداد': qty,
                        'شرح قطعات': description,
                        'شماره نقشه/مرجع': drawing_ref,
                        'جنس': 'Multimaterial',
                        'توصیه سازنده (سال)': '',
                        'توصیه پیمانکار (سال)': ''
                    })
                    
                    row_num += 1
                    
            except (ValueError, IndexError):
                continue
    
    return data

def extract_with_opencv_table_detection(image_path):
    """استخراج داده‌ها با تشخیص ساختار جدول"""
    
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    
    # تشخیص خطوط افقی
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    detect_horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    
    # تشخیص خطوط عمودی
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
    detect_vertical = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)
    
    # ترکیب خطوط
    table_structure = cv2.add(detect_horizontal, detect_vertical)
    
    # پیدا کردن contourها
    contours, _ = cv2.findContours(table_structure, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    
    # مرتب‌سازی contourها بر اساس موقعیت y
    contours = sorted(contours, key=lambda c: cv2.boundingRect(c)[1])
    
    return contours

def main():
    """تابع اصلی برای استخراج و ذخیره داده‌ها"""
    
    # مسیر فایل تصویر
    image_path = 'parts_list_image.png'  # نام فایل تصویر خود را وارد کنید
    
    print("در حال استخراج داده‌ها از تصویر...")
    
    # استخراج داده‌ها
    extracted_data = extract_table_data(image_path)
    
    # اگر داده استخراج نشد، از داده‌های دستی استفاده کنید
    if not extracted_data:
        print("استخراج خودکار موفق نبود. استفاده از داده‌های دستی...")
        extracted_data = [
            {'ردیف': 1, 'شماره': 42, 'واحد': 'PCS.', 'تعداد': 42, 'شرح قطعات': 'Fuse, 10x38 mm, 2A', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 11, 'توصیه پیمانکار (سال)': 33},
            {'ردیف': 2, 'شماره': 2, 'واحد': 'PCS.', 'تعداد': 2, 'شرح قطعات': 'Power Fuse, 350A', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 3, 'شماره': 2, 'واحد': 'PCS.', 'تعداد': 2, 'شرح قطعات': 'MCB - 6A, 1P, for heater & lighting', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 4, 'شماره': 20, 'واحد': 'PCS.', 'تعداد': 20, 'شرح قطعات': 'MCB - 10A, 2P, for distribution', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 3, 'توصیه پیمانکار (سال)': 9},
            {'ردیف': 5, 'شماره': 3, 'واحد': 'PCS.', 'تعداد': 3, 'شرح قطعات': 'MCB - 16A, 2P, for distribution', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 6, 'شماره': 5, 'واحد': 'PCS.', 'تعداد': 5, 'شرح قطعات': 'MCB - 20A, 2P, for distribution', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 7, 'شماره': 12, 'واحد': 'PCS.', 'تعداد': 12, 'شرح قطعات': 'Rec. Thyristor (SCR), SKKT106', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 4, 'توصیه پیمانکار (سال)': 12},
            {'ردیف': 8, 'شماره': 4, 'واحد': 'PCS.', 'تعداد': 4, 'شرح قطعات': 'STS Thyristor (SCR), SKKT106', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 9, 'شماره': 2, 'واحد': 'PCS.', 'تعداد': 2, 'شرح قطعات': 'Block Diode Module', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 10, 'شماره': 2, 'واحد': 'PCS.', 'تعداد': 2, 'شرح قطعات': 'IGBT (Transistor), CM400', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 11, 'شماره': 4, 'واحد': 'PCS.', 'تعداد': 4, 'شرح قطعات': 'Rectifier Thyristor Driver PCB (control cards)', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 12, 'شماره': 2, 'واحد': 'PCS.', 'تعداد': 2, 'شرح قطعات': 'IGBT Driver PCB (control cards)', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 13, 'شماره': 2, 'واحد': 'PCS.', 'تعداد': 2, 'شرح قطعات': 'STS Thyristor Driver PCB (control cards)', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 14, 'شماره': 1, 'واحد': 'PCS.', 'تعداد': 1, 'شرح قطعات': 'Isolator switch (make before break), 100A', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 1, 'توصیه پیمانکار (سال)': 3},
            {'ردیف': 15, 'شماره': 352, 'واحد': 'PCS.', 'تعداد': 352, 'شرح قطعات': 'Batteries (SBM200, 176 Cells) (2 Banks)', 'شماره نقشه/مرجع': 'SIEC-DGNRAC-ENEL-DIWI-0006', 'جنس': 'Multimaterial', 'توصیه سازنده (سال)': 18, 'توصیه پیمانکار (سال)': 54}
        ]
    
    # ایجاد DataFrame
    df = pd.DataFrame(extracted_data)
    
    # ذخیره در فایل اکسل
    output_file = 'parts_list_extracted.xlsx'
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    print(f"✅ فایل اکسل با موفقیت ذخیره شد: {output_file}")
    print(f"📊 تعداد رکوردها: {len(df)}")
    print("\n🔍 نمایش 5 رکورد اول:")
    print(df.head())
    
    # ذخیره در فایل CSV نیز
    csv_file = 'parts_list_extracted.csv'
    df.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"\n✅ فایل CSV نیز ذخیره شد: {csv_file}")

if __name__ == "__main__":
    # نصب کتابخانه‌های مورد نیاز (در صورت نیاز):
    # pip install pytesseract opencv-python pandas openpyxl Pillow
    
    main()