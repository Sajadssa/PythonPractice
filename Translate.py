import fitz
from PIL import Image
import pytesseract
from deep_translator import GoogleTranslator
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display
import re
import time
import io

# تنظیمات
INPUT_PDF = "SP-CA-SE-PD-0051.pdf"
OUTPUT_PDF = "ترجمه_صفحات_14_34.pdf"
START_PAGE = 14
END_PAGE = 34
FONT_PATH = "BNazanin.ttf"

# ⚠️ مسیر Tesseract رو تنظیم کن
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# دیکشنری اصطلاحات
OIL_GAS_TERMS = {
    'Maintenance': 'تعمیر و نگهداری',
    'Repair': 'تعمیر',
    'Inspection': 'بازرسی',
    'Preventive': 'پیشگیرانه',
    'Corrective': 'اصلاحی',
    'Equipment': 'تجهیزات',
    'Facility': 'تاسیسات',
    'Safety': 'ایمنی',
    'Operation': 'عملیات',
    'Procedure': 'روش اجرایی',
    'Standard': 'استاندارد',
    'Valve': 'شیر',
    'Pump': 'پمپ',
    'Pipeline': 'خط لوله',
    'Pressure': 'فشار',
    'Temperature': 'دما',
}

def extract_text_with_ocr(pdf_path, page_number):
    """استخراج متن با OCR از صفحه PDF"""
    doc = fitz.open(pdf_path)
    page = doc[page_number - 1]
    
    # تبدیل صفحه به تصویر
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # کیفیت بالا
    img_data = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))
    
    # OCR روی تصویر
    text = pytesseract.image_to_string(img, lang='eng')
    
    doc.close()
    return text

def translate_text_with_terms(text, chunk_size=4000):
    """ترجمه متن با حفظ اصطلاحات"""
    if not text or len(text.strip()) < 5:
        return ""
    
    translator = GoogleTranslator(source='en', target='fa')
    
    chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
    translated_chunks = []
    
    for i, chunk in enumerate(chunks):
        try:
            print(f"      ترجمه بخش {i+1}/{len(chunks)}...")
            translated = translator.translate(chunk)
            
            # اعمال اصطلاحات
            for eng, fa in OIL_GAS_TERMS.items():
                translated = re.sub(r'\b' + eng + r'\b', fa, translated, flags=re.IGNORECASE)
            
            translated_chunks.append(translated)
            time.sleep(0.3)
            
        except Exception as e:
            print(f"      ⚠️ خطا: {e}")
            translated_chunks.append(chunk)
    
    return " ".join(translated_chunks)

def create_pdf(output_path, pages_data, font_name):
    """ایجاد PDF با متن فارسی"""
    c = canvas.Canvas(output_path, pagesize=A4)
    page_width, page_height = A4
    
    for page_num, text in pages_data.items():
        print(f"   📄 ایجاد صفحه {page_num} در PDF...")
        
        y_position = page_height - 50
        
        # عنوان صفحه
        page_title = f"صفحه {page_num}"
        reshaped = arabic_reshaper.reshape(page_title)
        bidi_text = get_display(reshaped)
        c.setFont(font_name, 14)
        c.drawRightString(page_width - 50, y_position, bidi_text)
        y_position -= 40
        
        # متن ترجمه شده
        if text:
            reshaped = arabic_reshaper.reshape(text)
            bidi_text = get_display(reshaped)
            
            lines = bidi_text.split('\n')
            c.setFont(font_name, 10)
            
            for line in lines:
                if y_position < 50:
                    c.showPage()
                    y_position = page_height - 50
                
                # هر خط رو به چند قطعه کوچیک تقسیم کن
                max_width = 80
                words = line.split()
                current_line = ""
                
                for word in words:
                    if len(current_line) + len(word) < max_width:
                        current_line += word + " "
                    else:
                        c.drawRightString(page_width - 50, y_position, current_line)
                        y_position -= 15
                        current_line = word + " "
                        
                        if y_position < 50:
                            c.showPage()
                            y_position = page_height - 50
                
                if current_line:
                    c.drawRightString(page_width - 50, y_position, current_line)
                    y_position -= 15
        
        c.showPage()
    
    c.save()

def main():
    print("🚀 شروع فرآیند OCR و ترجمه...\n")
    
    # ثبت فونت
    try:
        pdfmetrics.registerFont(TTFont('BNazanin', FONT_PATH))
        print("✅ فونت فارسی بارگذاری شد\n")
    except Exception as e:
        print(f"❌ خطا در فونت: {e}")
        return
    
    pages_data = {}
    
    for page_num in range(START_PAGE, END_PAGE + 1):
        print(f"📄 پردازش صفحه {page_num}...")
        
        # استخراج با OCR
        text = extract_text_with_ocr(INPUT_PDF, page_num)
        
        if text and len(text.strip()) > 10:
            print(f"   ✅ OCR: {len(text)} کاراکتر استخراج شد")
            print(f"   🌐 در حال ترجمه...")
            translated = translate_text_with_terms(text)
            pages_data[page_num] = translated
        else:
            print(f"   ⚠️ متنی یافت نشد")
            pages_data[page_num] = ""
        
        print(f"✅ صفحه {page_num} تکمیل شد!\n")
        time.sleep(0.5)
    
    # ایجاد PDF
    print("📦 ایجاد PDF نهایی...")
    create_pdf(OUTPUT_PDF, pages_data, 'BNazanin')
    
    print(f"\n✅ تمام! فایل: {OUTPUT_PDF}")

if __name__ == "__main__":
    main()
