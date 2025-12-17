import os
import re
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import io
import cv2
import numpy as np

# سعی کنید هر دو را امتحان کنید
try:
    import easyocr
    EASYOCR_AVAILABLE = True
    print("✅ EasyOCR در دسترس است")
except:
    EASYOCR_AVAILABLE = False
    print("⚠️  EasyOCR در دسترس نیست")

try:
    from paddleocr import PaddleOCR
    PADDLE_AVAILABLE = True
    print("✅ PaddleOCR در دسترس است")
except:
    PADDLE_AVAILABLE = False
    print("⚠️  PaddleOCR در دسترس نیست")


class PDFProcessor:
    def __init__(self):
        self.easy_reader = None
        self.paddle_ocr = None
        
        # راه‌اندازی OCR engines
        if EASYOCR_AVAILABLE:
            try:
                print("🔧 راه‌اندازی EasyOCR...")
                self.easy_reader = easyocr.Reader(['en'], gpu=False, verbose=False)
                print("✅ EasyOCR آماده است")
            except Exception as e:
                print(f"⚠️  خطا در راه‌اندازی EasyOCR: {e}")
        
        if PADDLE_AVAILABLE:
            try:
                print("🔧 راه‌اندازی PaddleOCR...")
                self.paddle_ocr = PaddleOCR(use_angle_cls=True, lang='en', show_log=False)
                print("✅ PaddleOCR آماده است")
            except Exception as e:
                print(f"⚠️  خطا در راه‌اندازی PaddleOCR: {e}")
    
    def preprocess_image(self, image):
        """
        پیش‌پردازش تصویر برای بهبود OCR
        """
        # تبدیل PIL به numpy array
        if isinstance(image, Image.Image):
            img = np.array(image)
        else:
            img = image
        
        # تبدیل به grayscale
        if len(img.shape) == 3:
            gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
        else:
            gray = img
        
        # افزایش کنتراست
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        enhanced = clahe.apply(gray)
        
        # Denoising
        denoised = cv2.fastNlMeansDenoising(enhanced)
        
        # Thresholding
        _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        return binary
    
    def extract_text_with_ocr(self, image):
        """
        استخراج متن با استفاده از OCR (چند روش)
        """
        text_results = []
        
        # پیش‌پردازش تصویر
        processed_img = self.preprocess_image(image)
        
        # روش 1: EasyOCR
        if self.easy_reader:
            try:
                result = self.easy_reader.readtext(processed_img, detail=0, paragraph=True)
                text = " ".join(result)
                text_results.append(text)
            except Exception as e:
                print(f"    ⚠️  خطا در EasyOCR: {e}")
        
        # روش 2: PaddleOCR
        if self.paddle_ocr:
            try:
                result = self.paddle_ocr.ocr(processed_img, cls=True)
                if result and result[0]:
                    text = " ".join([line[1][0] for line in result[0]])
                    text_results.append(text)
            except Exception as e:
                print(f"    ⚠️  خطا در PaddleOCR: {e}")
        
        # ترکیب نتایج
        combined_text = "\n".join(text_results)
        return combined_text
    
    def extract_text_from_pdf(self, pdf_path):
        """
        استخراج متن از PDF (هم متنی و هم اسکن شده)
        """
        all_text = ""
        
        try:
            doc = fitz.open(pdf_path)
            print(f"  📄 تعداد صفحات: {len(doc)}")
            
            # بررسی 3 صفحه اول
            for page_num in range(min(3, len(doc))):
                page = doc[page_num]
                
                # ابتدا تلاش برای استخراج متن مستقیم
                page_text = page.get_text()
                
                if page_text and len(page_text.strip()) > 100:
                    print(f"  ✅ صفحه {page_num + 1}: متن مستقیم استخراج شد")
                    all_text += page_text + "\n"
                else:
                    print(f"  📷 صفحه {page_num + 1}: استفاده از OCR...")
                    
                    # تبدیل صفحه به تصویر با کیفیت بالا
                    mat = fitz.Matrix(3, 3)  # zoom factor = 3
                    pix = page.get_pixmap(matrix=mat)
                    
                    # تبدیل به PIL Image
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    
                    # OCR روی تصویر
                    ocr_text = self.extract_text_with_ocr(img)
                    
                    if ocr_text and len(ocr_text.strip()) > 50:
                        print(f"    ✅ OCR موفق: {len(ocr_text)} کاراکتر")
                        all_text += ocr_text + "\n"
                    else:
                        print(f"    ⚠️  OCR نتیجه کافی نداد")
                        
                        # تلاش با استخراج تصاویر موجود در PDF
                        image_list = page.get_images()
                        for img_index, img in enumerate(image_list[:3]):  # حداکثر 3 تصویر
                            try:
                                xref = img[0]
                                base_image = doc.extract_image(xref)
                                image_bytes = base_image["image"]
                                image = Image.open(io.BytesIO(image_bytes))
                                
                                print(f"    📷 OCR روی تصویر شماره {img_index + 1}...")
                                img_text = self.extract_text_with_ocr(image)
                                all_text += img_text + "\n"
                            except Exception as e:
                                print(f"    ⚠️  خطا در تصویر {img_index + 1}: {e}")
            
            doc.close()
            
        except Exception as e:
            print(f"  ❌ خطا در پردازش PDF: {str(e)}")
        
        return all_text
    
    def extract_doc_info(self, text):
        """
        استخراج اطلاعات از متن با الگوهای بهبود یافته
        """
        doc_no = None
        date = None
        rev = None
        number = None
        
        # تمیز کردن متن
        text = re.sub(r'\s+', ' ', text)
        
        print(f"  🔍 طول متن استخراج شده: {len(text)} کاراکتر")
        
        # الگوهای گسترده‌تر برای Doc No
        doc_patterns = [
            r'Doc\s*\.?\s*No\s*\.?\s*[:\-]?\s*([A-Z0-9\-\s]+?)(?:\s+Rev|\s+Date|\s+G\d{2}|$)',
            r'Document\s+No\s*\.?\s*[:\-]?\s*([A-Z0-9\-\s]+?)(?:\s+Rev|\s+Date|\s+G\d{2}|$)',
            r'DOC\s*\.?\s*NO\s*\.?\s*[:\-]?\s*([A-Z0-9\-\s]+?)(?:\s+Rev|\s+Date|\s+G\d{2}|$)',
            r'Doc\s+Number\s*[:\-]?\s*([A-Z0-9\-\s]+?)(?:\s+Rev|\s+Date|\s+G\d{2}|$)',
            r'([A-Z]{3,5}\-[A-Z]{3,10}\-[A-Z]{3,10}\-[A-Z]{3,10}\-\d+\-G\d{2})',
        ]
        
        for pattern in doc_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                doc_no = match.group(1).strip()
                doc_no = re.sub(r'\s+', '-', doc_no)
                doc_no = re.sub(r'-+', '-', doc_no)
                print(f"  ✅ Doc No یافت شد: {doc_no}")
                break
        
        # استخراج Number و Rev
        if doc_no:
            parts = [p.strip() for p in doc_no.split('-') if p.strip()]
            
            # جستجوی Rev (Gxx)
            for i, part in enumerate(parts):
                if re.match(r'G\d{2}', part, re.IGNORECASE):
                    rev = part.upper()
                    if i > 0:
                        number = parts[i-1]
                    print(f"  ✅ Number: {number}, Rev: {rev}")
                    break
            
            # اگر پیدا نشد، از آخرین قسمت‌ها استفاده کن
            if not rev and len(parts) >= 2:
                # بررسی آخرین قسمت
                if re.match(r'G?\d{2}', parts[-1]):
                    rev = 'G' + re.sub(r'[^0-9]', '', parts[-1])
                    number = parts[-2]
                    print(f"  ℹ️  Number: {number}, Rev: {rev} (استنباطی)")
        
        # الگوهای تاریخ
        date_patterns = [
            r'Date\s*[:\-]?\s*(\d{1,2}[\s/\-\.]\w+[\s/\-\.]\d{2,4})',
            r'DATE\s*[:\-]?\s*(\d{1,2}[\s/\-\.]\w+[\s/\-\.]\d{2,4})',
            r'(\d{1,2}[\s/\-\.](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s/\-\.]\d{2,4})',
            r'Date\s*[:\-]?\s*(\d{1,2}[\s/\-\.]\d{1,2}[\s/\-\.]\d{2,4})',
        ]
        
        for pattern in date_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                date = match.group(1).strip()
                date = re.sub(r'\s+', ' ', date)
                print(f"  ✅ Date یافت شد: {date}")
                break
        
        return {
            'doc_no': doc_no,
            'number': number,
            'rev': rev,
            'date': date
        }


def process_pdfs(directory_path):
    """
    پردازش تمام فایل‌های PDF
    """
    if not os.path.exists(directory_path):
        print(f"❌ پوشه '{directory_path}' یافت نشد!")
        return
    
    pdf_files = [f for f in os.listdir(directory_path) 
                 if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print("⚠️  هیچ فایل PDF یافت نشد!")
        return
    
    print(f"\n📁 تعداد {len(pdf_files)} فایل PDF یافت شد.\n")
    print("="*70)
    
    processor = PDFProcessor()
    
    if not processor.easy_reader and not processor.paddle_ocr:
        print("\n❌ هیچ OCR engine‌ای در دسترس نیست!")
        print("لطفاً یکی از این دستورات را اجرا کنید:")
        print("  pip install easyocr")
        print("  یا")
        print("  pip install paddlepaddle paddleocr")
        return
    
    results = []
    renamed_count = 0
    failed_count = 0
    
    for idx, pdf_file in enumerate(pdf_files, 1):
        print(f"\n[{idx}/{len(pdf_files)}] 🔍 {pdf_file}")
        print("-"*70)
        
        pdf_path = os.path.join(directory_path, pdf_file)
        
        try:
            # استخراج متن
            text = processor.extract_text_from_pdf(pdf_path)
            
            if len(text.strip()) < 50:
                print(f"  ⚠️  متن کافی استخراج نشد ({len(text)} کاراکتر)")
            
            # استخراج اطلاعات
            info = processor.extract_doc_info(text)
            
            if info['number'] and info['rev']:
                new_name = f"SJSC-GGNRSP-MADR-REWK-{info['number']}-{info['rev']}.pdf"
                new_path = os.path.join(directory_path, new_name)
                
                if not os.path.exists(new_path) and pdf_file != new_name:
                    os.rename(pdf_path, new_path)
                    print(f"  ✅ تغییر نام به: {new_name}")
                    renamed_count += 1
                    status = 'موفق'
                elif pdf_file == new_name:
                    print(f"  ℹ️  نام فایل از قبل صحیح است")
                    new_name = pdf_file
                    status = 'نام صحیح'
                else:
                    print(f"  ⚠️  فایل با این نام از قبل وجود دارد!")
                    new_name = pdf_file
                    status = 'تکراری'
                
                results.append({
                    'ردیف': idx,
                    'نام فایل قدیم': pdf_file,
                    'نام فایل جدید': new_name,
                    'Doc No': info['doc_no'],
                    'Number': info['number'],
                    'Rev': info['rev'],
                    'Date': info['date'],
                    'وضعیت': status
                })
            else:
                print(f"  ❌ اطلاعات کافی استخراج نشد")
                failed_count += 1
                results.append({
                    'ردیف': idx,
                    'نام فایل قدیم': pdf_file,
                    'نام فایل جدید': pdf_file,
                    'Doc No': info['doc_no'] or 'نامشخص',
                    'Number': info['number'] or 'نامشخص',
                    'Rev': info['rev'] or 'نامشخص',
                    'Date': info['date'] or 'نامشخص',
                    'وضعیت': 'ناموفق'
                })
        
        except Exception as e:
            print(f"  ❌ خطا: {str(e)}")
            import traceback
            traceback.print_exc()
            failed_count += 1
            results.append({
                'ردیف': idx,
                'نام فایل قدیم': pdf_file,
                'نام فایل جدید': pdf_file,
                'Doc No': 'خطا',
                'Number': 'خطا',
                'Rev': 'خطا',
                'Date': 'خطا',
                'وضعیت': f'خطا: {str(e)[:50]}'
            })
    
    # ذخیره گزارش
    df = pd.DataFrame(results)
    excel_path = os.path.join(directory_path, 'PDF_Report.xlsx')
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='گزارش', index=False)
        
        # تنظیم عرض ستون‌ها
        worksheet = writer.sheets['گزارش']
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
    print(f"\n{'='*70}")
    print(f"📊 خلاصه نتایج:")
    print(f"  ✅ موفق: {renamed_count} فایل")
    print(f"  ❌ ناموفق: {failed_count} فایل")
    print(f"  📁 کل: {len(pdf_files)} فایل")
    print(f"\n📄 گزارش کامل: {excel_path}")
    print(f"{'='*70}")


if __name__ == "__main__":
    folder_path = r"D:\Sepher_Pasargad\works\Maintenace\Maintenance Report\All_Extracted\weekly"
    
    print("🚀 شروع پردازش فایل‌های PDF...")
    print("="*70)
    
    process_pdfs(folder_path)
    
    print("\n✨ پردازش کامل شد!")