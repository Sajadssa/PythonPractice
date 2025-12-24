import os
import win32com.client

# مسیر ذخیره پیوست‌ها
save_path = r"D:\Sepher_Pasargad\works\qc\reports"
if not os.path.exists(save_path):
    os.makedirs(save_path)

# اتصال به Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox؛ اگر فولدر دیگه‌ای هست، تغییر بدید (مثل subfolder = inbox.Folders['نام فولدر'])

# فرستنده مورد نظر
sender_email = "d.bazargan@pogp.ir"

# فیلتر برای ایمیل‌های با پیوست از فرستنده خاص (برای Exchange)
restrict_criteria = f"[Attachment] > 0"  # فقط ایمیل‌های با پیوست

messages = inbox.Items.Restrict(restrict_criteria)

def get_sender_email(message):
    try:
        # برای حساب Exchange
        return message.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
    except:
        # برای حساب‌های غیر-Exchange
        try:
            return message.SenderEmailAddress.lower()
        except:
            return ""

def save_attachments_from_sender():
    count_emails_found = 0
    count_saved = 0
    for message in messages:
        try:
            if message.Class == 43:  # olMailItem
                current_sender = get_sender_email(message)
                if current_sender == sender_email.lower() and message.Attachments.Count > 0:
                    count_emails_found += 1
                    attachments = message.Attachments
                    for attachment in attachments:
                        # نام فایل اصلی (شامل .docx, .pdf و غیره)
                        filename = attachment.FileName
                        filepath = os.path.join(save_path, filename)
                        
                        # چک کردن تکراری بودن و اضافه کردن شماره اگر لازم باشه
                        i = 1
                        while os.path.exists(filepath):
                            name, ext = os.path.splitext(filename)
                            filepath = os.path.join(save_path, f"{name}_{i}{ext}")
                            i += 1
                        
                        # ذخیره پیوست
                        attachment.SaveAsFile(filepath)
                        count_saved += 1
                    
                    # اختیاری: علامت‌گذاری ایمیل به عنوان read
                    if message.UnRead:
                        message.UnRead = False
                        message.Save()
        except Exception as e:
            print(f"خطا در پردازش ایمیل: {e}")
    
    print(f"تعداد ایمیل‌های پیدا‌شده از فرستنده: {count_emails_found}")
    print(f"تعداد پیوست‌های ذخیره‌شده: {count_saved}")

# اجرا
save_attachments_from_sender()