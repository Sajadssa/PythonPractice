"""
Script: Update SharePoint Document Library Metadata from CSV
Description: Updates file metadata in SharePoint based on CSV data
"""

import pandas as pd
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from datetime import datetime
import sys
from urllib.parse import quote

# ==========================================
# تنظیمات اولیه
# ==========================================

# آدرس SharePoint Site
SITE_URL = "https://extranet.pedc.ir/pogp/PRD"

# نام Document Library
LIBRARY_NAME = "Production Engineering Report"

# مسیر فایل CSV
CSV_FILE_PATH = "D:\Sepher_Pasargad\works\Maintenace\PythonDataAnalysis\PythonPractice\\Weekly.csv"

# اطلاعات ورود (باید تغییر دهید)
USERNAME = "s.saeidi@pogp.ir"  # ایمیل یا username خود را وارد کنید
PASSWORD = "K@rensajad1367"          # پسورد خود را وارد کنید

# ==========================================
# توابع کمکی
# ==========================================

def connect_to_sharepoint(site_url, username, password):
    """اتصال به SharePoint"""
    try:
        print("🔗 Connecting to SharePoint...")
        credentials = UserCredential(username, password)
        ctx = ClientContext(site_url).with_credentials(credentials)
        
        # تست اتصال
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        
        print(f"✅ Connected successfully to: {web.properties['Title']}")
        return ctx
    except Exception as e:
        print(f"❌ Error connecting to SharePoint: {str(e)}")
        sys.exit(1)


def read_csv_file(csv_path):
    """خواندن فایل CSV"""
    try:
        print(f"\n📂 Reading CSV file: {csv_path}")
        
        # خواندن CSV با encoding مناسب
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
        
        # پاک کردن whitespace از نام ستون‌ها
        df.columns = df.columns.str.strip()
        
        print(f"✅ Found {len(df)} rows in CSV")
        print(f"📊 Columns: {', '.join(df.columns.tolist())}")
        
        return df
    except Exception as e:
        print(f"❌ Error reading CSV file: {str(e)}")
        sys.exit(1)


def get_all_files(ctx, library_name):
    """دریافت تمام فایل‌های Document Library"""
    try:
        print(f"\n📁 Getting all files from '{library_name}'...")
        
        # دریافت لیست
        list_obj = ctx.web.lists.get_by_title(library_name)
        
        # دریافت تمام آیتم‌ها
        items = list_obj.items.get_all(5000).execute_query()
        
        print(f"✅ Found {len(items)} files in library")
        
        # ساخت دیکشنری برای دسترسی سریع‌تر
        files_dict = {}
        for item in items:
            file_name = item.properties.get('FileLeafRef', '')
            if file_name:
                files_dict[file_name] = item
        
        return items, files_dict
    except Exception as e:
        print(f"❌ Error getting files: {str(e)}")
        return [], {}


def update_file_metadata(ctx, item, row_data, library_name):
    """آپدیت metadata یک فایل"""
    try:
        # آماده‌سازی داده‌ها برای آپدیت
        update_values = {}
        
        # ReportDate
        if pd.notna(row_data.get('ReportDate')):
            try:
                # تبدیل تاریخ به فرمت ISO
                date_str = str(row_data['ReportDate'])
                date_obj = pd.to_datetime(date_str)
                update_values['ReportDate'] = date_obj.strftime('%Y-%m-%dT%H:%M:%SZ')
            except:
                pass
        
        # سایر فیلدها
        field_mappings = {
            'Pttern': 'Pttern',
            'Rev': 'Rev',
            'Process': 'Process',
            'Subprocess': 'Subprocess',
            'Location': 'Location',
            'Subject': 'Subject',
            'Type': 'Type',
            'Contractor': 'Contractor',
            'MainGroup': 'MainGroup'
        }
        
        for csv_field, sp_field in field_mappings.items():
            value = row_data.get(csv_field)
            if pd.notna(value) and str(value).strip():
                update_values[sp_field] = str(value).strip()
        
        # آپدیت فقط اگر داده وجود داشته باشد
        if update_values:
            item.set_property_value_list(update_values)
            item.update()
            ctx.execute_query()
            return True, "Updated successfully"
        else:
            return False, "No values to update"
            
    except Exception as e:
        return False, f"Error: {str(e)}"


def find_matching_files(files_dict, report_no):
    """پیدا کردن فایل‌های مرتبط با Report No"""
    matching_files = []
    report_no_clean = str(report_no).strip()
    
    for file_name, item in files_dict.items():
        if report_no_clean in file_name:
            matching_files.append((file_name, item))
    
    return matching_files


# ==========================================
# تابع اصلی
# ==========================================

def main():
    print("=" * 70)
    print("SharePoint Document Library Metadata Updater")
    print("=" * 70)
    
    # 1. اتصال به SharePoint
    ctx = connect_to_sharepoint(SITE_URL, USERNAME, PASSWORD)
    
    # 2. خواندن CSV
    df = read_csv_file(CSV_FILE_PATH)
    
    # 3. دریافت فایل‌های موجود
    all_items, files_dict = get_all_files(ctx, LIBRARY_NAME)
    
    if not files_dict:
        print("❌ No files found in library!")
        return
    
    # 4. آپدیت فایل‌ها
    print("\n" + "=" * 70)
    print("Starting Update Process...")
    print("=" * 70)
    
    stats = {
        'total': len(df),
        'success': 0,
        'not_found': 0,
        'errors': 0,
        'no_update': 0
    }
    
    for index, row in df.iterrows():
        report_no = row.get('Report No', '')
        
        if pd.isna(report_no) or not str(report_no).strip():
            print(f"\n⚠️  Row {index + 1}: Missing Report No")
            stats['errors'] += 1
            continue
        
        # پیدا کردن فایل‌های مرتبط
        matching_files = find_matching_files(files_dict, report_no)
        
        if not matching_files:
            print(f"\n⚠️  Row {index + 1}: File not found for Report No: {report_no}")
            stats['not_found'] += 1
            continue
        
        # آپدیت هر فایل مرتبط
        for file_name, item in matching_files:
            print(f"\n📝 Row {index + 1}: Updating '{file_name}'")
            
            success, message = update_file_metadata(ctx, item, row, LIBRARY_NAME)
            
            if success:
                print(f"   ✅ {message}")
                stats['success'] += 1
            elif "No values" in message:
                print(f"   ⊘  {message}")
                stats['no_update'] += 1
            else:
                print(f"   ❌ {message}")
                stats['errors'] += 1
    
    # 5. نمایش خلاصه
    print("\n" + "=" * 70)
    print("Update Summary:")
    print("=" * 70)
    print(f"📊 Total rows in CSV:      {stats['total']}")
    print(f"✅ Successfully updated:   {stats['success']}")
    print(f"⊘  No values to update:    {stats['no_update']}")
    print(f"⚠️  Files not found:        {stats['not_found']}")
    print(f"❌ Errors:                 {stats['errors']}")
    print("=" * 70)
    
    print("\n✨ Process completed!")


# ==========================================
# اجرای برنامه
# ==========================================

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Process interrupted by user!")
        sys.exit(0)
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)