import pandas as pd
import os
from pathlib import Path
from datetime import datetime
import numpy as np

def find_header_row(df):
    """
    پیدا کردن ردیف هدر واقعی جدول
    """
    header_keywords = ['Location', 'Date', 'Point No', 'POS', 'Line Number', 'Material', 'N.Size', 'Class']
    
    for idx, row in df.iterrows():
        row_str = row.astype(str).str.lower()
        matches = sum([any(keyword.lower() in val for val in row_str) for keyword in header_keywords])
        
        if matches >= 4:
            return idx
    
    return None

def extract_location_date(df_raw, header_row):
    """
    استخراج Location و Date از قسمت بالای فایل (قبل از هدر جدول)
    """
    location = None
    date = None
    
    # جستجو در ردیف‌های قبل از هدر
    for idx in range(max(0, header_row - 10), header_row):
        row = df_raw.iloc[idx]
        row_str = ' '.join(row.astype(str).values)
        
        # جستجوی Location
        if 'LOCATION' in row_str.upper() or 'Location' in row_str:
            for cell in row.values:
                cell_str = str(cell).strip()
                # پیدا کردن مقدار Location (معمولاً به صورت JR-XX یا مشابه)
                if cell_str and cell_str not in ['LOCATION', 'Location', 'nan']:
                    location = cell_str
                    break
        
        # جستجوی Date
        if 'DATE' in row_str.upper() or 'Date' in row_str or 'REPORT' in row_str.upper():
            for cell in row.values:
                cell_str = str(cell).strip()
                # چک کردن اینکه آیا شامل تاریخ است
                if '/' in cell_str or '-' in cell_str:
                    if cell_str not in ['DATE', 'Date', 'nan']:
                        date = cell_str
                        break
    
    return location, date

def clean_dataframe(df, header_row, location_value=None, date_value=None):
    """
    پاکسازی و استاندارد کردن DataFrame
    """
    # استفاده از ردیف مشخص شده به عنوان هدر
    new_columns = df.iloc[header_row].values
    
    # رفع مشکل ستون‌های تکراری
    seen = {}
    unique_columns = []
    for col in new_columns:
        col_str = str(col).strip() if pd.notna(col) else 'Unnamed'
        if col_str in seen:
            seen[col_str] += 1
            unique_columns.append(f"{col_str}_{seen[col_str]}")
        else:
            seen[col_str] = 0
            unique_columns.append(col_str)
    
    df.columns = unique_columns
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    
    # حذف ستون‌های Unnamed و خالی
    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
    df = df.dropna(axis=1, how='all')
    
    # حذف ردیف‌های کاملاً خالی
    df = df.dropna(how='all')
    
    # فیلتر کردن ردیف‌های معتبر با Point No
    if 'Point No' in df.columns:
        pattern = r'^P\d+'
        mask = df['Point No'].astype(str).str.match(pattern, na=False)
        df = df[mask]
    
    # اضافه کردن Location و Date به ابتدای جدول
    if 'Location' not in df.columns:
        df.insert(0, 'Location', location_value)
    else:
        # اگر Location خالی است، از مقدار استخراج شده استفاده کن
        if df['Location'].isna().all() and location_value:
            df['Location'] = location_value
        # Forward fill
        df['Location'] = df['Location'].ffill()
    
    if 'Date' not in df.columns:
        df.insert(1, 'Date', date_value)
    else:
        # اگر Date خالی است، از مقدار استخراج شده استفاده کن
        if df['Date'].isna().all() and date_value:
            df['Date'] = date_value
        # Forward fill
        df['Date'] = df['Date'].ffill()
    
    return df

def combine_excel_files(source_folder, output_file=None):
    """
    ترکیب تمام فایل‌های اکسل با پیدا کردن خودکار هدر و پاکسازی داده‌ها
    """
    
    source_path = Path(source_folder)
    
    if not source_path.exists():
        print(f"❌ خطا: پوشه {source_folder} وجود ندارد!")
        return
    
    excel_files = list(source_path.glob('*.xlsx')) + list(source_path.glob('*.xls'))
    
    if not excel_files:
        print(f"❌ هیچ فایل اکسلی در پوشه {source_folder} پیدا نشد!")
        return
    
    print(f"📊 تعداد {len(excel_files)} فایل اکسل پیدا شد")
    print("="*80)
    
    all_dataframes = []
    total_rows = 0
    total_sheets = 0
    
    columns_to_fill = ['Location', 'Date']
    
    for idx, excel_file in enumerate(excel_files, 1):
        try:
            print(f"\n🔄 در حال پردازش فایل {idx}/{len(excel_files)}: {excel_file.name}")
            
            excel_data = pd.ExcelFile(excel_file)
            sheet_names = excel_data.sheet_names
            
            print(f"   📑 تعداد شیت‌ها: {len(sheet_names)}")
            
            for sheet_name in sheet_names:
                try:
                    df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                    
                    header_row = find_header_row(df_raw)
                    
                    if header_row is None:
                        print(f"   ⚠️  شیت '{sheet_name}': هدر پیدا نشد")
                        continue
                    
                    print(f"   📍 شیت '{sheet_name}': هدر در ردیف {header_row + 1} پیدا شد")
                    
                    # استخراج Location و Date از قسمت بالای فایل
                    location_from_header, date_from_header = extract_location_date(df_raw, header_row)
                    
                    df = clean_dataframe(df_raw, header_row, location_from_header, date_from_header)
                    
                    if df.empty:
                        print(f"   ⚠️  شیت '{sheet_name}': بعد از پاکسازی خالی شد")
                        continue
                    
                    # نمایش اطلاعات Location و Date
                    if location_from_header:
                        print(f"   📍 Location: {location_from_header}")
                    if date_from_header:
                        print(f"   📅 Date: {date_from_header}")
                    
                    # اطمینان از اینکه Location و Date پر شده‌اند
                    if 'Location' in df.columns:
                        df = df[df['Location'].notna()]
                    
                    if df.empty:
                        print(f"   ⚠️  شیت '{sheet_name}': داده معتبری پیدا نشد")
                        continue
                    
                    df.columns = df.columns.str.strip()
                    
                    all_dataframes.append(df)
                    total_rows += len(df)
                    total_sheets += 1
                    
                    print(f"   ✅ شیت '{sheet_name}': {len(df)} ردیف معتبر")
                    
                except Exception as e:
                    print(f"   ❌ خطا در شیت '{sheet_name}': {str(e)}")
                    continue
            
        except Exception as e:
            print(f"❌ خطا در فایل {excel_file.name}: {str(e)}")
            continue
    
    if not all_dataframes:
        print("\n❌ هیچ داده‌ای برای ترکیب پیدا نشد!")
        return
    
    print("\n" + "="*80)
    print("🔗 در حال ترکیب تمام داده‌ها...")
    
    # یکسان‌سازی ستون‌ها
    all_columns = set()
    for df in all_dataframes:
        all_columns.update(df.columns)
    
    standardized_dfs = []
    for df in all_dataframes:
        for col in all_columns:
            if col not in df.columns:
                df[col] = None
        df = df[sorted(df.columns)]
        standardized_dfs.append(df)
    
    # ترکیب
    try:
        combined_df = pd.concat(standardized_dfs, ignore_index=True)
    except Exception as e:
        print(f"❌ خطا در ترکیب: {str(e)}")
        combined_df = pd.DataFrame()
        for df in standardized_dfs:
            combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    combined_df = combined_df.dropna(how='all')
    
    # مرتب‌سازی ستون‌ها
    preferred_order = ['Location', 'Date', 'Point No', 'POS', 'Line Number', 
                      'Material', 'N.Size', 'Class', 'N.W.T', 'W.T Measurement (mm)',
                      'C.R', 'C.A', 'M.A.W.P', 'M.R.T', 'Next Ins.']
    
    existing_cols = [col for col in preferred_order if col in combined_df.columns]
    other_cols = [col for col in combined_df.columns if col not in existing_cols]
    combined_df = combined_df[existing_cols + other_cols]
    
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = source_path / f"Combined_Thickness_Report_{timestamp}.xlsx"
    else:
        output_file = Path(output_file)
    
    try:
        print(f"💾 در حال ذخیره فایل نهایی...")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            combined_df.to_excel(writer, sheet_name='Combined_Data', index=False)
            
            worksheet = writer.sheets['Combined_Data']
            for idx, col in enumerate(combined_df.columns):
                max_length = max(
                    combined_df[col].astype(str).apply(len).max(),
                    len(str(col))
                ) + 2
                if idx < 26:
                    col_letter = chr(65 + idx)
                else:
                    col_letter = chr(65 + idx // 26 - 1) + chr(65 + idx % 26)
                worksheet.column_dimensions[col_letter].width = min(max_length, 50)
        
        print(f"\n✅ عملیات با موفقیت انجام شد!")
        print(f"📁 فایل خروجی: {output_file}")
        print(f"📊 تعداد کل ردیف‌ها: {len(combined_df):,}")
        print(f"📋 تعداد ستون‌ها: {len(combined_df.columns)}")
        print(f"📑 تعداد شیت‌های پردازش شده: {total_sheets}")
        
        print("\n📝 ستون‌های موجود:")
        for i, col in enumerate(combined_df.columns, 1):
            print(f"   {i}. {col}")
        
        return combined_df
        
    except Exception as e:
        print(f"\n❌ خطا در ذخیره فایل: {str(e)}")
        return None


if __name__ == "__main__":
    source_folder = r"D:\Sepher_Pasargad\works\qc\report\thickness"
    
    print("🚀 شروع فرآیند ترکیب فایل‌های اکسل...")
    print(f"📂 پوشه مبدا: {source_folder}")
    print("="*80)
    
    result = combine_excel_files(source_folder)
    
    if result is not None:
        print("\n" + "="*80)
        print("🎉 پردازش کامل شد!")
        print("\n📋 نمونه 20 ردیف اول:")
        print(result.head(20).to_string())
        print("\n💡 نکات:")
        print("   ✅ هدرهای صحیح به صورت خودکار شناسایی شدند")
        print("   ✅ فقط ردیف‌های داده (با Point No معتبر) نگهداری شدند")
        print("   ✅ ردیف‌های توضیحی و خالی حذف شدند")
        print("   ✅ Location و Date برای تمام ردیف‌ها تکرار شدند")