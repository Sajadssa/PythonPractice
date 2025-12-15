"""
Final Form Sync: Insert ALL forms WITHOUT Discipline validation
This bypasses Foreign Key issues by setting DisciplineCode to NULL initially

Usage: python sync_all_forms_final.py
"""

import pyodbc
import win32com.client
from datetime import datetime

# ============================================
# CONFIGURATION - UPDATE THESE
# ============================================
ACCESS_DB_PATH = r"C:\Users\SaeeidiAzad\Desktop\IDMS_Rev_2.1.2.accdb"  # UPDATE THIS
SQL_SERVER = "DCC-SAEEDI"
SQL_DATABASE = "IDMS_WRFM"
SQL_USERNAME = None
SQL_PASSWORD = None
# Build connection string
if SQL_USERNAME:
    SQL_CONN = f"DRIVER={{SQL Server}};SERVER={SQL_SERVER};DATABASE={SQL_DATABASE};UID={SQL_USERNAME};PWD={SQL_PASSWORD}"
else:
    SQL_CONN = f"DRIVER={{SQL Server}};SERVER={SQL_SERVER};DATABASE={SQL_DATABASE};Trusted_Connection=yes"

# ============================================
# Step 0: Check and add missing disciplines
# ============================================
def ensure_all_disciplines(cursor, conn):
    """Add all possible discipline codes to avoid FK errors"""
    
    print("[0] Ensuring all discipline codes exist...")
    
    all_disciplines = [
        ('ALL', 'All Disciplines', 'همه رشته‌ها', None),
        ('PI', 'Piping & Inspection', 'لوله کشی و بازرسی', 'Pi_'),
        ('CV', 'Civil', 'عمران', 'Cv_'),
        ('COR', 'Corrosion', 'خوردگی', 'Cor_'),
        ('EL', 'Electrical', 'برق', 'El_'),
        ('INS', 'Instrument', 'ابزار دقیق', 'Ins_'),
        ('ST', 'Structure', 'سازه', 'St_'),
        ('PA', 'Painting', 'رنگ', 'Pa_'),
        ('WH', 'Warehouse', 'انبار', 'Wh_'),
        ('DC', 'Document Control', 'کنترل مدارک', 'Dc_'),
        ('CP', 'Construction Planning', 'برنامه‌ریزی', 'Cp_'),
        ('GEN', 'General', 'عمومی', 'Gen_'),
        ('SYS', 'System Forms', 'فرم‌های سیستمی', 'Frm_'),
        ('FRM', 'Forms', 'فرم', 'frm_'),
        ('FORM', 'Form', 'فرم', 'Form_'),
    ]
    
    for code, name, name_fa, prefix in all_disciplines:
        try:
            # Check if exists
            cursor.execute("SELECT COUNT(*) FROM tbl_Disciplines WHERE DisciplineCode = ?", code)
            if cursor.fetchone()[0] == 0:
                # Insert if not exists
                cursor.execute("""
                    INSERT INTO tbl_Disciplines (DisciplineCode, DisciplineName, DisciplineNameFA, DisciplinePrefix, IsActive)
                    VALUES (?, ?, ?, ?, 1)
                """, code, name, name_fa, prefix)
                conn.commit()
                print(f"    + Added: {code}")
        except Exception as e:
            print(f"    - Could not add {code}: {str(e)}")
            continue
    
    print("    ✓ Discipline check complete\n")

# ============================================
# Simple discipline detection
# ============================================
def get_discipline_safe(form_name):
    """
    Get discipline code with guaranteed match in tbl_Disciplines
    Returns 'SYS' as safe default if nothing matches
    """
    name_lower = form_name.lower()
    
    # Check exact prefix matches (case insensitive)
    if name_lower.startswith('cor'):
        return 'COR'
    elif name_lower.startswith('cv'):
        return 'CV'
    elif name_lower.startswith('cp'):
        return 'CP'
    elif name_lower.startswith('dc'):
        return 'DC'
    elif name_lower.startswith('el'):
        return 'EL'
    elif name_lower.startswith('ins'):
        return 'INS'
    elif name_lower.startswith('pa'):
        return 'PA'
    elif name_lower.startswith('pi'):
        return 'PI'
    elif name_lower.startswith('st'):
        return 'ST'
    elif name_lower.startswith('wh'):
        return 'WH'
    elif name_lower.startswith('gen'):
        return 'GEN'
    elif name_lower.startswith('frm'):
        return 'FRM'
    elif name_lower.startswith('form'):
        return 'FORM'
    else:
        return 'SYS'  # Safe default

def get_category_safe(form_name):
    """Simple category detection"""
    name_lower = form_name.lower()
    
    if 'subform' in name_lower:
        return 'SubForm'
    elif 'navigation' in name_lower:
        return 'Navigation'
    elif 'chart' in name_lower or 'report' in name_lower:
        return 'Report'
    elif 'edit' in name_lower:
        return 'Edit'
    elif 'history' in name_lower:
        return 'History'
    else:
        return 'Main'

# ============================================
# Main sync function
# ============================================
def sync_all_forms_final():
    print("=" * 80)
    print("FINAL FORM SYNC - Insert ALL Forms (No FK Issues)")
    print("=" * 80)
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Step 1: Connect to Access
    print("[1] Connecting to Access...")
    try:
        access = win32com.client.Dispatch("Access.Application")
        access.Visible = False
        access.OpenCurrentDatabase(ACCESS_DB_PATH)
        print(f"    ✓ Connected\n")
    except Exception as e:
        print(f"    ✗ ERROR: {str(e)}")
        return False
    
    # Step 2: Get ALL forms
    print("[2] Reading ALL forms...")
    all_forms = []
    
    try:
        forms_collection = access.CurrentProject.AllForms
        
        for i in range(forms_collection.Count):
            form_name = forms_collection.Item(i).Name
            
            # Skip ONLY MSys and temp forms
            if not (form_name.startswith('MSys') or form_name.startswith('~')):
                all_forms.append(form_name)
        
        print(f"    ✓ Found {len(all_forms)} forms\n")
        
    except Exception as e:
        print(f"    ✗ ERROR: {str(e)}")
        access.CloseCurrentDatabase()
        access.Quit()
        return False
    
    access.CloseCurrentDatabase()
    access.Quit()
    print("    ✓ Access closed\n")
    
    # Step 3: Connect to SQL
    print("[3] Connecting to SQL Server...")
    try:
        conn = pyodbc.connect(SQL_CONN)
        cursor = conn.cursor()
        print(f"    ✓ Connected\n")
    except Exception as e:
        print(f"    ✗ ERROR: {str(e)}")
        return False
    
    # Step 3.5: Ensure all disciplines exist
    ensure_all_disciplines(cursor, conn)
    
    # Step 4: Insert forms
    print("[4] Inserting forms...")
    print("-" * 80)
    
    success = 0
    failed = 0
    failed_forms = []
    
    for idx, form_name in enumerate(all_forms, 1):
        try:
            discipline = get_discipline_safe(form_name)
            category = get_category_safe(form_name)
            
            # Use simple INSERT with IF NOT EXISTS (safer than MERGE)
            sql = """
            IF NOT EXISTS (SELECT 1 FROM tbl_Forms WHERE FormName = ?)
            BEGIN
                INSERT INTO tbl_Forms (FormName, FormDisplayName, FormCategory, DisciplineCode, IsActive)
                VALUES (?, ?, ?, ?, 1)
            END
            ELSE
            BEGIN
                UPDATE tbl_Forms 
                SET FormCategory = ?, DisciplineCode = ?
                WHERE FormName = ?
            END
            """
            
            cursor.execute(sql, 
                         form_name,      # IF NOT EXISTS check
                         form_name,      # INSERT FormName
                         form_name,      # INSERT FormDisplayName
                         category,       # INSERT FormCategory
                         discipline,     # INSERT DisciplineCode
                         category,       # UPDATE FormCategory
                         discipline,     # UPDATE DisciplineCode
                         form_name)      # UPDATE WHERE
            
            conn.commit()
            success += 1
            
            if idx % 20 == 0 or idx == len(all_forms):
                print(f"  Progress: {idx}/{len(all_forms)} ({int(idx/len(all_forms)*100)}%) - {form_name[:50]}")
            
        except Exception as e:
            failed += 1
            failed_forms.append((form_name, str(e)))
            print(f"  ✗ Failed: {form_name}")
            print(f"     Error: {str(e)[:100]}")
            continue
    
    print("-" * 80)
    
    # Step 5: Verify
    print("\n[5] Verification...")
    try:
        cursor.execute("SELECT COUNT(*) FROM tbl_Forms")
        total = cursor.fetchone()[0]
        
        cursor.execute("""
            SELECT DisciplineCode, COUNT(*) as cnt 
            FROM tbl_Forms 
            GROUP BY DisciplineCode 
            ORDER BY cnt DESC
        """)
        
        print(f"    Total forms in SQL: {total}\n")
        print("    By Discipline:")
        for disc, count in cursor.fetchall():
            print(f"    {disc:<8} : {count:>4} forms")
        
    except Exception as e:
        print(f"    ✗ Error: {str(e)}")
    
    cursor.close()
    conn.close()
    
    # Final summary
    print("\n" + "=" * 80)
    print("SYNC COMPLETE")
    print("=" * 80)
    print(f"Total forms:       {len(all_forms)}")
    print(f"Successfully:      {success}")
    print(f"Failed:            {failed}")
    print(f"Success rate:      {int(success/len(all_forms)*100)}%")
    print("=" * 80)
    
    if failed > 0:
        print(f"\n⚠ {failed} forms failed:")
        for form, err in failed_forms[:10]:
            print(f"  - {form}: {err[:80]}")
        if len(failed_forms) > 10:
            print(f"  ... and {len(failed_forms)-10} more")
    
    if failed == 0:
        print("\n✓✓✓ ALL FORMS SYNCED SUCCESSFULLY! ✓✓✓")
        return True
    else:
        return False

# ============================================
# Run
# ============================================
if __name__ == "__main__":
    try:
        result = sync_all_forms_final()
        exit(0 if result else 1)
    except KeyboardInterrupt:
        print("\n\nCancelled")
        exit(1)
    except Exception as e:
        print(f"\n\nFATAL ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        exit(1)