"""
Debug Script: Identify which 146 forms failed and WHY
This will help us fix the sync issue
"""

import pyodbc
import win32com.client
from datetime import datetime
import re

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
# Analyze forms
# ============================================
def analyze_forms():
    print("=" * 80)
    print("FORM ANALYSIS & DEBUG TOOL")
    print("=" * 80)
    
    # Step 1: Get all forms from Access
    print("\n[1] Reading forms from Access...")
    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False
    access.OpenCurrentDatabase(ACCESS_DB_PATH)
    
    access_forms = []
    all_forms = access.CurrentProject.AllForms
    
    for i in range(all_forms.Count):
        form_obj = all_forms.Item(i)
        form_name = form_obj.Name
        if not (form_name.startswith('~') or form_name.startswith('MSys')):
            access_forms.append(form_name)
    
    access.CloseCurrentDatabase()
    access.Quit()
    
    print(f"   Total forms in Access: {len(access_forms)}")
    
    # Step 2: Get all forms from SQL
    print("\n[2] Reading forms from SQL Server...")
    sql_conn = pyodbc.connect(SQL_CONN)
    sql_cursor = sql_conn.cursor()
    
    sql_cursor.execute("SELECT FormName FROM tbl_Forms")
    sql_forms = [row[0] for row in sql_cursor.fetchall()]
    
    print(f"   Total forms in SQL: {len(sql_forms)}")
    
    # Step 3: Find missing forms
    print("\n[3] Finding missing forms...")
    missing_forms = [f for f in access_forms if f not in sql_forms]
    
    print(f"   Missing forms: {len(missing_forms)}")
    
    if len(missing_forms) == 0:
        print("\n✓ All forms are synced!")
        return
    
    # Step 4: Analyze missing forms
    print("\n" + "=" * 80)
    print("MISSING FORMS ANALYSIS")
    print("=" * 80)
    
    # Group by prefix
    prefix_groups = {}
    for form in missing_forms:
        # Extract prefix (everything before first underscore)
        match = re.match(r'^([^_]+)_', form)
        if match:
            prefix = match.group(1)
        else:
            prefix = "NO_PREFIX"
        
        if prefix not in prefix_groups:
            prefix_groups[prefix] = []
        prefix_groups[prefix].append(form)
    
    print(f"\nMissing forms by prefix:")
    print("-" * 80)
    for prefix, forms in sorted(prefix_groups.items(), key=lambda x: len(x[1]), reverse=True):
        print(f"\n{prefix}_ : {len(forms)} forms")
        for form in forms[:5]:  # Show first 5
            print(f"  - {form}")
        if len(forms) > 5:
            print(f"  ... and {len(forms) - 5} more")
    
    # Step 5: Check for special characters
    print("\n" + "=" * 80)
    print("SPECIAL CHARACTER ANALYSIS")
    print("=" * 80)
    
    special_chars = {}
    for form in missing_forms:
        # Find non-standard characters
        for char in form:
            if not (char.isalnum() or char in ['_', '-', ' ']):
                if char not in special_chars:
                    special_chars[char] = []
                special_chars[char].append(form)
    
    if special_chars:
        print("\nForms with special characters:")
        for char, forms in special_chars.items():
            print(f"\n'{char}' (ASCII {ord(char)}): {len(forms)} forms")
            for form in forms[:3]:
                print(f"  - {form}")
    else:
        print("\n✓ No special characters found")
    
    # Step 6: Check for case sensitivity
    print("\n" + "=" * 80)
    print("CASE SENSITIVITY CHECK")
    print("=" * 80)
    
    lowercase_prefix = [f for f in missing_forms if f[0].islower()]
    print(f"\nForms starting with lowercase: {len(lowercase_prefix)}")
    if lowercase_prefix:
        for form in lowercase_prefix[:10]:
            print(f"  - {form}")
    
    # Step 7: Try to sync missing forms with better error handling
    print("\n" + "=" * 80)
    print("ATTEMPTING TO SYNC MISSING FORMS")
    print("=" * 80)
    
    success = 0
    failed = 0
    
    for idx, form_name in enumerate(missing_forms, 1):
        try:
            # Determine discipline and category
            discipline = get_discipline_code_enhanced(form_name)
            category = get_form_category(form_name)
            
            # Clean form name (remove special chars that might cause issues)
            clean_name = form_name.replace("'", "''")  # Escape single quotes
            
            merge_sql = """
            MERGE INTO tbl_Forms AS target
            USING (SELECT ? AS FormName) AS source
            ON target.FormName = source.FormName
            WHEN MATCHED THEN
                UPDATE SET 
                    FormCategory = ?,
                    DisciplineCode = ?
            WHEN NOT MATCHED THEN
                INSERT (FormName, FormDisplayName, FormCategory, DisciplineCode, IsActive)
                VALUES (?, ?, ?, ?, 1);
            """
            
            sql_cursor.execute(merge_sql, 
                             form_name, category, discipline,
                             form_name, form_name, category, discipline)
            sql_conn.commit()
            
            success += 1
            print(f"{idx:3d}. ✓ {form_name} [{discipline}]")
            
        except Exception as e:
            failed += 1
            print(f"{idx:3d}. ✗ {form_name}")
            print(f"       ERROR: {str(e)}")
    
    sql_cursor.close()
    sql_conn.close()
    
    # Final summary
    print("\n" + "=" * 80)
    print("FINAL SUMMARY")
    print("=" * 80)
    print(f"Missing forms processed:  {len(missing_forms)}")
    print(f"Successfully synced:      {success}")
    print(f"Still failed:             {failed}")
    print("=" * 80)

# ============================================
# Enhanced discipline detection
# ============================================
def get_discipline_code_enhanced(form_name):
    """Enhanced discipline detection with case-insensitive matching"""
    upper_name = form_name.upper()
    
    # Check for various patterns
    if upper_name.startswith('COR_') or upper_name.startswith('COR'):
        return 'COR'
    elif upper_name.startswith('CV_') or upper_name.startswith('CV'):
        return 'CV'
    elif upper_name.startswith('CP_'):
        return 'CP'
    elif upper_name.startswith('DC_'):
        return 'DC'
    elif upper_name.startswith('EL_'):
        return 'EL'
    elif upper_name.startswith('INS_'):
        return 'INS'
    elif upper_name.startswith('PA_'):
        return 'PA'
    elif upper_name.startswith('PI_'):
        return 'PI'
    elif upper_name.startswith('ST_'):
        return 'ST'
    elif upper_name.startswith('WH_'):
        return 'WH'
    elif upper_name.startswith('GEN_'):
        return 'GEN'
    elif upper_name.startswith('FRM_') or upper_name.startswith('FORM_'):
        return 'SYS'
    
    # Check for discipline name in middle of form name
    if 'CORROSION' in upper_name or '_COR_' in upper_name:
        return 'COR'
    elif 'CIVIL' in upper_name or '_CV_' in upper_name:
        return 'CV'
    elif 'PIPING' in upper_name or '_PI_' in upper_name:
        return 'PI'
    elif 'PAINTING' in upper_name or '_PA_' in upper_name:
        return 'PA'
    elif 'STRUCTURE' in upper_name or '_ST_' in upper_name:
        return 'ST'
    elif 'ELECTRICAL' in upper_name or '_EL_' in upper_name:
        return 'EL'
    elif 'INSTRUMENT' in upper_name or '_INS_' in upper_name:
        return 'INS'
    
    return 'SYS'

def get_form_category(form_name):
    """Determine form category"""
    upper_name = form_name.upper()
    
    if 'SUBFORM' in upper_name:
        return 'SubForm'
    elif 'NAVIGATION' in upper_name:
        return 'Navigation'
    elif 'CHART' in upper_name:
        return 'Report'
    elif 'EDIT' in upper_name:
        return 'Edit'
    elif 'HISTORY' in upper_name:
        return 'History'
    elif 'REPORT' in upper_name:
        return 'Report'
    elif upper_name == 'FRM_LOGIN':
        return 'System'
    elif upper_name.startswith('FRM_MAIN'):
        return 'Main'
    else:
        return 'Main'

# ============================================
# Run analysis
# ============================================
if __name__ == "__main__":
    try:
        analyze_forms()
    except Exception as e:
        print(f"\nFATAL ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
