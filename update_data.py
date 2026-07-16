import pandas as pd
import json
from datetime import datetime
import re

def generate_code(last_name, first_name, index):
    """Generate consistent code matching parent_codes.xlsx format"""
    code = (str(last_name)[:3].upper() + 
            str(first_name)[:3].upper() + 
            str(int(index) + 1000))
    code = re.sub(r'[^A-Z0-9]', '', code)
    return code

def update_school_data(excel_path='2025-2026 SHARED FEE.xlsx'):
    print("🔄 Processing latest fee data...")
    
    # Read the main fee data
    df = pd.read_excel(excel_path, sheet_name='3RD TERM 2025-2026 (2)')
    df = df.iloc[:-1].fillna('')  # Remove total row
    
    # Define the columns we want (J through AB)
    # Based on your Excel columns:
    # J: AD_FORM, K: TUITION, L: PTA&EXAM, M: REPORT_CARD, N: PRACTICAL
    # O: TEXTBOOKS, P: NB & STAT, Q: UNIFORM, R: HOODY, S: SPORTSWEAR
    # T: AFTERSCHOOL, U: EXCURSION, V: EXTERNAL EXAMINATION
    # W: GRADUATION/ SPORTS/PARTY, X: HOLIDAY LESSON, Y: OUTSTANDING
    # Z: TOTAL FEE, AA: TOTAL PAID, AB: BALANCE
    
    # Column mapping (Excel column letter -> column name)
    fee_columns = [
        'AD_FORM',           # J
        'TUITION',           # K
        'PTA&EXAM',          # L
        'REPORT_CARD',       # M
        'PRACTICAL',         # N
        'TEXTBOOKS',         # O
        'NB & STAT',         # P
        'UNIFORM',           # Q
        'HOODY',             # R
        'SPORTSWEAR',        # S
        'AFTERSCHOOL',       # T
        'EXCURSION',         # U
        'EXTERNAL EXAMINATION',  # V
        'GRADUATION/ SPORTS/PARTY',  # W
        'HOLIDAY LESSON',    # X
        'OUTSTANDING',       # Y
        'TOTAL FEE',         # Z
        'TOTAL PAID',        # AA
        'BALANCE'            # AB
    ]
    
    # Also keep these for reference but not in details
    # A: LAST NAME, B: FIRST NAME, C: CLASS, D: DISCOUNT, E: AD-FORM
    # F: FULL TUITION, G: PTA/EXAM, H: REPORT CARD, I: PRACTICALS
    
    # Generate codes
    df['code'] = df.apply(
        lambda row: generate_code(row['LAST NAME'], row['FIRST NAME'], row.name),
        axis=1
    )
    
    # Build student data
    data = {}
    for _, row in df.iterrows():
        # Get clean name
        name = f"{row['LAST NAME']} {row['FIRST NAME']}".strip()
        
        # Convert numeric fields
        def safe_float(val):
            try:
                return float(val) if val and val != '' else 0
            except:
                return 0
        
        # Build details from columns J through AB only
        details = {}
        for col in fee_columns:
            if col in row.index and str(row[col]).strip():
                val = row[col]
                # Try to format as number if it's numeric
                try:
                    if isinstance(val, (int, float)):
                        details[col] = str(val)
                    else:
                        details[col] = str(val)
                except:
                    details[col] = str(val)
        
        # Add basic info to details (for reference)
        details['LAST NAME'] = str(row['LAST NAME'])
        details['FIRST NAME'] = str(row['FIRST NAME'])
        details['CLASS'] = str(row['CLASS'])
        
        student = {
            'name': name,
            'class': row['CLASS'],
            'total_fee': safe_float(row.get('TOTAL FEE', 0)),
            'total_paid': safe_float(row.get('TOTAL PAID', 0)),
            'balance': safe_float(row.get('BALANCE', 0)),
            'details': details
        }
        
        data[row['code']] = student
    
    # Save students.json
    with open('students.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    # Generate parent codes for reference
    codes_df = pd.DataFrame({
        'Code': df['code'],
        'Student': df.apply(lambda r: f"{r['LAST NAME']} {r['FIRST NAME']}".strip(), axis=1),
        'Class': df['CLASS']
    }).sort_values('Code')
    
    codes_df.to_excel('parent_codes.xlsx', index=False)
    
    # Print summary
    print(f"✅ Updated {len(data)} students on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"📊 Total students: {len(data)}")
    print(f"📋 Fee columns: {len(fee_columns)} columns")
    print(f"🔑 Codes saved to parent_codes.xlsx")
    
    # Show sample
    if data:
        sample_code = list(data.keys())[0]
        sample = data[sample_code]
        print(f"\n📌 Sample: {sample_code} - {sample['name']}")
        print(f"   Total Fee: ₦{sample['total_fee']:,.2f}")
        print(f"   Paid: ₦{sample['total_paid']:,.2f}")
        print(f"   Balance: ₦{sample['balance']:,.2f}")
        print(f"   Details: {len(sample['details'])} fields")
    
    return data

if __name__ == "__main__":
    update_school_data()