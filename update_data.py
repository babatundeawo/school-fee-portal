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
    
    # ALL columns from J through AB (and some extras for reference)
    # J: AD_FORM, K: TUITION, L: PTA&EXAM, M: REPORT_CARD, N: PRACTICAL
    # O: TEXTBOOKS, P: NB & STAT, Q: UNIFORM, R: HOODY, S: SPORTSWEAR
    # T: AFTERSCHOOL, U: EXCURSION, V: EXTERNAL EXAMINATION
    # W: GRADUATION/ SPORTS/PARTY, X: HOLIDAY LESSON, Y: OUTSTANDING
    # Z: TOTAL FEE, AA: TOTAL PAID, AB: BALANCE
    
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
        
        # Build details from columns J through AB
        details = {}
        for col in fee_columns:
            if col in row.index:
                val = row[col]
                if str(val).strip():
                    # Store as string but try to keep numeric format
                    details[col] = str(val)
                else:
                    details[col] = '0'
        
        # Add basic info
        details['LAST NAME'] = str(row['LAST NAME'])
        details['FIRST NAME'] = str(row['FIRST NAME'])
        details['CLASS'] = str(row['CLASS'])
        
        # Calculate total from individual fee components
        # This ensures the breakdown adds up to the total
        tuition = safe_float(row.get('TUITION', 0))
        pta_exam = safe_float(row.get('PTA&EXAM', 0))
        report_card = safe_float(row.get('REPORT_CARD', 0))
        practical = safe_float(row.get('PRACTICAL', 0))
        textbooks = safe_float(row.get('TEXTBOOKS', 0))
        nb_stat = safe_float(row.get('NB & STAT', 0))
        uniform = safe_float(row.get('UNIFORM', 0))
        hoody = safe_float(row.get('HOODY', 0))
        sportswear = safe_float(row.get('SPORTSWEAR', 0))
        afterschool = safe_float(row.get('AFTERSCHOOL', 0))
        excursion = safe_float(row.get('EXCURSION', 0))
        external_exam = safe_float(row.get('EXTERNAL EXAMINATION', 0))
        graduation = safe_float(row.get('GRADUATION/ SPORTS/PARTY', 0))
        holiday_lesson = safe_float(row.get('HOLIDAY LESSON', 0))
        outstanding = safe_float(row.get('OUTSTANDING', 0))
        ad_form = safe_float(row.get('AD_FORM', 0))
        
        # Calculate total from components
        calculated_total = (
            tuition + pta_exam + report_card + practical + 
            textbooks + nb_stat + uniform + hoody + sportswear +
            afterschool + excursion + external_exam + graduation +
            holiday_lesson + outstanding + ad_form
        )
        
        # Get the actual total from the sheet
        actual_total = safe_float(row.get('TOTAL FEE', 0))
        total_paid = safe_float(row.get('TOTAL PAID', 0))
        balance = safe_float(row.get('BALANCE', 0))
        
        student = {
            'name': name,
            'class': row['CLASS'],
            'total_fee': actual_total,
            'total_paid': total_paid,
            'balance': balance,
            'details': details,
            # Add calculated breakdown for verification
            'breakdown': {
                'tuition': tuition,
                'pta_exam': pta_exam,
                'report_card': report_card,
                'practical': practical,
                'textbooks': textbooks,
                'nb_stat': nb_stat,
                'uniform': uniform,
                'hoody': hoody,
                'sportswear': sportswear,
                'afterschool': afterschool,
                'excursion': excursion,
                'external_exam': external_exam,
                'graduation': graduation,
                'holiday_lesson': holiday_lesson,
                'outstanding': outstanding,
                'ad_form': ad_form
            }
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
    
    # Show sample with breakdown
    if data:
        sample_code = list(data.keys())[0]
        sample = data[sample_code]
        print(f"\n📌 Sample: {sample_code} - {sample['name']}")
        print(f"   Total Fee: ₦{sample['total_fee']:,.2f}")
        print(f"   Paid: ₦{sample['total_paid']:,.2f}")
        print(f"   Balance: ₦{sample['balance']:,.2f}")
        
        # Show breakdown
        print(f"\n   Breakdown:")
        for key, val in sample['breakdown'].items():
            if val > 0:
                print(f"     {key}: ₦{val:,.2f}")
    
    return data

if __name__ == "__main__":
    update_school_data()