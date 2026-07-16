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
    
    # Get all column names from the dataframe
    all_columns = list(df.columns)
    print(f"📋 Found columns: {all_columns}")
    
    # Define the exact columns J through AB by position
    # Since pandas uses 0-based indexing, column J is index 9, AB is index 27
    # J=9, K=10, L=11, M=12, N=13, O=14, P=15, Q=16, R=17, S=18
    # T=19, U=20, V=21, W=22, X=23, Y=24, Z=25, AA=26, AB=27
    
    # Get the actual column names at positions 9 through 27
    fee_columns = all_columns[9:28]  # This gets columns J through AB (9 to 27 inclusive)
    
    print(f"📊 Fee columns (J through AB):")
    for i, col in enumerate(fee_columns):
        print(f"   Position {chr(74+i)} ({9+i}): '{col}'")
    
    # Generate codes
    df['code'] = df.apply(
        lambda row: generate_code(row['LAST NAME'], row['FIRST NAME'], row.name),
        axis=1
    )
    
    # Build student data
    data = {}
    for idx, row in df.iterrows():
        # Get clean name
        name = f"{row['LAST NAME']} {row['FIRST NAME']}".strip()
        
        # Convert numeric fields
        def safe_float(val):
            try:
                return float(val) if val and val != '' else 0
            except:
                return 0
        
        # Build details from columns J through AB ONLY
        details = {}
        for col in fee_columns:
            if col in row.index:
                val = row[col]
                # Store all values, even if they're 0 or empty
                if str(val).strip():
                    details[col] = str(val)
                else:
                    details[col] = '0'
        
        # Get totals from the sheet
        total_fee = safe_float(row.get('TOTAL FEE', 0))
        total_paid = safe_float(row.get('TOTAL PAID', 0))
        balance = safe_float(row.get('BALANCE', 0))
        
        student = {
            'name': name,
            'class': row['CLASS'],
            'total_fee': total_fee,
            'total_paid': total_paid,
            'balance': balance,
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
    print(f"\n✅ Updated {len(data)} students on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"📊 Total students: {len(data)}")
    print(f"📋 Fee columns: {len(fee_columns)} columns (J through AB)")
    print(f"🔑 Codes saved to parent_codes.xlsx")
    
    # Show sample with all fee columns
    if data:
        sample_code = list(data.keys())[0]
        sample = data[sample_code]
        print(f"\n📌 Sample: {sample_code} - {sample['name']}")
        print(f"   Class: {sample['class']}")
        print(f"   Total Fee: ₦{sample['total_fee']:,.2f}")
        print(f"   Paid: ₦{sample['total_paid']:,.2f}")
        print(f"   Balance: ₦{sample['balance']:,.2f}")
        print(f"\n   Fee Details (J through AB):")
        for col in fee_columns:
            val = sample['details'].get(col, '0')
            if val != '0':
                try:
                    num_val = float(val)
                    print(f"     {col}: ₦{num_val:,.2f}")
                except:
                    print(f"     {col}: {val}")
    
    return data

if __name__ == "__main__":
    update_school_data()