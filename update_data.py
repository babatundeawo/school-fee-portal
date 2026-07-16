import pandas as pd
import json
from datetime import datetime
import re

def generate_code(last_name, first_name, index):
    """Generate consistent code matching parent_codes.xlsx format"""
    # Take first 3 letters of last name and first name, uppercase
    code = (str(last_name)[:3].upper() + 
            str(first_name)[:3].upper() + 
            str(int(index) + 1000))
    # Remove any spaces or special characters
    code = re.sub(r'[^A-Z0-9]', '', code)
    return code

def update_school_data(excel_path='2025-2026 SHARED FEE.xlsx'):
    print("🔄 Processing latest fee data...")
    
    # Read the main fee data
    df = pd.read_excel(excel_path, sheet_name='3RD TERM 2025-2026 (2)')
    df = df.iloc[:-1].fillna('')  # Remove total row
    
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
        
        student = {
            'name': name,
            'class': row['CLASS'],
            'total_fee': safe_float(row.get('TOTAL FEE', 0)),
            'total_paid': safe_float(row.get('TOTAL PAID', 0)),
            'balance': safe_float(row.get('BALANCE', 0)),
            'details': {}
        }
        
        # Add all details
        for k, v in row.items():
            if str(v).strip():
                student['details'][str(k)] = str(v)
        
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
    
    print(f"✅ Updated {len(data)} students on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"📊 Total students: {len(data)}")
    print(f"🔑 Codes saved to parent_codes.xlsx")
    
    return data

if __name__ == "__main__":
    update_school_data()
