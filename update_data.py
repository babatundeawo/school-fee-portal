import pandas as pd
import json
from datetime import datetime

def update_school_data(excel_path='/home/workdir/attachments/2025-2026 SHARED FEE.xlsx'):
    print("🔄 Processing latest fee data...")
    df = pd.read_excel(excel_path, sheet_name='3RD TERM 2025-2026 (2)')
    df = df.iloc[:-1].fillna('')  # Remove total row
    
    df['code'] = df.apply(
        lambda row: (str(row['LAST NAME'])[:3].upper() + 
                     str(row['FIRST NAME'])[:3].upper() + 
                     str(int(row.name) + 1000)).replace(' ', ''), axis=1)
    
    data = {}
    for _, row in df.iterrows():
        student = {
            'name': f"{row['LAST NAME']} {row['FIRST NAME']}".strip(),
            'class': row['CLASS'],
            'total_fee': float(row.get('TOTAL FEE', 0) or 0),
            'total_paid': float(row.get('TOTAL PAID', 0) or 0),
            'balance': float(row.get('BALANCE', 0) or 0),
            'details': {str(k).strip(): str(v) for k, v in row.items() if str(v).strip()}
        }
        data[row['code']] = student
    
    with open('students.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    # Printable codes
    codes = pd.DataFrame({
        'Code': df['code'],
        'Student': df.apply(lambda r: f"{r['LAST NAME']} {r['FIRST NAME']}", axis=1),
        'Class': df['CLASS']
    })
    codes.to_excel('parent_codes.xlsx', index=False)
    
    print(f"✅ Updated {len(data)} students on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    
if __name__ == "__main__":
    update_school_data()
