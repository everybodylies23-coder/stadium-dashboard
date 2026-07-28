import pandas as pd
import json

file_path = r"c:\Users\user\Desktop\Antigravity\データ分析自動化 - スタジアム\スタジアム_データ.xlsx"

try:
    xl = pd.ExcelFile(file_path)
    
    output = {}
    
    # Read sheets
    for sheet_name in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet_name)
        
        # Convert to string to avoid NaNs and Datetime serialization issues in json
        df = df.fillna('').astype(str)
        
        # Try to extract the first 100 rows or find relevant data
        if "6月" in sheet_name or "データ" in sheet_name:
            output[sheet_name] = df.head(50).to_dict(orient='records')
        else:
            output[sheet_name] = df.head(10).to_dict(orient='records')
            
    with open('extracted_data.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print("Extracted successfully to extracted_data.json")
except Exception as e:
    print("Error:", e)
