import pandas as pd
import io
import os

# Copy the logic from main.py to test it
def clean_currency(value):
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    
    s_value = str(value).replace('$', '').strip()
    if not s_value or s_value == '.00':
        return 0.0

    last_dot = s_value.rfind('.')
    last_comma = s_value.rfind(',')
    
    if last_dot > last_comma:
        s_value = s_value.replace(',', '')
    elif last_comma > last_dot:
        s_value = s_value.replace('.', '').replace(',', '.')
    else:
        s_value = s_value.replace(',', '.')

    try:
        return float(s_value)
    except ValueError:
        return 0.0

def process_test(filename):
    print(f"Testing with {filename}...")
    with open(filename, 'rb') as f:
        contents = f.read()
    
    df_raw = pd.read_excel(io.BytesIO(contents), header=None)
    header_row_index = 0
    target_keywords = ['fecha', 'descripcion', 'valor', 'saldo']
    for i, row in df_raw.iterrows():
        row_values = [str(x).lower().strip() for x in row.values if not pd.isna(x)]
        matches = sum(1 for kw in target_keywords if any(kw in rv for rv in row_values))
        if matches >= 2:
            header_row_index = i
            break
    
    print(f"Header found at row {header_row_index}")
    df = pd.read_excel(io.BytesIO(contents), header=header_row_index)
    df.columns = [str(c).strip() for c in df.columns]
    
    valor_cols = [c for c in df.columns if 'valor' in c.lower()]
    saldo_cols = [c for c in df.columns if 'saldo' in c.lower()]
    cols_to_fix = valor_cols + saldo_cols
    
    print(f"Columns to fix: {cols_to_fix}")
    for col in cols_to_fix:
        df[col] = df[col].apply(clean_currency)
    
    print("Sample of processed columns:")
    print(df[cols_to_fix].head())
    
    output_path = f"test_processed_{os.path.basename(filename)}"
    df.to_excel(output_path, index=False)
    print(f"Saved to {output_path}")

process_test("78100004719_MAR2026.xlsx")
process_test("data.xlsx")
