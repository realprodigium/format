import pandas as pd

def inspect_file(filename):
    print(f"\n--- Inspecting {filename} ---")
    try:
        df = pd.read_excel(filename, header=None)
        print("Rows 10 to 25:")
        print(df.iloc[10:25])
        
        # Look for the header row definitively
        target_keywords = ['fecha', 'descripcion', 'valor', 'saldo']
        for i, row in df.iterrows():
            row_values = [str(x).lower().strip() for x in row.values if not pd.isna(x)]
            # If at least 3 keywords are in the row
            matches = sum(1 for kw in target_keywords if any(kw in rv for rv in row_values))
            if matches >= 2:
                print(f"Definitive header found at row {i}")
                print(f"Columns: {list(row.values)}")
                break
    except Exception as e:
        print(f"Error reading {filename}: {e}")

inspect_file("78100004719_MAR2026.xlsx")
