import pandas as pd

def inspect_file(filename):
    print(f"\n--- Inspecting {filename} ---")
    try:
        # Try reading without header first to see what's at the top
        df = pd.read_excel(filename, header=None)
        print("First 10 rows (without header logic):")
        print(df.head(10))
        
        # Try to find where the header might be
        # Look for row that has 'Fecha', 'Valor', 'Saldo' or similar
        for i, row in df.iterrows():
            row_str = " ".join([str(x).lower() for x in row.values if not pd.isna(x)])
            if 'fecha' in row_str or 'descripcion' in row_str or 'valor' in row_str:
                print(f"Potential header found at row {i}")
                break
    except Exception as e:
        print(f"Error reading {filename}: {e}")

inspect_file("data.xlsx")
inspect_file("78100004719_MAR2026.xlsx")
