import pandas as pd

def inspect_file(filename):
    print(f"\n--- Inspecting {filename} ---")
    try:
        df_full = pd.read_excel(filename, header=None)
        header_row = 14
        df = pd.read_excel(filename, header=header_row)
        print("First 5 rows of actual data:")
        print(df.head())
        print("\nData types:")
        print(df.dtypes)
        print("\nExample values in 'VALOR':")
        print(df['VALOR'].head().tolist())
        print("\nExample values in 'SALDO':")
        print(df['SALDO'].head().tolist())
    except Exception as e:
        print(f"Error reading {filename}: {e}")

inspect_file("78100004719_MAR2026.xlsx")
