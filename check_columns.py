import pandas as pd
import sys

try:
    df = pd.read_excel('PAY.xlsx')
    print("COLUMNS IN YOUR EXCEL FILE:")
    print("="*60)
    for i, col in enumerate(df.columns, 1):
        print(f"{i}. '{col}'")
    print("\n" + "="*60)
    print(f"Total rows: {len(df)}")
    print(f"Total columns: {len(df.columns)}")
    
    print("\n\nFIRST ROW DATA:")
    print("="*60)
    if len(df) > 0:
        for col in df.columns:
            print(f"{col}: {df.iloc[0][col]}")
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)
