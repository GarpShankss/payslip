import pandas as pd

df = pd.read_excel('PAY.xlsx')
print('Columns:', list(df.columns))
print('\nFirst row:')
print(df.iloc[0].to_dict())
print('\nTotal rows:', len(df))
