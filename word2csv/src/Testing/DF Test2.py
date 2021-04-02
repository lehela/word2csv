import pandas as pd

df = pd.DataFrame(columns=['Product', 'Price', 'Buy', 'Sell'])
df.loc[len(df.index)] = ["Apple", 1.50, 3, 2]
df.loc[len(df.index)] = ["Banana", 0.75, -8, 4]
df.loc[len(df.index)] = ["Carrot", 2.00, -6, -3]
df.loc[len(df.index)] = ["Blueberry", 0.05, 5, 6]

df['Ratio'] = df.apply(
    lambda x: (x.Price / x.Sell) if abs(x.Buy) < abs(x.Sell) else (x.Price / x.Buy),
    axis=1)

print('Done')