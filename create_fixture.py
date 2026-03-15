
import pandas as pd
import os

os.makedirs('test/fixtures', exist_ok=True)

df = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'City': ['New York', 'London', 'Paris']
})

df.to_excel('test/fixtures/simple.xlsx', index=False)
print("Created test/fixtures/simple.xlsx")
