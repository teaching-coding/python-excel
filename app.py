import pandas as pd

# Reading an Excel file into a DataFrame
df = pd.read_excel('projectManagement.xlsx', sheet_name='Sheet1')

# Displaying the first few rows
print(df.head())
