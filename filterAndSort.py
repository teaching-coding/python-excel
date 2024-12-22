import pandas as pd

# Read the Excel file
df = pd.read_excel("data.xlsx")

# Display original data
print("Original Data:")
print(df)

# Filter data
filtered_df = df.loc[(df["Age"] > 25) & (df["Department"] == "HR")]
print("\nFiltered Data (Age > 25 and Department == 'HR'):")
print(filtered_df)

# Sort data
sorted_df = df.sort_values(by="Salary", ascending=False)
print("\nData Sorted by Salary (descending):")
print(sorted_df)

# Save filtered data to a new file
filtered_df.to_excel("filtered_data.xlsx", index=False)
print("\nFiltered data saved to 'filtered_data.xlsx'")
