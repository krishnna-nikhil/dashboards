import pandas as pd

# Path to the Excel file
file_path = "/Users/devnikhil/Downloads/Feb _ March _ April _  Operation Master Sheet(Expenses )-2.xlsx"

# Load the Excel file into a DataFrame
df = pd.read_excel(file_path)

df.info()

print("Column Names in DataFrame:")
print(df.columns)

df.columns = df.columns.str.strip()

# Now, retrieve unique dates from the cleaned "Date" column
unique_dates = df['Date'].unique()

# Print unique dates
print("Unique Dates:")
print(unique_dates)


df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce')

# Clean the 'Amount' column
df['Amount'] = (
    df['Amount']
    .astype(str)
    .str.replace('₹', '', regex=False)
    .str.replace('?', '', regex=False)
    .str.replace(',', '')
    .str.strip()
)

# Convert to numeric, forcing errors to NaN
df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')

# Drop rows where Amount is NaN (could also fill with 0 if needed)
df = df.dropna(subset=['Amount'])

# Group by date and sum
date_wise_expense = df.groupby('Date')['Amount'].sum().reset_index()

# Display results
print("Date-wise Expense:")
print(date_wise_expense)

# Convert the 'Region' column to lowercase
df['City Name'] = df['City Name'].str.lower()

# Get unique regions after standardization
unique_regions_standardized = df['City Name'].unique()

# Convert to a list for easier reading
unique_regions_standardized_list = unique_regions_standardized.tolist()

# Print the standardized unique regions
print("Standardized Unique Regions:")
print(unique_regions_standardized_list)

# Total amount by Region
total_by_region = df.groupby('City Name')['Amount'].sum().reset_index()
print("Total Amount by Region:")
print(total_by_region)

# Total amount by Purpose (Product)
df['Category'] = df['Category'].str.lower().str.strip()
total_by_purpose = df.groupby('Category')['Amount'].sum().reset_index()
print("\nTotal Amount by Purpose:")
print(total_by_purpose)

df['Department'] = df['Department'].str.lower().str.strip()
total_by_department = df.groupby('Department')['Amount'].sum().reset_index()
print("\nTotal Amount by Department:")
print(total_by_department)

df['Project'] = df['Project'].str.lower().str.strip()
total_by_purpose = df.groupby('Project')['Amount'].sum().reset_index()
print("\nTotal Amount by Purpose:")
print(total_by_purpose)


# Replace 'YourDataFrame' with the name of your DataFrame
strawberry_products = ['STARWBERRY PCK PC', 'Strawberry', 'Strawberry Punnet PKD PC']

# Filter and calculate day-wise totals for the specified products
day_wise_totals = df[df['Project'].isin(strawberry_products)].groupby('Date')['Amount'].sum()

# Display the result
print(day_wise_totals)

# Calculate grouped summaries
total_by_region = df.groupby('City Name')['Amount'].sum().reset_index()
total_by_purpose_category = df.groupby('Category')['Amount'].sum().reset_index()
total_by_department = df.groupby('Department')['Amount'].sum().reset_index()
total_by_purpose_project = df.groupby('Project')['Amount'].sum().reset_index()

# Define the output file path
output_file = '/Users/devnikhil/Downloads/expenses_maymonth.xlsx'

# Save results to separate sheets in the same Excel file
with pd.ExcelWriter(output_file) as writer:
    total_by_region.to_excel(writer, sheet_name='Total by Region', index=False)
    total_by_purpose_category.to_excel(writer, sheet_name='Total by Category', index=False)
    total_by_department.to_excel(writer, sheet_name='Total by Department', index=False)
    total_by_purpose_project.to_excel(writer, sheet_name='Total by Project', index=False)

print(f"Results saved to {output_file}")
