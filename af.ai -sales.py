#!/usr/bin/env python
# coding: utf-8

# In[72]:


import pandas as pd

# Path to the Excel file
file_path = "/Users/devnikhil/Downloads/Final_May(1).xlsx"

# Load the Excel file into a DataFrame
df = pd.read_excel(file_path)


# In[73]:


df.info()


# In[76]:


df.head()


# In[75]:


df.rename(columns={"Customer Category": "Trade Type"}, inplace=True)


# In[77]:


df['Trade Type'] = df['Trade Type'].str.strip().str.lower()

unique_gt_mt_values = df['Trade Type'].unique()

# Print the unique values
print("Unique values in the 'MT/GT' column:")
for value in unique_gt_mt_values:
    print(value)

# Optionally, count the number of unique values
print(f"\nTotal number of unique values in the 'MT/GT' column: {len(unique_gt_mt_values)}")


# In[78]:


# Filter the DataFrame for rows where 'GT/MT' column has the value 'MT'
filtered_df = df[df['Trade Type'] == 'gt']

# Group by 'Customer Name' and sum the 'Item Total', then sort in descending order
customer_totals_GT = (
    filtered_df.groupby('Customer Name')['Item Total']
    .sum()
    .sort_values(ascending=False)
)

# Display the result
print("Customer Name GT wise Item Total (descending order):")
print(customer_totals_GT)

# Save to a CSV file if needed



# In[79]:


# Filter the DataFrame for rows where 'GT/MT' column has the value 'MT'
filtered_df1 = df[df['Trade Type'] == 'mt']

# Group by 'Customer Name' and sum the 'Item Total', then sort in descending order
customer_totals_MT = (
    filtered_df1.groupby('Customer Name')['Item Total']
    .sum()
    .sort_values(ascending=False)
)

# Display the result
print("Customer Name MT wise Item Total (descending order):")
print(customer_totals_MT)

# Save to a CSV file if needed


# In[80]:


# Date-wise GRN Amount total
date_wise_grn = df.groupby('Invoice Date')['Item Total'].sum().reset_index()
df['Line Item Location Name'] = df['Line Item Location Name'].str.lower().str.strip()
# City-wise GRN Amount total
city_wise_grn = df.groupby('Line Item Location Name')['Item Total'].sum().reset_index()
# Brand Name-wise total (excluding nulls)

# Display the results
print("Date-wise GRN Amount Total:\n", date_wise_grn)
print("\nCity-wise GRN Amount Total:\n", city_wise_grn)


# In[81]:


# Pivot the data to create the desired format
pivoted_df = df.pivot_table(
    index='Invoice Date',  # Use Invoice Date as the index
    columns='Project ',   # Use unique Item Name as the columns
    values='Item Total',   # Fill the cells with the Item Total values
    aggfunc='sum',         # Aggregate Item Total values by summing them (if there are duplicates)
    fill_value=0           # Replace NaN with 0 for missing values
).reset_index()



# Display the transformed DataFrame
print(pivoted_df)

# Output the file path



# In[82]:


# Date-wise GRN Amount total
project_wise = df.groupby('Item Name')['Item Total'].sum().sort_values(ascending=False)



# Display the results
print("Date-wise GRN Amount Total:\n", project_wise)


# In[83]:


# Filter the DataFrame for rows where 'GT/MT' column has the value 'MT'
filtered_df = df[df['Trade Type'] == 'gt']

# Group by 'Customer Name' and sum the 'Item Total', then sort in descending order
customer_totals_GT = (
    filtered_df.groupby('Customer Name')['Item Total']
    .sum()
    .sort_values(ascending=False)
)

# Display the result
print("Customer Name GT wise Item Total (descending order):")
print(customer_totals_GT)

# Save to a CSV file if needed


filtered_df1 = df[df['Trade Type'] == 'mt']

# Group by 'Customer Name' and sum the 'Item Total', then sort in descending order
customer_totals_MT = (
    filtered_df1.groupby('Customer Name')['Item Total']
    .sum()
    .sort_values(ascending=False)
)

# Display the result
print("Customer Name MT wise Item Total (descending order):")
print(customer_totals_MT)

# Date-wise GRN Amount total
date_wise_grn = df.groupby('Invoice Date')['Item Total'].sum().reset_index()

# City-wise GRN Amount total
city_wise_grn = df.groupby('Line Item Location Name')['Item Total'].sum().reset_index()
# Brand Name-wise total (excluding nulls)

# Display the results
print("Date-wise GRN Amount Total:\n", date_wise_grn)
print("\nCity-wise GRN Amount Total:\n", city_wise_grn)


# Pivot the data to create the desired format
pivoted_df = df.pivot_table(
    index='Invoice Date',  # Use Invoice Date as the index
    columns='Project ',   # Use unique Item Name as the columns
    values='Item Total',   # Fill the cells with the Item Total values
    aggfunc='sum',         # Aggregate Item Total values by summing them (if there are duplicates)
    fill_value=0           # Replace NaN with 0 for missing values
).reset_index()



# Display the transformed DataFrame
print(pivoted_df)

# Date-wise GRN Amount total
project_wise = df.groupby('Item Name')['Item Total'].sum().sort_values(ascending=False)



# Display the results
print("Date-wise GRN Amount Total:\n", project_wise)

import pandas as pd

# Define the file path
file_path = "/Users/devnikhil/Downloads/fullmay.xlsx"

# Create an Excel writer object
with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
    # Save Customer Name GT wise Item Total
    customer_totals_GT.to_frame().to_excel(writer, sheet_name="Customer_GT")

    # Save Customer Name MT wise Item Total
    customer_totals_MT.to_frame().to_excel(writer, sheet_name="Customer_MT")

    # Save Date-wise GRN Amount Total
    date_wise_grn.to_excel(writer, sheet_name="Datewise_GRN", index=False)

    # Save City-wise GRN Amount Total
    city_wise_grn.to_excel(writer, sheet_name="Citywise_GRN", index=False)

    # Save Pivoted DataFrame
    pivoted_df.to_excel(writer, sheet_name="Itemwise_Pivot", index=False)

    # Save Project-wise GRN Amount Total
    project_wise.to_frame().to_excel(writer, sheet_name="Projectwise_GRN")

print(f"File saved successfully: {file_path}")


# In[ ]:





# In[ ]:




