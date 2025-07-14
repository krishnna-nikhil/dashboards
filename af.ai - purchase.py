#!/usr/bin/env python
# coding: utf-8

# In[28]:


import pandas as pd

# Path to the Excel file
file_path = "/Users/devnikhil/Downloads/Bill (17).xlsx"

# Load the Excel file into a DataFrame
df = pd.read_excel(file_path)
df.info()


# In[29]:


df.head()


# In[30]:


print(df['Bill Date'].sort_values().unique())


# In[31]:


df['Item Name'] = df['Item Name'].str.strip().str.lower()
df['Line Item Location Name'] = df['Line Item Location Name'].str.strip().str.lower()
df_grouped = df.groupby('Item Name')['Item Total'].sum().sort_values(ascending=False)
print(df_grouped)

df_ware = df.groupby('Line Item Location Name')['Item Total'].sum().sort_values(ascending=False)
print(df_ware)


# In[32]:


import pandas as pd

# Define file path
file_path = "/Users/devnikhil/Downloads/may_full_pur.xlsx"

# Save both dataframes in different sheets of the same file
with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
    df_grouped.to_frame().to_excel(writer, sheet_name="Itemwise_Total")
    df_ware.to_frame().to_excel(writer, sheet_name="Warehousewise_Total")

print(f"File saved successfully: {file_path}")


# In[ ]:





# In[ ]:




