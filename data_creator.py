
import pandas as pd
import numpy as np
from datetime import datetime

# This makes sure we get the same random data every time we run the script.
np.random.seed(123)

# Create a list of dates for the whole year of 2024
dates = pd.date_range(start='2024-01-01', end='2024-12-31', freq='D')

# Create lists of products, sales reps, and regions
products_list = ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Headphones']
sales_reps = ['Alex', 'Sam', 'Jordan', 'Taylor']
regions_list = ['North', 'South', 'East', 'West']

# Create an empty list to hold each sales record
all_sales = []

# Loop through each day in the year and create random sales
for date in dates:
    # For each day, create between 5 and 15 sales transactions
    for _ in range(np.random.randint(5, 16)):
        product = np.random.choice(products_list)
        units = np.random.randint(1, 4) # Sell 1-3 units per transaction
        price = 0
        
        # Set a price based on the product
        if product == 'Laptop':
            price = 1000
        elif product == 'Monitor':
            price = 300
        elif product == 'Keyboard':
            price = 80
        elif product == 'Mouse':
            price = 40
        elif product == 'Headphones':
            price = 150
            
        sale_amount = units * price
        cost_amount = sale_amount * np.random.uniform(0.6, 0.8) # Cost is 60-80% of sale price
        
        # Create a sales record and add it to the list
        sale_record = {
            'Date': date,
            'Product': product,
            'SalesRep': np.random.choice(sales_reps),
            'Region': np.random.choice(regions_list),
            'UnitsSold': units,
            'SaleAmount': sale_amount,
            'CostAmount': round(cost_amount, 2)
        }
        all_sales.append(sale_record)

# Convert the list of sales into a DataFrame
sales_df = pd.DataFrame(all_sales)

# Calculate a Profit amount for each sale
sales_df['Profit'] = sales_df['SaleAmount'] - sales_df['CostAmount']

# Create a separate "Products" table
products_df = pd.DataFrame({
    'Product': products_list,
    'Category': ['Electronics', 'Accessory', 'Accessory', 'Electronics', 'Electronics']
})

# Create a separate "Sales Team" table
team_df = pd.DataFrame({
    'SalesRep': sales_reps,
    'HireDate': ['2022-01-15', '2021-08-03', '2023-02-20', '2020-11-10']
})

# Save all the data to one Excel file, but on different sheets
with pd.ExcelWriter('data_dashboard.xlsx') as writer:
    sales_df.to_excel(writer, sheet_name='Sales', index=False)
    products_df.to_excel(writer, sheet_name='Products', index=False)
    team_df.to_excel(writer, sheet_name='SalesTeam', index=False)

print("Done! The file 'data_dashboard.xlsx' has been created.")
