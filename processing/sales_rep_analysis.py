import pandas as pd
def sales_analysis(merged_df, rfq_won, cus_number_by_sales_by_year, customer_to_sales):
    
    # Calculate new customers for each Sales Rep
    def calculate_new_customers(df):
        # Initialize an empty list to store the results
        results = []

        # Loop through the years 2021 to 2025
        for year in range(2025, max(df['Year']) + 1):
            # Filter the data for the current year
            current_year_data = df[df['Year'] == year]

            # Filter the data for the previous year
            prev_year_data = df[df['Year'] == (year - 1)]

            # Get the unique customers for the current and previous years
            current_year_customers = set(current_year_data['Customer'])
            prev_year_customers = set(prev_year_data['Customer'])

            # Find customers in the current year but not in the previous year
            new_customers = current_year_customers - prev_year_customers

            # For each Sales Rep, count the number of new customers
            for rep in current_year_data['Sales Rep'].unique():
                rep_customers = current_year_data[current_year_data['Sales Rep'] == rep]['Customer']
                # Calculate how many of the customers in this rep's list are new
                new_customers_for_rep = len(set(rep_customers) & new_customers)
                results.append({'Year': year, 'Sales Rep': rep, 'New Customers': new_customers_for_rep})
        
        # Return the results as a DataFrame
        return pd.DataFrame(results)
    
    # Abbreviation function
    def abbrevName(name):
        # Split the name by spaces to get the first and last names
        names = name.split()
        # Return the initials of the first and last names in uppercase
        return f"{names[0][0]}{names[1][0]}".upper()
    
    def factory_submitted(merged_df):
        sales_submit = merged_df.groupby(['RFQID', 'FactorySubmittedQuote']).last().reset_index()
        sales_submit = sales_submit.groupby(['Sales Rep', 'Year', 'FactorySubmittedQuote','FactoryQuoteRespone']).size().reset_index(name='Count')

        # Pivot the DataFrame, with 'FactoryUsed' as columns and 'ResultCode' as another dimension in the columns
        sales_submit = sales_submit.pivot_table(index=['Sales Rep', 'Year', 'FactorySubmittedQuote'], 
                                    columns=['FactoryQuoteRespone'], 
                                    values='Count', 
                                    aggfunc='sum', 
                                    fill_value=0)
        sales_submit = sales_submit.reset_index()

    def factory_used(merged_df):
        sales_received = merged_df.groupby(['RFQID', 'FactoryUsed']).last().reset_index()
        sales_received = sales_received.groupby(['Sales Rep', 'Year', 'FactoryUsed', 'ResultCode']).size().reset_index(name='Count')

        # Pivot the DataFrame, with 'FactoryUsed' as columns and 'ResultCode' as another dimension in the columns
        sales_received = sales_received.pivot_table(index=['Sales Rep', 'Year', 'FactoryUsed'], 
                                    columns=['ResultCode'], 
                                    values='Count', 
                                    aggfunc='sum', 
                                    fill_value=0)
        sales_received = sales_received.reset_index()

    
    # Sales Rep Analysis
    sales_result0 = merged_df.groupby(['RFQID']).last().reset_index()

    # Result Code
    sales_result = sales_result0.groupby(['Sales Rep', 'Year', 'ResultCode']).size().reset_index(name='Count')

    # Pivot the dataframe so that 'Customer Result' values become columns
    sales_result = sales_result.pivot_table(index=['Sales Rep', 'Year'], 
                                    columns='ResultCode', 
                                    values='Count', 
                                    aggfunc='sum', 
                                    fill_value=0)
    sales_result = sales_result.reset_index()

    # Total Received RFQs
    sales_received2 = sales_result0.groupby(['Sales Rep', 'Year']).size().reset_index(name='Total Received RFQs')

    # Total Quoted RFQs
    sales_received3 = sales_result0.copy()
    sales_received3 = sales_received3[sales_received3['ResultCode'] != 'No Quote']
    sales_received3 = sales_received3.groupby(['Sales Rep', 'Year']).size().reset_index(name='Total Quoted RFQs')

    # Merge
    sales_ana = sales_received2.merge(sales_received3, on=['Sales Rep', 'Year'], how='outer')
    sales_ana = sales_ana.merge(sales_result, on=['Sales Rep', 'Year'], how='outer')

    # Win rate
    sales_ana['Win Rate'] = sales_ana['Won']/sales_ana['Total Quoted RFQs']

    # Total Amount
    total_amount_won_3 = rfq_won.copy()
    total_amount_won_3 = total_amount_won_3.drop(columns=['Customer', 'Factory', 'RFQID'])
    total_amount_won_3 = total_amount_won_3.groupby(['Year', 'Sales Rep']).sum().reset_index()
    sales_ana = sales_ana.merge(total_amount_won_3, on=['Year', 'Sales Rep'], how='left')
    sales_ana['Total Amount'] = sales_ana['Total Amount'].fillna(0)

    # Calculate the number of customer of the sales by year
    keep_columns = ['Customer', 'Year', 'Sales Rep']
    cus_number_by_sales_by_year = cus_number_by_sales_by_year[keep_columns]

    new_cus = cus_number_by_sales_by_year.copy()
    new_cus['Year'] = new_cus['Year'].astype(int)

    cus_number_by_sales_by_year = cus_number_by_sales_by_year.groupby(['Sales Rep', 'Year']).size().reset_index(name='Number of Customers')
    sales_ana = sales_ana.merge(cus_number_by_sales_by_year, on=['Sales Rep', 'Year'], how='left')

    # Calculate new customers
    new_customers = calculate_new_customers(new_cus)
    sales_ana = sales_ana.merge(new_customers, on=['Sales Rep', 'Year'], how='left')
    sales_ana = sales_ana.fillna(0)
    sales_ana.to_excel('outputs/raw/RFQ_Analysis_by_Sales.xlsx', index=False, float_format='%.3f')

    # Sales List
    sales_list_order = ['Sales Rep']
    sales_list = customer_to_sales[sales_list_order]
    sales_list = sales_list.drop_duplicates(subset='Sales Rep', keep='last')
    sales_list = sales_list.fillna('N A')

    # Create a new 'Sales Rep Initials' column
    sales_list['Sales Code'] = sales_list['Sales Rep'].apply(abbrevName)
    sales_list =sales_list.rename(columns={'Sales Rep': 'Sales Name'})

    sales_list.to_excel('outputs/raw/Sales_List.xlsx', index=False)