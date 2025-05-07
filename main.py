#%%
from datetime import datetime
import os

from data.mock_generator import generate_rfq_data, generate_rfq_to_factory, generate_rfq_quote_submit
from processing.clean_rfq import clean_rfq_1, clean_rfq_2
from processing.calculate_won_amount import add_won_amount
from processing.factory_analysis import factory_analysis
from outputs.format_factory_excel import format_factory_excel
from processing.customer_analysis import customer_analysis
from outputs.format_customer_excel import format_customer_excel
from processing.customer_factory_analysis import customer_factory_analysis
from processing.sales_rep_analysis import sales_analysis
from processing.ml_model import win_rate_model


def clean_rfq_fac(rfq_fac):
    rfq_fac = rfq_fac.sort_values(by=['RFQID'])
    rfq_fac = rfq_fac.fillna('Blanks')
    rfq_fac = rfq_fac[rfq_fac['FactorySubmittedQuote'] != 'Blanks']
    rfq_fac = rfq_fac.groupby(['RFQID', 'FactorySubmittedQuote']).last().reset_index()
    return rfq_fac

def add_FactoryQuoteRespone(df, rfq_fac):
    df = df.merge(rfq_fac, on ='RFQID', how='left')
    merged_df_order =['RFQID', 'Customer', 'Year', 'Sales Rep', 'FactoryUsed', 'ResultCode', 'FactorySubmittedQuote', 'FactoryQuoteRespone']
    df = df[merged_df_order]

    # Split the rows where there's a comma and then explode the resulting lists into separate rows
    df['FactoryUsed'] = df['FactoryUsed'].str.split(';')

    # Explode the lists into individual rows
    df = df.explode('FactoryUsed', ignore_index=True)
    df = df.fillna('Blanks')
    return df



# Generate all mock datasets
rfq = generate_rfq_data()
rfq_fac = generate_rfq_to_factory()
rfq_q_submit = generate_rfq_quote_submit()


rfq.to_excel('data/rfq.xlsx', index=False)
rfq_fac.to_excel('data/rfq_fac.xlsx', index=False)
rfq_q_submit.to_excel('data/rfq_q_submit.xlsx', index=False)

# Clean rfq
rfq = clean_rfq_1(rfq)

# Update the sales rep for each customer
customer_to_sales = rfq.copy()
customer_to_sales_column = ['Customer', 'Sales Rep']
customer_to_sales = customer_to_sales[customer_to_sales_column]
customer_to_sales = customer_to_sales.groupby('Customer').last().reset_index()
customer_to_sales = customer_to_sales.fillna('Blanks Blanks')

rfq = clean_rfq_2(rfq, customer_to_sales)
rfq.to_excel('outputs/raw/RFQ_clean.xlsx', index=False)

# Clean rfq_fac to extract FactoryQuoteRespone
rfq_fac = clean_rfq_fac(rfq_fac)

# Add FactoryQuoteRespone to rfq
rfq2 = rfq.copy()
merged_df = add_FactoryQuoteRespone(rfq2, rfq_fac)

# Add max won amount
rfq_won = add_won_amount(merged_df, rfq_q_submit)
rfq_won.to_excel('outputs/raw/Awarded_Amount_List.xlsx', index=False)

# RFQ analysis by factory
factory_ana = factory_analysis(merged_df)

# Save RFQ analysis by factory
current_date = datetime.now().strftime("%m.%d.%y")
factory_ana_name1 = os.path.join("outputs/raw", "RFQ_Analysis_by_Factory.xlsx")
factory_ana_name2 = os.path.join("outputs/formatted", f"RFQ_Analysis_by_Factory_{current_date}.xlsx")
factory_ana.to_excel(factory_ana_name1, index=False, float_format='%.3f')

# Formate excel file
format_factory_excel(factory_ana_name1, factory_ana_name2, rfq_won)

# RFQ analysis by customer
cus_subm_result_1 = customer_analysis(merged_df, customer_to_sales, rfq_won)
pivot_cus_name1 = os.path.join("outputs/raw", 'RFQ_Analysis_by_Customer.xlsx')
pivot_cus_name2 = os.path.join("outputs/formatted", f'RFQ_Analysis_by_Customer_{current_date}.xlsx')
cus_subm_result_1.to_excel(pivot_cus_name1, index=False, float_format='%.3f')

cus_number_by_sales_by_year = cus_subm_result_1.copy()

cus_list_ = ['Customer']
customer_list = cus_subm_result_1[cus_list_]
customer_list = customer_list.drop_duplicates(subset='Customer', keep='last')
customer_list.to_excel('outputs/raw/Customer_List.xlsx', index=False)

# Formate excel file
format_customer_excel(pivot_cus_name1, pivot_cus_name2)

# Customer and used factory analysis
cus_n_fac = customer_factory_analysis(merged_df, customer_to_sales, rfq_won)

# RFQ analysis by sales rep
sales_analysis(merged_df, rfq_won, cus_number_by_sales_by_year, customer_to_sales)

# Win Rate Predictions
win_rate_model(cus_n_fac)
# %%
