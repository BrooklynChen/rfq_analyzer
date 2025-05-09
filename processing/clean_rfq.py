import pandas as pd
import numpy as np

# Clean RFQ file
def clean_rfq_1(df):
    df = df.rename(columns={'RFQDate':'Original RFQ Date', 'FQUDate':'Original Quote Date', 'AssignedSalesPerson':'Sales Rep'})
    df = df.sort_values(by=['RFQID'])
    df = df.replace({None: np.nan, '': np.nan})
    df = df.apply(lambda x: np.nan if isinstance(x, str) and x.strip() == '' else x)
    df['Customer'] = df['Customer'].str.upper()

    # Add column 'Year'
    df['Original RFQ Date'] = pd.to_datetime(df['Original RFQ Date'])
    df['Original Quote Date'] = pd.to_datetime(df['Original Quote Date'])
    df['Year'] = df['Original RFQ Date'].fillna(df['Original Quote Date'] - pd.Timedelta(days=5))
    df['Year'] = df['Year'].fillna(df['RFQ Due Date'] - pd.Timedelta(days=5))
    df['Year'] = df['Year'].dt.year
    df['FactoryUsed'] = df['FactoryUsed'].fillna('Blanks')
    df['ResultCode'] = df['ResultCode'].fillna('Blanks')
    df = df.drop(columns=['Original Quote Date', 'Material', 'Email', 'Original RFQ Date', 'RFQ Due Date', 'TotalAmount'])
    return df


# Add new sales rep to the RFQ table
def clean_rfq_2(df, customer_to_sales):
    df = df.drop(columns=['Sales Rep'])
    df = df.merge(customer_to_sales, on='Customer', how='left')

    # Filter the dataset
    df = df[df['InquiryType'] == 'New Project']
    df = df[df['Year'] > 2019]
    df = df[(df['Quoted'] == True)]
    df['Sales Rep'] = df['Sales Rep'].fillna('Blanks Blanks')
    return df