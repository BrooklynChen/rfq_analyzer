# Calculate maximum WON amount of each RFQ
import pandas as pd

def add_won_amount(merged_df, rfq_q_submit):
    
    def calculate_max_amount(df):
        pd.options.display.float_format = '{:.3f}'.format

        df['Q1P1'] = df['Qty'] * df['PiecePrice']
        df['Q2P2'] = df['Qty2'] * df['PiecePrice2']
        df['Q3P3'] = df['Qty3'] * df['PiecePrice3']
        df['Q4P4'] = df['Qty4'] * df['PiecePrice4']
        df['Q5P5'] = df['Qty5'] * df['PiecePrice5']
        df['Q1P6'] = df['Qty'] * df['PiecePrice6']
        df['Q2P7'] = df['Qty2'] * df['PiecePrice7']
        df['Q3P8'] = df['Qty3'] * df['PiecePrice8']
        df['Q4P9'] = df['Qty4'] * df['PiecePrice9']
        df['Q5P10'] = df['Qty5'] * df['PiecePrice10']
        df['Q1P11'] = df['Qty'] * df['PiecePrice11']
        df['Q2P12'] = df['Qty2'] * df['PiecePrice12']
        df['Q3P13'] = df['Qty3'] * df['PiecePrice13']
        df['Q4P14'] = df['Qty4'] * df['PiecePrice14']
        df['Q5P15'] = df['Qty5'] * df['PiecePrice15']

        df.loc[:, 'Max_Amount'] = df[['Q1P1', 'Q2P2', 'Q3P3', 'Q4P4', 'Q5P5', 
                                        'Q1P6', 'Q2P7', 'Q3P8', 'Q4P9', 
                                        'Q5P10', 'Q1P11', 'Q2P12', 'Q3P13', 
                                        'Q4P14', 'Q5P15']].max(axis=1)

        df['Max_Amount'] = df['Max_Amount'].round(3)
        df = df.groupby(['RFQID']).sum().reset_index()
        rfq_q_submit_order = ['RFQID', 'Max_Amount']
        df = df[rfq_q_submit_order]
        return df

    def explode_factories(row):
        # Ensure 'FactoryUsed' is a string first
        if isinstance(row['FactoryUsed'], str):
            factories = row['FactoryUsed'].split(';')  # If it's a string, split it
        else:
            factories = row['FactoryUsed']  # If it's already a list, use it directly

        # Adjust the 'Max_Amount' accordingly if there are multiple factories
        num_factories = len(factories)
        if num_factories > 1:
            max_amounts = [row['Max_Amount'] / num_factories] * num_factories
        else:
            max_amounts = [row['Max_Amount']]

        # Create new rows with each factory and corresponding amount
        exploded_rows = pd.DataFrame({
            'RFQID': [row['RFQID']] * num_factories,
            'Customer': [row['Customer']] * num_factories,
            'Year': [row['Year']] * num_factories,
            'FactoryUsed': factories,
            'Sales Rep': [row['Sales Rep']] * num_factories,
            'Max_Amount': max_amounts
        })
        return exploded_rows
    

    rfq_q_submit = calculate_max_amount(rfq_q_submit)

    rfq_won = merged_df[merged_df['ResultCode'] == 'Won']
    rfq_won_order = ['RFQID', 'Customer', 'Year', 'FactoryUsed', 'Sales Rep']
    rfq_won = rfq_won[rfq_won_order]
    rfq_won = rfq_won.merge(rfq_q_submit, on='RFQID', how='left')

    # Apply the function and concatenate the results
    rfq_won = pd.concat(rfq_won.apply(explode_factories, axis=1).tolist(), ignore_index=True)
    rfq_won = rfq_won.groupby(['RFQID', 'Customer', 'Year', 'FactoryUsed', 'Sales Rep']).last().reset_index()
    rfq_won.rename(columns={'FactoryUsed': 'Factory', 'Max_Amount': 'Total Amount'}, inplace=True)
    return rfq_won
