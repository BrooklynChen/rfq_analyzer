def customer_factory_analysis(merged_df, customer_to_sales, rfq_won):

    # Factory used and the result by customer
    cus_n_fac = merged_df.groupby(['RFQID', 'Customer', 'Year', 'FactoryUsed', 'ResultCode']).last().reset_index()
    cus_n_fac = cus_n_fac.groupby(['Customer', 'Year', 'FactoryUsed', 'ResultCode']).size().reset_index(name='Count')

    # Pivot the DataFrame, with 'FactoryUsed' as columns and 'ResultCode' as another dimension in the columns
    cus_n_fac = cus_n_fac.pivot_table(index=['Customer', 'Year', 'FactoryUsed'], 
                                columns=['ResultCode'], 
                                values='Count', 
                                aggfunc='sum', 
                                fill_value=0)
    cus_n_fac = cus_n_fac.reset_index()

    # Total quoted number
    cus_n_fac['Total Quoted'] = cus_n_fac['Blanks']+cus_n_fac['Lost']+cus_n_fac['Pending']+cus_n_fac['Won']
    cus_n_fac = cus_n_fac.rename(columns={'FactoryUsed': 'Factory'})

    # Total amount
    total_amount_won_5 = rfq_won.copy()
    total_amount_won_5 = total_amount_won_5.drop(columns=['RFQID', 'Sales Rep'])
    total_amount_won_5 = total_amount_won_5.groupby(['Year', 'Customer', 'Factory']).sum().reset_index()
    cus_n_fac = cus_n_fac.merge(total_amount_won_5, on=['Year', 'Customer', 'Factory'], how='left')
    cus_n_fac['Total Amount'] = cus_n_fac['Total Amount'].fillna(0)

    # Add win rate
    cus_n_fac['Win Rate'] = cus_n_fac['Won']/cus_n_fac['Total Quoted']

    # Factory response by customer
    fac_response = merged_df.groupby(['RFQID', 'Customer', 'Year', 'FactorySubmittedQuote', 'FactoryQuoteRespone']).last().reset_index()
    fac_response = fac_response.groupby(['Customer', 'Year', 'FactorySubmittedQuote', 'FactoryQuoteRespone']).size().reset_index(name='Count')

    # Pivot the DataFrame, with 'FactoryUsed' as columns and 'ResultCode' as another dimension in the columns
    fac_response = fac_response.pivot_table(index=['Customer', 'Year', 'FactorySubmittedQuote'], 
                                columns=['FactoryQuoteRespone'], 
                                values='Count', 
                                aggfunc='sum', 
                                fill_value=0)

    fac_response = fac_response.reset_index()
    fac_response['Total Submitted RFQs'] = fac_response.loc[:, 'Advised Drop':'Quoted'].sum(axis=1)

    # Add factory response
    cus_submfac_response_ord = ['Customer', 'Year', 'FactorySubmittedQuote', 'Total Submitted RFQs']
    fac_response = fac_response[cus_submfac_response_ord]
    fac_response = fac_response.rename(columns={'FactorySubmittedQuote': 'Factory'})
    cus_n_fac = cus_n_fac.merge(fac_response, on=['Year', 'Customer', 'Factory'], how='outer')

    # Add sales rep
    cus_n_fac = cus_n_fac.merge(customer_to_sales, on='Customer', how='left')

    reset_order = ['Customer', 'Year', 'Factory', 'Sales Rep',
        'Total Submitted RFQs', 'Total Quoted', 'Total Amount',
        'Won', 'Blanks', 'Lost', 'No Quote', 'Pending', 'Win Rate']
    cus_n_fac = cus_n_fac[reset_order]
    cus_n_fac = cus_n_fac.fillna(0)

    cus_n_fac.to_excel('outputs/raw/Customer_UsedFactory_Result.xlsx', index=False)
    return cus_n_fac
