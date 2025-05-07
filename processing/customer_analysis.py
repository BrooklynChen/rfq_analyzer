def customer_analysis(merged_df, customer_to_sales, rfq_won):
    # Number of each result by customer
    cus_subm_result = merged_df.groupby(['RFQID', 'Customer', 'Year']).last().reset_index()
    cus_subm_result = cus_subm_result.groupby(['Customer', 'Year', 'ResultCode']).size().reset_index(name='Count')
    cus_subm_result = cus_subm_result.pivot_table(index=['Customer', 'Year'], 
                                columns=['ResultCode'], 
                                values='Count', 
                                aggfunc='sum', 
                                fill_value=0)

    cus_subm_result = cus_subm_result.reset_index()
    cus_subm_result.rename(columns={'FactoryUsed': 'Factory'}, inplace=True)

    # Add sales rep
    cus_subm_result = cus_subm_result.merge(customer_to_sales, on='Customer', how='left')

    # Add won amount
    cus_amount = rfq_won.copy()
    cus_amount = cus_amount.drop(columns=['RFQID', 'Factory', 'Sales Rep'])
    cus_amount = cus_amount.groupby(['Customer', 'Year']).sum().reset_index()
    cus_subm_result = cus_subm_result.merge(cus_amount, on=['Customer', 'Year'], how='left')

    cus_subm_result_1 = cus_subm_result.copy()
    cus_subm_result_order = ['Customer', 'Year', 'Sales Rep', 'Won', 'Lost', 'No Quote', 'Pending', 'Blanks']
    cus_subm_result_1 = cus_subm_result_1[cus_subm_result_order]

    # Total Submitted RFQs
    cus_total = merged_df.groupby(['RFQID']).last().reset_index()
    cus_total2 = cus_total.groupby(['Customer', 'Year']).size().reset_index(name='Total Submitted RFQs')

    # Total Quoted RFQs
    cus_total3 = cus_total[cus_total['ResultCode'] != 'No Quote']
    cus_total3 = cus_total3.groupby(['Customer', 'Year']).size().reset_index(name='Total Quoted RFQs')

    cus_subm_result_1 = cus_subm_result_1.merge(cus_total2, on=['Customer', 'Year'], how='outer')
    cus_subm_result_1 = cus_subm_result_1.merge(cus_total3, on=['Customer', 'Year'], how='outer')
    cus_subm_result_1 = cus_subm_result_1.fillna(0)

    # Add won amount
    total_amount_won_4 = rfq_won.copy()
    total_amount_won_4 = total_amount_won_4.drop(columns=['Sales Rep', 'Factory', 'RFQID'])
    total_amount_won_4 = total_amount_won_4.groupby(['Year', 'Customer']).sum().reset_index()
    cus_subm_result_1 = cus_subm_result_1.merge(total_amount_won_4, on=['Year', 'Customer'], how='left')

    # Add win rate
    cus_subm_result_1['Total Amount'] = cus_subm_result_1['Total Amount'].fillna(0)
    cus_subm_result_1['Win Rate'] = cus_subm_result_1['Won']/cus_subm_result_1['Total Quoted RFQs']
    cus_subm_result_1 = cus_subm_result_1.fillna(0)

    # Reorder
    cus_subm_result_order = ['Customer', 'Year', 'Sales Rep', 'Total Submitted RFQs', 'Total Quoted RFQs', 'Total Amount', 'Win Rate', 'Won', 'Lost', 'No Quote', 'Pending', 'Blanks']
    cus_subm_result_1 = cus_subm_result_1[cus_subm_result_order]
    return cus_subm_result_1


