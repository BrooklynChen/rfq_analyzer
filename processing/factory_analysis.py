def factory_analysis(merged_df):
    # FactoryUsed and Result
    def fac_result(merged_df):
        fac_used_result = merged_df.groupby(['RFQID','FactoryUsed']).last().reset_index()
        fac_used_result = fac_used_result.groupby(['Year', 'FactoryUsed', 'ResultCode']).size().reset_index(name='Count')

        # Pivot the dataframe so that 'ResultCode' values become columns
        fac_used_result = fac_used_result.pivot_table(index=['FactoryUsed', 'Year'], 
                                        columns='ResultCode', 
                                        values='Count', 
                                        aggfunc='sum', 
                                        fill_value=0)

        # Flatten the columns to a more readable format
        fac_used_result.columns = [f'{col}' for col in fac_used_result.columns]
        fac_used_result['Total Used'] = fac_used_result.sum(axis=1)

        # Reset the index to make 'Factory' and 'Year' regular columns
        fac_used_result = fac_used_result.reset_index()
        fac_used_result = fac_used_result.rename(columns={'FactoryUsed':'Factory'})
        return fac_used_result

    # FactorySubmitted and Responses
    def fac_response(merged_df):
        fac_subm_response = merged_df.groupby(['RFQID','FactorySubmittedQuote']).last().reset_index()

        # Group by 'FactorySubmittedQuote', 'Year', and 'FactoryQuoteRespone' and count the number of rows
        fac_subm_response = fac_subm_response.groupby(['FactorySubmittedQuote', 'Year', 'FactoryQuoteRespone']).size().reset_index(name='Count')

        # Pivot the dataframe so that 'FactoryQuoteRespone' values become columns
        fac_subm_response = fac_subm_response.pivot_table(index=['FactorySubmittedQuote', 'Year'], 
                                        columns='FactoryQuoteRespone', 
                                        values='Count', 
                                        aggfunc='sum', 
                                        fill_value=0)

        fac_subm_response.columns = [f'{col}' for col in fac_subm_response.columns]
        fac_subm_response['Total Submitted Quote'] = fac_subm_response.sum(axis=1)

        # Reset the index to make 'Factory' and 'Year' regular columns
        fac_subm_response = fac_subm_response.reset_index()
        fac_subm_response = fac_subm_response.rename(columns={'FactorySubmittedQuote':'Factory'})
        return fac_subm_response

    # Merge
    fac_used_result = fac_result(merged_df)
    fac_subm_response = fac_response(merged_df)
    factory_ana = fac_subm_response.merge(fac_used_result, on=['Factory', 'Year'], how='outer')

    new_order = ['Factory', 'Year', 'Total Submitted Quote', 'Quoted', 'No Quote_x', 'No Response', 'Advised Drop',
                'Estimate', 'Closed', 'Blanks_x', 'Total Used', 'Won', 'Lost', 'Pending', 'No Quote_y', 'Blanks_y']

    factory_ana = factory_ana[new_order]
    factory_ana = factory_ana.fillna(0)
    return factory_ana