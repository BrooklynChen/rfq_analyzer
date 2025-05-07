
rfq_fac = rfq_fac.sort_values(by=['RFQID'])
rfq_fac = rfq_fac.fillna('Blanks')
rfq_fac = rfq_fac[rfq_fac['FactorySubmittedQuote'] != 'Blanks']
rfq_fac = rfq_fac.groupby(['RFQID', 'FactorySubmittedQuote']).last().reset_index()