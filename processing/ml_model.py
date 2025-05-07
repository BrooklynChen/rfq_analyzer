import pandas as pd
import numpy as np
import joblib

def win_rate_model(cus_n_fac):
    df = cus_n_fac.copy()
    df['Total Submitted RFQs'] = pd.to_numeric(df['Total Submitted RFQs'], errors='coerce')
    df['Total Quoted'] = pd.to_numeric(df['Total Quoted'], errors='coerce')
    df['Blanks'] = pd.to_numeric(df['Blanks'], errors='coerce')
    df['Lost'] = pd.to_numeric(df['Lost'], errors='coerce')
    df['No Quote'] = pd.to_numeric(df['No Quote'], errors='coerce')
    df['Pending'] = pd.to_numeric(df['Pending'], errors='coerce')
    df['Won'] = pd.to_numeric(df['Won'], errors='coerce')
    df['Total Amount'] = pd.to_numeric(df['Total Amount'], errors='coerce')

    df = df.drop(columns=['Win Rate', 'Year', 'Sales Rep'])
    df = df.groupby(['Customer', 'Factory']).sum().reset_index()
    df['Win Rate'] = df['Won']/df['Total Quoted']
    df['Customer-Factory'] = df['Customer'].astype(str) + '-' + df['Factory'].astype(str)

    # Create customer-level and factory-level feature
    df2 = df.copy()

    customer_avg = df2.groupby('Customer').sum(numeric_only=True).reset_index()
    customer_avg['Cus_Win Rate'] = customer_avg['Won']/customer_avg['Total Quoted']
    customer_avg = customer_avg.drop(columns=['Win Rate', 'Blanks', 'Lost', 'No Quote', 'Pending'])
    customer_avg = customer_avg.rename(columns={'Total Submitted RFQs': 'Cus_Total Submitted RFQs', 'Total Quoted':'Cus_Total Quoted', 'Total Amount':'Cus_Total Amount', 'Won':'Cus_Won'})

    factory_avg = df2.groupby('Factory').sum(numeric_only=True).reset_index()
    factory_avg['Fac_Win Rate'] = factory_avg['Won']/factory_avg['Total Quoted']
    factory_avg = factory_avg.drop(columns=['Win Rate', 'Blanks', 'Lost', 'No Quote', 'Pending'])
    factory_avg = factory_avg.rename(columns={'Total Submitted RFQs': 'Fac_Total Submitted RFQs', 'Total Quoted':'Fac_Total Quoted', 'Total Amount':'Fac_Total Amount', 'Won':'Fac_Won'})

    # Merge these averages back into the original data
    df = df.merge(customer_avg, on='Customer')
    df = df.merge(factory_avg, on='Factory')

    df['Cus_Win Rate1'] = np.log1p(df['Cus_Win Rate'])
    df = df.fillna(0)

    rf_best = joblib.load(r'model\best_model.pkl')

    X = df[['Won', 'Cus_Win Rate1', 'Cus_Won', 'Fac_Won', 'Fac_Win Rate', 'Total Submitted RFQs']]

    log_pred_all = rf_best.predict(X)

    y_pred_all = np.expm1(log_pred_all)

    df_new = df.copy()
    df_new['Actual Win Rate'] = df['Win Rate']
    df_new['Predicted Win Rate'] = y_pred_all
    df_new.to_excel('outputs/raw/Win_Rate_Predictions.xlsx', index=False)