import pandas as pd
import random
from datetime import datetime, timedelta
import numpy as np
from faker import Faker

fake = Faker()

# Generate mock data for the RFQ table
def generate_rfq_data(n=500):
    df = pd.DataFrame([{
        'RFQID': f'RFQ{i}',
        'Customer': random.choice([f'Customer{i}' for i in range(1, 26)]),
        'RFQDate': fake.date_between(start_date='-5y', end_date='today'),
        'FactoryUsed': random.choice(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'A;B', 'E;G', 'B;D']),
        'AssignedSalesPerson': random.choice([
            'Randy Byerly', 'Robyn George', 'Alfred Ludwig',
            'James Aaron', 'Heath Mellen', 'Laura Jackson', 'David Tarver']),
        'Material': random.choice(['Fabric', 'Metal', 'Rubber', 'Plastic', 'Wire']),
        'ResultCode': random.choice(['Lost', 'Won', 'Pending', 'No Quote', np.nan]),
        'TotalAmount': round(random.uniform(1000, 100000), 2),
        'InquiryType': random.choice(['New Project', 'Price Update']),
        'Email': fake.email(),  # ✅ Fixed here
        'Quoted': random.choice([0,1])
    } for i in range(n)])
    df['RFQ Due Date'] = df['RFQDate'] + timedelta(days=10)
    df['FQUDate'] = df['RFQ Due Date'] + timedelta(days=3)
    return df

# Generate mock data for the RFQtoFactory table
def generate_rfq_to_factory(n=800):
    return pd.DataFrame([{
        'RFQID': f'RFQ{random.randint(0, 500)}',  # Foreign key to RFQ table
        'FactorySubmittedQuote': random.choice(['A', 'B', 'C', 'D', 'E', 'F', 'G', np.nan]),
        'FactoryQuoteRespone': random.choice(['Quoted', 'No Response', 'No Quote', 'Advised Drop', 'Estimate', 'Closed', np.nan])
    } for _ in range(n)])

# Generate mock data for the RFQQuoteSubmit table
def generate_rfq_quote_submit(n=700):
    return pd.DataFrame([{
        'RFQID': f'RFQ{random.randint(0, n)}',
        'PartNumber':f'PN{random.randint(1000, 9999)}',
        'Qty': random.randint(1, 10000),
        'Qty2': random.randint(1, 10000),
        'Qty3': random.randint(1, 10000),
        'Qty4': random.randint(1, 10000),
        'Qty5': random.randint(1, 10000),
        'PiecePrice': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice2': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice3': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice4': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice5': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice6': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice7': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice8': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice9': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice10': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice11': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice12': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice13': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice14': round(random.uniform(1.0, 10.0), 2),
        'PiecePrice15': round(random.uniform(1.0, 10.0), 2),
    } for _ in range(n)])

