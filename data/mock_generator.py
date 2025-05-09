import pandas as pd
import random
from datetime import timedelta
import numpy as np
from faker import Faker

fake = Faker()

# Generate mock data for the RFQ table
def generate_rfq_data(n=700):
    df = pd.DataFrame([{
        'RFQID': f'RFQ{i}',
        'Customer': random.choice([f'Cust_{i}' for i in range(1, 20)]),
        'RFQDate': fake.date_between(start_date='-5y', end_date='today'),
        'FactoryUsed': random.choices(['A', 'B', 'C', 'D', 'E', 'A;B', np.nan], weights = [0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.4], k=1)[0],
        'AssignedSalesPerson': random.choice([
            'Randy Byerly', 'Robyn George', 'Alfred Ludwig',
            'James Aaron', 'Heath Mellen', 'Laura Jackson', 'David Tarver']),
        'Material': random.choice(['Fabric', 'Metal', 'Rubber', 'Plastic', 'Wire']),
        'ResultCode': random.choices(['Lost', 'Won', 'Pending', 'No Quote', np.nan], weights = [0.1, 0.6, 0.1, 0.1, 0.1], k=1)[0],
        'TotalAmount': round(random.uniform(1000, 100000), 2),
        'InquiryType': random.choices(['New Project', 'Price Update'], weights = [0.90, 0.10], k=1)[0],
        'Email': fake.email(),
        'Quoted': random.choice([0, 1])
    } for i in range(n)])
    df['RFQ Due Date'] = df['RFQDate'] + timedelta(days=10)
    df['FQUDate'] = df['RFQ Due Date'] + timedelta(days=3)
    return df

# Generate mock data for the RFQtoFactory table
def generate_rfq_to_factory(n=2000):
    return pd.DataFrame([{
        'RFQID': f'RFQ{random.randint(0, 699)}',  # Ensure RFQID matches the range from 0 to 699
        'FactorySubmittedQuote': random.choices(['A', 'B', 'C', 'D', 'E'], weights = [0.2, 0.20, 0.20, 0.10, 0.30], k=1)[0],
        'FactoryQuoteRespone': random.choices(['Quoted', 'No Response', 'No Quote', 'Advised Drop', 'Estimate', 'Closed', np.nan],
                                              weights=[0.94, 0.01, 0.01, 0.01, 0.01, 0.01, 0.01], k=1)[0],
    } for _ in range(n)])

# Generate mock data for the RFQQuoteSubmit table
def generate_rfq_quote_submit(n=700):
    return pd.DataFrame([{
        'RFQID': f'RFQ{i}',
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
    } for i in range(n)])

