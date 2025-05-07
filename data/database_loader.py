import pandas as pd
from sqlalchemy import create_engine
import urllib
import logging

logging.basicConfig(filename='rfq_analysis.log', level=logging.DEBUG)
logging.info("Starting script execution")

# Download data from Access
db_file = 'file_path'
username = 'username'
password = 'password'

# Set up the connection string with credentials
connection_string = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    f'DBQ={db_file};'
    f'UID={username};'
    f'PWD={password};'
)

# URL encode the connection string
quoted_connection_string = urllib.parse.quote_plus(connection_string)

# Create a SQLAlchemy engine using the connection string
engine = create_engine(f'access+pyodbc:///?odbc_connect={quoted_connection_string}')

# Execute a query to retrieve data from a specific table
query_rfq = 'SELECT * FROM RFQ'
query_rfqtofactory = 'SELECT * FROM RFQtoFactory'
query_rfqquotesubmit = 'SELECT * FROM RFQQuoteSubmit'
query_leadsinfo = 'SELECT * FROM "Leads Information"'

try:
    # Read the data directly into a pandas DataFrame
    rfq = pd.read_sql(query_rfq, engine)
    rfq_fac = pd.read_sql(query_rfqtofactory, engine)
    rfq_q_submit = pd.read_sql(query_rfqquotesubmit, engine)
    leads = pd.read_sql(query_leadsinfo, engine)


except Exception as e:
    print("Error:", e)
finally:
    pass