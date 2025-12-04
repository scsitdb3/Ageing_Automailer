import os
import pyodbc
from dotenv import load_dotenv
#load_dotenv(os.path.join('venv', '.env'))
load_dotenv()  # Automatically loads .env from the root


def connect_to_db():
    server = os.getenv('Server')
    database = os.getenv('Database')  
    username = 'Utkrishtsa' 
    password = os.getenv('Password') 
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={username};"
        f"PWD={password};"
        f"MARS_Connection=Yes;"
        f"Encrypt=yes;"
        f"TrustServerCertificate=yes;"
        
    )
    conn = pyodbc.connect(conn_str, autocommit=True)
    return conn
  
   



