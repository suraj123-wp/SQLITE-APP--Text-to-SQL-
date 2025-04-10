import pandas as pd
import sqlite3
import os

# Path to your Excel file
file_path = r"C:\Users\suraj\Desktop\Assignment 1 - Sample Data.xlsx"

# Function to read Excel file and return data
def read_excel_file(file_path):
    try:
        # Load Excel data using pandas
        data = pd.read_excel(file_path)
        print("Data loaded successfully!")
        print(data.head())  # Print the first few rows to verify
        return data
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return None

# Function to check if the data already exists in the database
def data_exists_in_db(db_path, sale_date, product_name, city):
    try:
        # Connect to SQLite database
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Query to check if the record already exists
        query = """
        SELECT COUNT(*) FROM sales_data WHERE sale_date = ? AND product_name = ? AND city = ?
        """
        cursor.execute(query, (sale_date, product_name, city))
        result = cursor.fetchone()

        conn.close()
        return result[0] > 0  # Returns True if record exists, otherwise False
    except sqlite3.Error as e:
        print(f"SQLite error: {e}")
        return False

# Function to insert data into SQLite database if not already present
def insert_data_if_not_exists(db_path, data):
    try:
        # Connect to SQLite database
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # SQL Insert query
        insert_query = """
        INSERT INTO sales_data (sale_date, channel, product_name, city, quantity, sales)
        VALUES (?, ?, ?, ?, ?, ?)
        """

        for index, row in data.iterrows():
            sale_date = row['Date']
            product_name = row['Product Name']
            city = row['City']

            # Check if the data already exists before inserting
            if not data_exists_in_db(db_path, sale_date, product_name, city):
                # Insert the new record if it doesn't exist
                cursor.execute(insert_query, (
                    sale_date, row['Channel'], product_name,
                    city, row['Quantity'], row['Sales']
                ))
                print(f"Inserting row: {row}")  # Debugging line
            else:
                print(f"Data for {sale_date} - {product_name} in {city} already exists.")

        # Commit changes and close connection
        conn.commit()
        conn.close()
        print(f"Data insertion completed.")
    except sqlite3.Error as e:
        print(f"SQLite error: {e}")
    except Exception as e:
        print(f"General error: {e}")

# Function to execute SQL query and fetch results
def execute_sql_query(db_path, sql_query):
    try:
        # Connect to SQLite database
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Execute SQL query
        cursor.execute(sql_query)

        # Fetch all rows from the result
        rows = cursor.fetchall()
        
        # Fetch column names
        col_names = [description[0] for description in cursor.description]
        
        conn.close()
        return rows, col_names
    except sqlite3.Error as e:
        print(f"SQLite error: {e}")
        return [], []
    except Exception as e:
        print(f"General error: {e}")
        return [], []

# Function to fetch data from the sales_data table using a SQL query
def fetch_data_from_sales_data(db_path, query):
    rows, columns = execute_sql_query(db_path, query)
    
    if rows:
        df = pd.DataFrame(rows, columns=columns)
        return df
    else:
        return None

# Main execution
db_path = 'sales_data.db'  # Your SQLite database path
data = read_excel_file(file_path)  # Load data from Excel file

if data is not None:
    # Insert data into SQLite DB only if not already present
    insert_data_if_not_exists(db_path, data)

# Example SQL Queries to fetch data (13 queries)
queries = {
    "Show total sales and quantity per city": """
        SELECT City, SUM(Sales) AS Total_Sales, SUM(Quantity) AS Total_Quantity
        FROM sales_data
        GROUP BY City
    """,

    "Which city had the highest sales in 2024": """
        SELECT City, SUM(Sales) AS Total_Sales
        FROM sales_data
        WHERE sale_date BETWEEN '2024-01-01' AND '2024-12-31'
        GROUP BY City
        ORDER BY Total_Sales DESC
        LIMIT 1
    """,

    "Get monthly sales for Product 2 in 2025": """
        SELECT strftime('%Y-%m', sale_date) AS Month, SUM(Sales) AS Total_Sales
        FROM sales_data
        WHERE Product_Name = 'Product 2' AND sale_date BETWEEN '2025-01-01' AND '2025-12-31'
        GROUP BY Month
        ORDER BY Month
    """,

    "Show top 3 cities by total quantity sold": """
        SELECT City, SUM(Quantity) AS Total_Quantity
        FROM sales_data
        GROUP BY City
        ORDER BY Total_Quantity DESC
        LIMIT 3
    """,

    "List product names with their total sales": """
        SELECT Product_Name, SUM(Sales) AS Total_Sales
        FROM sales_data
        GROUP BY Product_Name
    """,

    "Find total quantity sold for each channel in the last 6 months": """
        SELECT Channel, SUM(Quantity)
        FROM sales_data
        WHERE sale_date >= date('now', '-6 months')
        GROUP BY Channel
    """,

    "What is the average sales per transaction for Product 2": """
        SELECT AVG(Sales)
        FROM sales_data
        WHERE Product_Name = 'Product 2'
    """,

    "Rank cities based on total sales": """
        SELECT City, SUM(Sales) AS Total_Sales,
               RANK() OVER (ORDER BY SUM(Sales) DESC) AS Rank
        FROM sales_data
        GROUP BY City
    """,

    "Get sales in City1 for Channel 1 in October 2024": """
        SELECT *
        FROM sales_data
        WHERE City = 'City1' AND Channel = 'Channel 1'
        AND sale_date BETWEEN '2024-10-01' AND '2024-10-31'
    """,

    "Compare sales in January and February 2025": """
        SELECT strftime('%Y-%m', sale_date) AS Month, SUM(Sales)
        FROM sales_data
        WHERE sale_date BETWEEN '2025-01-01' AND '2025-02-28'
        GROUP BY Month
    """,

    "What are the monthly sales across platform1 since Jan 2025?": """
        SELECT strftime('%Y-%m', sale_date) AS Month, SUM(Sales) AS Total_Sales
        FROM sales_data
        WHERE Channel = 'Platform1' AND sale_date >= '2025-01-01'
        GROUP BY Month
        ORDER BY Month
    """,

    "What is the share of units sold across various platforms since Jan 2025?": """
        SELECT Channel, SUM(Quantity) AS Total_Quantity,
        (SUM(Quantity) / (SELECT SUM(Quantity) FROM sales_data WHERE sale_date >= '2025-01-01')) * 100 AS Share_Percent
        FROM sales_data
        WHERE sale_date >= '2025-01-01'
        GROUP BY Channel
    """,

    "Can you tell me the top 5 days with the highest daily units sold?": """
        SELECT sale_date, SUM(Quantity) AS Total_Quantity
        FROM sales_data
        GROUP BY sale_date
        ORDER BY Total_Quantity DESC
        LIMIT 5
    """
}

# Fetching data using the queries defined above
for query_name, query in queries.items():
    print(f"\nQuery: {query_name}")
    fetched_data = fetch_data_from_sales_data(db_path, query)

    # Displaying results
    if fetched_data is not None:
        print(fetched_data)
    else:
        print("No data found for the query.")
