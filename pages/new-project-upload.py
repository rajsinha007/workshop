import streamlit as st
import pandas as pd
import pyodbc

# Database details
server = 'sql-prod-ap-exp-cfit-b2dprdsvr-pldnsazr.database.windows.net'
database = 'b2d'
username = 'csvadm'
password = '4g#w5G#d6g2%q'
table_name = 'project_r2'

# Establish database connection
conn_string = f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database};UID={username};PWD={password}'

# Create Streamlit web app
st.title("Upload Excel File and Insert into Database")

# File upload section
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    # Read Excel file into a DataFrame
    df = pd.read_excel(uploaded_file)

    # Convert start_date and end_date columns to datetime
    df['start_date'] = pd.to_datetime(df['start_date'])
    df['end_date'] = pd.to_datetime(df['end_date'])

    # Calculate total days
    df['total_days'] = (df['end_date'] - df['start_date']).dt.days

    # Display DataFrame
    st.write("Uploaded Data:")
    st.dataframe(df)

    # Button to insert data into database
    if st.button("Insert into Database"):
        # Create a table if it does not exist
        create_table_query = f"""
        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='{table_name}' AND xtype='U')
        CREATE TABLE {table_name} (
            team VARCHAR(50),
            product_name VARCHAR(50),
            order_type VARCHAR(50),
            order_no VARCHAR(50),
            order_name VARCHAR(50),
            order_material VARCHAR(50),
            project_manager VARCHAR(50),
            order_value FLOAT,
            start_date DATE,
            end_date DATE,
            total_days INT
        );
        """

        # Insert data into the database table
        def insert_into_database(data):
            with pyodbc.connect(conn_string) as conn:
                cursor = conn.cursor()
                for _, row in data.iterrows():
                    values = (
                        row['team'], row['product'], row['order_type'], row['order_no'], row['order_name'],
                        row['order_material'], row['project_manager'], row['order_value'], row['start_date'].date(),
                        row['end_date'].date(), row['total_days']
                    )
                    cursor.execute(f"INSERT INTO {table_name} VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", values)
                conn.commit()

        # Execute the create table query
        with pyodbc.connect(conn_string) as conn:
            cursor = conn.cursor()
            cursor.execute(create_table_query)

        # Insert data into the database table
        insert_into_database(df)

        # Display success message
        st.success("Data inserted into the database table: " + table_name)
