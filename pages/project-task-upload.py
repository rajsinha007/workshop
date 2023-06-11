import streamlit as st
import pandas as pd
import pyodbc
from datetime import datetime

# Database connection details
server = 'sql-prod-ap-exp-cfit-b2dprdsvr-pldnsazr.database.windows.net'
database = 'b2d'
username = 'csvadm'
password = '4g#w5G#d6g2%q'
table_name = 'project_tasks_r2'

def calculate_est_effort(row):
    duration = row['totaldays']
    est_effort = round(duration * row['percentage'] / 100, 2)
    return est_effort

def calculate_est_charge(row):
    return row['est_effort'] * 345

def process_excel(file):
    df = pd.read_excel(file)
    df['totaldays'] = (df['end_date'] - df['start_date']).dt.days
    df['percentage'] = df['percentage'].astype(int)  # Convert percentage to integer
    df['est_effort'] = df.apply(calculate_est_effort, axis=1)
    df['est_charge'] = df.apply(calculate_est_charge, axis=1)
    # Reorder columns
    df = df[['team', 'product', 'order_no', 'order_material', 'task', 'assignee',
             'start_date', 'end_date', 'totaldays', 'percentage', 'est_effort', 'est_charge']]
    return df

def create_table_if_not_exists(connection):
    cursor = connection.cursor()
    cursor.execute(
        f"IF OBJECT_ID('{table_name}', 'U') IS NULL "
        f"CREATE TABLE {table_name} "
        "(team VARCHAR(50), product VARCHAR(50), order_no VARCHAR(50), order_material VARCHAR(50), "
        "task VARCHAR(50), assignee VARCHAR(50), start_date DATE, end_date DATE, totaldays INT, "
        "percentage INT, est_effort DECIMAL(10,2), est_charge DECIMAL(10,2))"
    )
    connection.commit()

def insert_data_to_table(connection, df):
    cursor = connection.cursor()
    for _, row in df.iterrows():
        cursor.execute(
            f"INSERT INTO {table_name} "
            "(team, product, order_no, order_material, task, assignee, start_date, end_date, "
            "totaldays, percentage, est_effort, est_charge) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            row.tolist()
        )
    connection.commit()

def main():
    st.title("Excel File Processor")

    # File upload section
    st.header("Upload Excel File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Check file extension
        if uploaded_file.name != 'project-task-setup.xlsx':
            st.error("Invalid file. Please upload 'project-task-setup.xlsx' only.")
        else:
            df = process_excel(uploaded_file)
            st.header("Processed Data")
            st.dataframe(df)

            # Database connection and insertion
            if st.button("Insert into Database"):
                try:
                    connection = pyodbc.connect(
                        f"DRIVER={{SQL Server}};"
                        f"SERVER={server};DATABASE={database};UID={username};PWD={password}"
                    )

                    create_table_if_not_exists(connection)
                    insert_data_to_table(connection, df)

                    st.success("Data inserted into the database successfully!")
                except pyodbc.Error as e:
                    st.error(f"An error occurred while connecting to the database: {e}")

# Run the app
if __name__ == "__main__":
    main()
