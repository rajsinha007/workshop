import streamlit as st
import pandas as pd
import numpy as np
import pyodbc

# Database connection string
server = 'sql-prod-ap-exp-cfit-b2dprdsvr-pldnsazr.database.windows.net'
database = 'b2d'
username = 'csvadm'
password = '4g#w5G#d6g2%q'

connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"

@st.cache(allow_output_mutation=True)
def load_data(file):
    if file.name.endswith('.xlsx'):
        if file.name == 'project-task-setup.xlsx':
            df = pd.read_excel(file, engine='openpyxl')
            df = process_task_setup(df)
        elif file.name == 'project-order-setup.xlsx':
            df = pd.read_excel(file, engine='openpyxl')
            df = process_order_setup(df)
        else:
            st.warning("Invalid file uploaded. Please upload 'project-task-setup.xlsx' or 'project-order-setup.xlsx'.")
            df = None
    else:
        st.warning("Invalid file format. Please upload an Excel file.")
        df = None
    return df

def process_task_setup(df):
    expected_columns = [
        'team', 'product', 'order_no', 'order_material', 'task', 'assignee',
        'start_date', 'end_date', 'percentage'
    ]
    df = df[expected_columns].copy()
    df['start_date'] = pd.to_datetime(df['start_date'], errors='coerce')
    df['end_date'] = pd.to_datetime(df['end_date'], errors='coerce')
    df['est_effort'] = (df['end_date'] - df['start_date']).dt.days
    df['est_effort'] = df['est_effort'] * df['percentage'] 
    df['est_charge'] = df['est_effort'] * 345
    df['percentage'] = df['percentage'].astype(str) + '%'
    df['est_effort'] = df['est_effort'].round(0)
    df = df.dropna(how='all', axis=1)
    return df

def process_order_setup(df):
    expected_columns = [
        'order_type', 'order_no', 'order_name', 'order_material', 'project_manager',
        'order_value', 'start_date', 'end_date'
    ]
    df = df[expected_columns].copy()
    df['start_date'] = pd.to_datetime(df['start_date'], errors='coerce')
    df['end_date'] = pd.to_datetime(df['end_date'], errors='coerce')
    
    if 'total_days' not in df.columns:
        df['total_days'] = (df['end_date'] - df['start_date']).dt.days
        df['total_days'] = df['total_days'].fillna(0).astype(int)
    
    df = df.dropna(how='all', axis=1)
    return df


def create_table():
    try:
        with pyodbc.connect(connection_string) as conn:
            with conn.cursor() as cursor:
                cursor.execute(
                    """
                    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='project_task_r2')
                    CREATE TABLE project_task_r2 (
                        team VARCHAR(255),
                        product VARCHAR(255),
                        order_no VARCHAR(255),
                        order_material VARCHAR(255),
                        task VARCHAR(255),
                        assignee VARCHAR(255),
                        start_date DATE,
                        end_date DATE,
                        percentage FLOAT,
                        est_effort INT,
                        est_charge FLOAT,
                        total_days INT
                    )
                    """
                )
            conn.commit()
        st.success("Table created successfully.")
    except pyodbc.Error as e:
        st.error(f"Error creating table: {e}")


def insert_records(df):
    try:
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        for _, row in df.iterrows():
            cursor.execute("""
                INSERT INTO project_task_r2 (
                    team, product, order_no, order_material, task, assignee, start_date, 
                    end_date, percentage, est_effort, est_charge, total_days
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                row.get('team', None), row.get('product', None), row.get('order_no', None),
                row.get('order_material', None), row.get('task', None), row.get('assignee', None),
                row.get('start_date', None), row.get('end_date', None), row.get('percentage', None),
                row.get('est_effort', None), row.get('est_charge', None), row.get('total_days', None)
            ))
        conn.commit()
        st.success("Records inserted successfully.")
    except pyodbc.Error as e:
        st.error(f"Error inserting records: {e}")

# Main program
st.title("DHL Project Forecast App")
st.markdown("**DHL Project Forecast App**")
st.markdown("---")

# Upload file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    df = load_data(uploaded_file)
    if df is not None:
        st.subheader("Data Frame")
        st.dataframe(df)

        if uploaded_file.name == 'project-task-setup.xlsx':
            create_table()
            if st.button("Insert Records"):
                insert_records(df)
        elif uploaded_file.name == 'project-order-setup.xlsx':
            if 'total_days' in df.columns:
                create_table()
                if st.button("Insert Records"):
                    insert_records(df)
            else:
                st.warning("Missing 'total_days' column in the file.")
    else:
        st.error("Error processing file.")

# 