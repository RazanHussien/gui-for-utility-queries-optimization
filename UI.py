# Standard Library
import os
import sys
import tkinter as tk
from tkinter import ttk

# Third Party Libraries
import pandas as pd
from sqlalchemy import create_engine
from dotenv import load_dotenv
import cx_Oracle
import logging

# Function to get the correct path to the resource files
def get_resource_path(relative_path: str):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Configure logging
log_file = get_resource_path("UI.log")
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler(log_file),
                        logging.StreamHandler(sys.stdout)
                    ])

# Load environment variables from the .env file
dotenv_path = get_resource_path(".env")
load_dotenv(dotenv_path)

# Connection String
connection_string_c89qry = os.getenv("CIS")

if not connection_string_c89qry:
    logging.error("CIS connection string not found in environment variables.")
else:
    engine_c89qry = create_engine(f"oracle+cx_oracle://{connection_string_c89qry}")

# Read SQL Queries
def read_query(file_path: str) -> str:
    try:
        with open(file_path, 'r') as file:
            logging.info(f"Reading query file {file_path}")
            return file.read().strip()
    except Exception as e:
        logging.error(f"Error reading query file {file_path}: {e}")
        return ""

# Define the paths to the SQL query files
water_query_path = get_resource_path("WATER MEP-MRP.sql")
gas_query_path = get_resource_path("GAS MEP-MRP.sql")
electric_query_path = get_resource_path("ELECTRIC MEP-MRP.sql")

# Log the paths to ensure they are correct
logging.debug(f"Water Query Path: {water_query_path}")
logging.debug(f"Gas Query Path: {gas_query_path}")
logging.debug(f"Electric Query Path: {electric_query_path}")

# Read the SQL query files
water_query = read_query(water_query_path)
gas_query = read_query(gas_query_path)
electric_query = read_query(electric_query_path)

# Execute Query and Save to Excel with Progress Bar
def execute_query_to_excel(query: str, output_filename: str, sheet_name: str, progress_var: tk.IntVar, root: tk.Tk):
    logging.debug(f"Executing query and saving to {output_filename}")
    try:
        with cx_Oracle.connect(connection_string_c89qry) as conn:
            with conn.cursor() as cursor:
                cursor.execute(query)
                result = cursor.fetchall()
                columns = [col[0] for col in cursor.description]
                progress_var.set(50)
                root.update_idletasks()
                
        df = pd.DataFrame(result, columns=columns)
        
        for col in df.columns:
            if col in ['RECEIVE_DT', 'INSTALL_DT', 'START_DT', 'END_DT']:
                df[col] = pd.to_datetime(df[col])
                df[col] = df[col].dt.strftime('%m-%d-%Y')
        df.to_excel(output_filename, sheet_name=sheet_name, index=False)
        progress_var.set(75)
        root.update_idletasks()
    except cx_Oracle.DatabaseError as e:
        logging.error(f"Database error: {e}")
    except Exception as e:
        logging.error(f"Error executing query and saving to Excel: {e}")

# Button Command Functions
def execute_and_open_water_query():
    progress_var.set(0)
    root.update_idletasks()
    logging.debug("Executing water query")
    execute_query_to_excel(water_query, 'water_query_output.xlsx', 'Water Query', progress_var, root)
    open_excel('water_query_output.xlsx')
    progress_var.set(100)
    root.update_idletasks()

def execute_and_open_gas_query():
    progress_var.set(0)
    root.update_idletasks()
    logging.debug("Executing gas query")
    execute_query_to_excel(gas_query, 'gas_query_output.xlsx', 'Gas Query', progress_var, root)
    open_excel('gas_query_output.xlsx')
    progress_var.set(100)
    root.update_idletasks()

def execute_and_open_electric_query():
    progress_var.set(0)
    root.update_idletasks()
    logging.debug("Executing electric query")
    execute_query_to_excel(electric_query, 'electric_query_output.xlsx', 'Electric Query', progress_var, root)
    open_excel('electric_query_output.xlsx')
    progress_var.set(100)
    root.update_idletasks()

# Functions to Open the Corresponding Excel Files
def open_excel(file_name: str) -> None:
    try:
        os.system(f"start EXCEL.EXE {file_name}")
        logging.info(f"Opening {file_name} in Excel")
    except Exception as e:
        logging.error(f"Error opening Excel file {file_name}: {e}")

# GUI Setup
def create_button(root: tk.Tk, text: str, command: callable, bg_color: str) -> None:
    button_style = {'font': ("Helvetica", 12), 'width': 20, 'height': 2}
    button = tk.Button(root, text=text, command=command, bg=bg_color, fg='white', **button_style)
    button.pack(pady=10, anchor='center')

root = tk.Tk()
root.title('Click the Button to Execute the Query and View the Excel File')
root.geometry("600x400")
root.configure(bg='white')

# Progress Bar
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate", maximum=100, variable=progress_var)
progress_bar.pack(pady=20)

# Create Buttons
create_button(root, "Water MEP-MRP", execute_and_open_water_query, '#1A1A2E')
create_button(root, "Gas MEP-MRP", execute_and_open_gas_query, '#162447')
create_button(root, "Electric MEP-MRP", execute_and_open_electric_query, '#0F3460')

root.mainloop()