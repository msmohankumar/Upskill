import streamlit as st
import openpyxl
import pandas as pd

# Path to the Excel file
file_path = r"U:\2025\Data science app\Upskilling app\Characteristics Competency Map_Design_Services_M S Mohan Kumar V2.xlsx"

# Function to load the Excel file
def load_excel(file_path):
    # Open the existing Excel file
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    # Convert the Excel data to a pandas DataFrame for easy manipulation
    data = pd.read_excel(file_path, sheet_name=sheet.title)
    return wb, sheet, data

# Function to update a rating
def update_rating(row, col, value):
    wb, sheet, data = load_excel(file_path)
    sheet.cell(row=row, column=col, value=value)  # Update the specified cell
    wb.save(file_path)  # Save the file
    return data

# Streamlit UI elements
st.title("Upskilling Monitoring App")
st.sidebar.header("Controls")

# Display the data in a table
wb, sheet, data = load_excel(file_path)
st.dataframe(data)

# Year selection filter
year = st.selectbox("Select Year", ["2022", "2023", "2024", "2025"])

# Show ratings and allow users to update them
row_to_edit = st.number_input("Enter the row number you want to update", min_value=2, max_value=len(data))
column_to_edit = st.number_input("Enter the column number (Rating column)", min_value=5, max_value=10)
new_value = st.number_input("Enter new rating", min_value=0, max_value=3, step=0.5)

if st.button("Update Rating"):
    updated_data = update_rating(row_to_edit, column_to_edit, new_value)
    st.success(f"Rating for row {row_to_edit}, column {column_to_edit} updated to {new_value}!")
    st.dataframe(updated_data)
