import pandas as pd
import streamlit as st

# Set page configuration for wide view
st.set_page_config(layout="wide")

def sort_attendance(uploaded_file):
    if uploaded_file is not None:
        # Read the Excel file, skipping the first six rows
        df = pd.read_excel(uploaded_file, skiprows=6)
        
        # Add filtering options
        st.sidebar.write("### Filter Options")
        student_id_filter = st.sidebar.text_input("Filter by Student ID")
        student_name_filter = st.sidebar.text_input("Filter by Student Name")
        
        # Sort data by Student ID and Student Name
        df_sorted = df.sort_values(by=["Student ID", "Student Name"], ascending=[True, True])
        
        # Apply filters
        if student_id_filter:
            df_sorted = df_sorted[df_sorted["Student ID"].astype(str).str.contains(student_id_filter, case=False, na=False)]
        if student_name_filter:
            df_sorted = df_sorted[df_sorted["Student Name"].str.contains(student_name_filter, case=False, na=False)]
        
        # Column selection
        st.sidebar.write("### Select Columns to Display")
        selected_columns = st.sidebar.multiselect("Choose columns", df_sorted.columns.tolist(), default=df_sorted.columns.tolist())
        
        # Display the sorted, filtered, and selected columns result in wide format
        st.write("### Sorted, Filtered, and Selected Columns Attendance Record:")
        st.dataframe(df_sorted[selected_columns], width=1500)

# Streamlit UI
st.title("Student Attendance Sorter")
st.write("Upload an Excel file to sort and filter attendance records in real-time.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xls"])

if uploaded_file is not None:
    sort_attendance(uploaded_file)

# Footer
st.markdown("---")
st.markdown("<p style='text-align: center;'>Developed with ❤️ by Syed Aun Muhammad</p>", unsafe_allow_html=True)
