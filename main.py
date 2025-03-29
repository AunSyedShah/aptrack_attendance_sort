import pandas as pd
import streamlit as st
import numpy as np

# Set page configuration for wide view
st.set_page_config(layout="wide")

def sort_attendance(uploaded_file):
    if uploaded_file is not None:
        try:
            # Determine the file extension and select the appropriate engine
            file_extension = uploaded_file.name.split(".")[-1].lower()
            engine = "xlrd" if file_extension == "xls" else "openpyxl"
            
            # Read the Excel file, skipping the first six rows
            df = pd.read_excel(uploaded_file, skiprows=6, engine=engine)
            
            # Strip any leading or trailing spaces from column names
            df.columns = df.columns.str.strip()
            
            # Drop unwanted columns
            columns_to_exclude = ["Sr. No.", "Center", "Student Signature", "Remark"]
            df = df.drop(columns=[col for col in columns_to_exclude if col in df.columns], errors='ignore')
            
            # Check if required columns are present
            required_columns = ["Student ID", "Student Name", "Date", "Faculty"]
            for col in required_columns:
                if col not in df.columns:
                    st.error(f"Error: Required column '{col}' not found in the uploaded file.")
                    return
            
            # Convert Date column to datetime for accurate grouping
            df["Date"] = pd.to_datetime(df["Date"], errors='coerce')
            
            # Compute attendance counts
            df["duplicate_attendance_count"] = df.groupby("Student ID")["Student ID"].transform("count")
            df["unique_attendance_count"] = df.drop_duplicates(subset=["Student ID", "Date"]).groupby("Student ID")["Date"].transform("count")
            
            # Add filtering options
            st.sidebar.write("### Filter Options")
            student_id_filter = st.sidebar.text_input("Filter by Student ID")
            student_name_filter = st.sidebar.text_input("Filter by Student Name")
            faculty_filter = st.sidebar.text_input("Filter by Faculty")
            
            # Sort data by Student ID and Student Name
            df_sorted = df.sort_values(by=["Student ID", "Student Name"], ascending=[True, True])
            
            # Apply filters
            if student_id_filter:
                df_sorted = df_sorted[df_sorted["Student ID"].astype(str).str.contains(student_id_filter, case=False, na=False)]
            if student_name_filter:
                df_sorted = df_sorted[df_sorted["Student Name"].str.contains(student_name_filter, case=False, na=False)]
            if faculty_filter:
                df_sorted = df_sorted[df_sorted["Faculty"].str.contains(faculty_filter, case=False, na=False)]
            
            # Column selection
            st.sidebar.write("### Select Columns to Display")
            selected_columns = st.sidebar.multiselect("Choose columns", df_sorted.columns.tolist(), default=df_sorted.columns.tolist())
            
            # Display the sorted, filtered, and selected columns result in wide format
            st.write("### Sorted, Filtered, and Selected Columns Attendance Record:")
            st.dataframe(df_sorted[selected_columns], width=1500)
            
            # Create another DataFrame for summary view
            summary_df = df.drop_duplicates(subset=["Student ID", "Date"]).groupby(["Student ID", "Student Name"], as_index=False).agg(Classes_Taken=("Date", "count"))
            
            # Display the summary dataframe
            st.write("### Summary: Student Attendance Record")
            st.dataframe(summary_df, width=800)
        
        except Exception as e:
            st.error(f"An error occurred: {e}")

# Streamlit UI
st.title("Student Attendance Sorter")
st.write("Upload an Excel file to sort and filter attendance records in real-time.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xls", "xlsx"])

if uploaded_file is not None:
    sort_attendance(uploaded_file)

# Footer
st.markdown("---")
st.markdown("<p style='text-align: center;'>Developed with ❤️ by Syed Aun Muhammad</p>", unsafe_allow_html=True)