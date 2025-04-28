import pandas as pd
import streamlit as st
import io

# Set page configuration for wide view
st.set_page_config(layout="wide")

def sort_attendance(uploaded_file):
    if uploaded_file is not None:
        try:
            # Determine file extension and select appropriate engine
            file_extension = uploaded_file.name.split(".")[-1].lower()
            engine = "xlrd" if file_extension == "xls" else "openpyxl"

            # Read the Excel file, skipping the first six rows
            df = pd.read_excel(uploaded_file, skiprows=6, engine=engine)

            # Strip spaces from column names
            df.columns = df.columns.str.strip()

            # Drop unwanted columns
            columns_to_exclude = ["Sr. No.", "Center", "Student Signature", "Remark"]
            df = df.drop(columns=[col for col in columns_to_exclude if col in df.columns], errors='ignore')

            # Required columns check
            required_columns = ["Student ID", "Student Name", "Date", "Faculty", "Batch"]
            for col in required_columns:
                if col not in df.columns:
                    st.error(f"Error: Required column '{col}' not found in the uploaded file.")
                    return

            # Convert Date column to datetime
            df["Date"] = pd.to_datetime(df["Date"], errors='coerce')

            # Compute duplicate and unique attendance counts
            df["duplicate_attendance_count"] = df.groupby("Student ID")["Student ID"].transform("count")
            df["unique_attendance_count"] = df.drop_duplicates(subset=["Student ID", "Date"]).groupby("Student ID")["Date"].transform("count")

            # Sort data
            df_sorted = df.sort_values(by=["Student ID", "Student Name"], ascending=[True, True])

            # Sidebar Filters
            st.sidebar.write("### Filter Options")
            student_id_filter = st.sidebar.text_input("Filter by Student ID")
            student_name_filter = st.sidebar.text_input("Filter by Student Name")
            faculty_filter = st.sidebar.text_input("Filter by Faculty")

            # New Batch filter
            batch_list = df_sorted["Batch"].dropna().unique().tolist()
            batch_selected = st.sidebar.selectbox("Select Batch", options=[""] + sorted(batch_list))

            # Apply filters
            if student_id_filter:
                df_sorted = df_sorted[df_sorted["Student ID"].astype(str).str.contains(student_id_filter, case=False, na=False)]
            if student_name_filter:
                df_sorted = df_sorted[df_sorted["Student Name"].str.contains(student_name_filter, case=False, na=False)]
            if faculty_filter:
                df_sorted = df_sorted[df_sorted["Faculty"].str.contains(faculty_filter, case=False, na=False)]
            if batch_selected:
                df_sorted = df_sorted[df_sorted["Batch"] == batch_selected]

            # If no batch selected, stop here
            if not batch_selected:
                st.warning("Please select a Batch to generate Month-wise Attendance export.")
                return

            # Generate Month-wise Attendance Sheets
            df_filtered = df_sorted.drop_duplicates(subset=["Student ID", "Date"])
            df_filtered["Month_Year"] = df_filtered["Date"].dt.to_period("M")

            output = io.BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for month_year, month_df in df_filtered.groupby("Month_Year"):
                    month_str = month_year.strftime("%B %Y")

                    # Pivot table
                    pivot = month_df.pivot_table(
                        index=["Student ID", "Student Name"],
                        columns=month_df["Date"].dt.day,
                        values="Date",
                        aggfunc="count",
                        fill_value=""
                    )

                    # Replace non-empty cells with "."
                    pivot = pivot.applymap(lambda x: "." if x != "" else "")

                    # Sort columns (days of month)
                    pivot = pivot.reindex(sorted(pivot.columns), axis=1)

                    # Write each month's data in separate sheets
                    pivot.to_excel(writer, sheet_name=month_str[:31])  # Sheet names max 31 characters

            st.success("Excel file ready for download!")

            # Provide download button
            st.download_button(
                label="📥 Download Month-wise Attendance Excel",
                data=output.getvalue(),
                file_name=f"{batch_selected}_Attendance_Monthwise.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")

# Streamlit UI
st.title("Student Attendance Sorter & Exporter")
st.write("Upload an Excel file to sort, filter, and export month-wise attendance per batch.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xls", "xlsx"])

if uploaded_file is not None:
    sort_attendance(uploaded_file)

# Footer
st.markdown("---")
st.markdown("<p style='text-align: center;'>Developed with ❤️ by Syed Aun Muhammad</p>", unsafe_allow_html=True)
