import pandas as pd
import streamlit as st
import numpy as np
from io import BytesIO

# Set page configuration
st.set_page_config(layout="wide")

def process_attendance(attendance_file, extra_session_file):
    if attendance_file is not None and extra_session_file is not None:
        try:
            # Determine engine based on file extension
            def get_engine(file):
                return "xlrd" if file.name.split(".")[-1].lower() == "xls" else "openpyxl"
            
            # Load attendance data
            attendance_df = pd.read_excel(attendance_file, skiprows=6, engine=get_engine(attendance_file))
            attendance_df.columns = attendance_df.columns.str.strip()

            # Drop unwanted columns
            columns_to_exclude = ["Sr. No.", "Center", "Student Signature", "Remark"]
            attendance_df = attendance_df.drop(columns=[col for col in columns_to_exclude if col in attendance_df.columns], errors='ignore')

            # Ensure required columns exist
            required_columns = ["Student ID", "Student Name", "Date", "Batch", "Faculty"]
            for col in required_columns:
                if col not in attendance_df.columns:
                    st.error(f"Error: Required column '{col}' not found in Attendance File.")
                    return

            # Process attendance date
            attendance_df["Date"] = pd.to_datetime(attendance_df["Date"], errors='coerce')

            # Load extra session data
            extra_session_df = pd.read_excel(extra_session_file, skiprows=6, engine=get_engine(extra_session_file))
            extra_session_df.columns = extra_session_df.columns.str.strip()
            if "Extra Session Attendance Date" not in extra_session_df.columns or "Student ID" not in extra_session_df.columns:
                st.error("Error: Required columns not found in Extra Session File.")
                return
            
            # Process extra session date
            extra_session_df["Extra Session Attendance Date"] = pd.to_datetime(extra_session_df["Extra Session Attendance Date"], errors='coerce')

            # Sidebar - Batch Filter
            st.sidebar.write("### Filter Options")
            batches = attendance_df["Batch"].dropna().unique()
            selected_batch = st.sidebar.selectbox("Select Batch", options=sorted(batches))

            # Sidebar - Student ID Filter
            student_ids = attendance_df["Student ID"].dropna().unique()
            selected_student_id = st.sidebar.selectbox("Select Student ID", options=["All"] + list(sorted(student_ids)))

            # Filter attendance and extra sessions for selected batch
            batch_attendance = attendance_df[attendance_df["Batch"] == selected_batch]
            batch_extra_sessions = extra_session_df[extra_session_df["Batch"] == selected_batch]

            # Apply Student ID filter
            if selected_student_id != "All":
                batch_attendance = batch_attendance[batch_attendance["Student ID"] == selected_student_id]
                batch_extra_sessions = batch_extra_sessions[batch_extra_sessions["Student ID"] == selected_student_id]

            if batch_attendance.empty:
                st.warning("No attendance records found for the selected Batch or Student ID.")
                return

            # Prepare month-wise attendance
            output = {}
            batch_attendance_nodup = batch_attendance.drop_duplicates(subset=["Student ID", "Date"])

            months = batch_attendance_nodup["Date"].dt.to_period("M").dropna().unique()

            for month in months:
                month_str = month.strftime("%B-%Y")
                month_df = batch_attendance_nodup[batch_attendance_nodup["Date"].dt.to_period("M") == month]

                # Prepare list of dates for the month
                start_date = pd.Timestamp(month.start_time)
                end_date = pd.Timestamp(month.end_time)
                date_range = pd.date_range(start=start_date, end=end_date)

                # Prepare pivot table
                students = month_df[["Student ID", "Student Name"]].drop_duplicates()

                pivot_df = students.copy()
                for date in date_range:
                    pivot_df[date.strftime("%-d-%b")] = ""

                # Mark attendance
                for idx, row in month_df.iterrows():
                    student_id = row["Student ID"]
                    date_col = row["Date"].strftime("%-d-%b")
                    pivot_df.loc[pivot_df["Student ID"] == student_id, date_col] = "."

                # Mark extra session attendance
                for idx, row in batch_extra_sessions.iterrows():
                    student_id = row["Student ID"]
                    date = row["Extra Session Attendance Date"]
                    if pd.isna(date):
                        continue
                    if date.to_period("M") == month:
                        date_col = date.strftime("%-d-%b")
                        if date_col in pivot_df.columns:
                            pivot_df.loc[pivot_df["Student ID"] == student_id, date_col] = "E"

                output[month_str] = pivot_df

            # Create Excel file for download
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                for month, df_month in output.items():
                    df_month.to_excel(writer, sheet_name=month[:31], index=False)
            excel_buffer.seek(0)

            # Provide download button
            st.success("Processed Successfully! Ready for download.")
            st.download_button(
                label="Download Batch Attendance Excel",
                data=excel_buffer,
                file_name=f"{selected_batch}_Attendance.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Show the attendance data for UI with duplicates and extra sessions marked
            st.write("### Attendance and Extra Session Data (UI)")

            # Compute duplicate and unique attendance counts
            batch_attendance["duplicate_attendance_count"] = batch_attendance.groupby("Student ID")["Student ID"].transform("count")
            batch_attendance["unique_attendance_count"] = batch_attendance.drop_duplicates(subset=["Student ID", "Date"]).groupby("Student ID")["Date"].transform("count")

            # Show the filtered and processed attendance data in UI
            st.write(f"### Filtered Attendance Data for Batch: {selected_batch} and Student ID: {selected_student_id}")
            st.dataframe(batch_attendance)

        except Exception as e:
            st.error(f"An error occurred: {e}")

# Streamlit UI
st.title("Student Attendance Processor with Extra Sessions")
st.write("Upload the Main Attendance file and the Extra Sessions file to generate batch-wise month-wise attendance.")

attendance_file = st.file_uploader("Upload Main Attendance File", type=["xls", "xlsx"])
extra_session_file = st.file_uploader("Upload Extra Session Attendance File", type=["xls", "xlsx"])

if attendance_file and extra_session_file:
    process_attendance(attendance_file, extra_session_file)

# Footer
st.markdown("---")
st.markdown("<p style='text-align: center;'>Developed with ❤️ by Syed Aun Muhammad</p>", unsafe_allow_html=True)
