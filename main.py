import pandas as pd
import streamlit as st
import numpy as np
from io import BytesIO

st.set_page_config(layout="wide")

def load_excel(file, skiprows):
    file_extension = file.name.split(".")[-1].lower()
    engine = "xlrd" if file_extension == "xls" else "openpyxl"
    df = pd.read_excel(file, skiprows=skiprows, engine=engine)
    df.columns = df.columns.str.strip()
    return df

def display_filtered_attendance(df):
    st.sidebar.write("### Filter Options")
    student_id_filter = st.sidebar.text_input("Filter by Student ID")
    student_name_filter = st.sidebar.text_input("Filter by Student Name")
    faculty_filter = st.sidebar.text_input("Filter by Faculty")

    df_sorted = df.sort_values(by=["Student ID", "Student Name"], ascending=True)

    if student_id_filter:
        df_sorted = df_sorted[df_sorted["Student ID"].astype(str).str.contains(student_id_filter, case=False, na=False)]
    if student_name_filter:
        df_sorted = df_sorted[df_sorted["Student Name"].str.contains(student_name_filter, case=False, na=False)]
    if faculty_filter:
        df_sorted = df_sorted[df_sorted["Faculty"].str.contains(faculty_filter, case=False, na=False)]

    st.sidebar.write("### Select Columns to Display")
    selected_columns = st.sidebar.multiselect("Choose columns", df_sorted.columns.tolist(), default=df_sorted.columns.tolist())

    st.write("### Sorted, Filtered, and Selected Columns Attendance Record:")
    st.dataframe(df_sorted[selected_columns], width=1500)

    summary_df = df.drop_duplicates(subset=["Student ID", "Date"]).groupby(["Student ID", "Student Name"], as_index=False).agg(Classes_Taken=("Date", "count"))
    st.write("### Summary: Student Attendance Record")
    st.dataframe(summary_df, width=800)

def generate_batch_reports(attendance_df, extra_session_df):
    st.sidebar.write("### Batch & Student Filters")

    batches = sorted(attendance_df["Batch"].dropna().unique().tolist())
    selected_batch = st.sidebar.selectbox("Select Batch", options=["All"] + batches)

    student_ids = sorted(attendance_df["Student ID"].dropna().unique().tolist())
    selected_student_id = st.sidebar.selectbox("Select Student ID", options=["All"] + student_ids)

    # Filter by batch if not 'All'
    if selected_batch != "All":
        attendance_df = attendance_df[attendance_df["Batch"] == selected_batch]
        extra_session_df = extra_session_df[extra_session_df["Batch"] == selected_batch]

    if selected_student_id != "All":
        attendance_df = attendance_df[attendance_df["Student ID"] == selected_student_id]
        extra_session_df = extra_session_df[extra_session_df["Student ID"] == selected_student_id]

    if attendance_df.empty:
        st.warning("No attendance data found for selected filters.")
        return

    attendance_df_nodup = attendance_df.drop_duplicates(subset=["Student ID", "Date"])
    months = attendance_df_nodup["Date"].dt.to_period("M").dropna().unique()

    output = {}
    for month in months:
        month_str = month.strftime("%B-%Y")
        month_df = attendance_df_nodup[attendance_df_nodup["Date"].dt.to_period("M") == month]

        start_date, end_date = pd.Timestamp(month.start_time), pd.Timestamp(month.end_time)
        date_range = pd.date_range(start=start_date, end=end_date)

        students = month_df[["Student ID", "Student Name"]].drop_duplicates()
        pivot_df = students.copy()
        for date in date_range:
            pivot_df[date.strftime("%-d-%b")] = "A"

        for _, row in month_df.iterrows():
            date_col = row["Date"].strftime("%-d-%b")
            pivot_df.loc[pivot_df["Student ID"] == row["Student ID"], date_col] = "P"

        for _, row in extra_session_df.iterrows():
            if pd.isna(row["Extra Session Attendance Date"]): continue
            if row["Extra Session Attendance Date"].to_period("M") != month: continue
            date_col = row["Extra Session Attendance Date"].strftime("%-d-%b")
            if date_col in pivot_df.columns:
                pivot_df.loc[pivot_df["Student ID"] == row["Student ID"], date_col] = "E"

        output[month_str] = pivot_df

    # Export Excel
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        for sheet_name, sheet_df in output.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    excel_buffer.seek(0)

    st.success("Processed Successfully! Ready for download.")
    filename = f"{selected_batch if selected_batch != 'All' else 'AllBatches'}_Attendance.xlsx"
    st.download_button("Download Attendance Excel", data=excel_buffer, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Show Data on UI
    attendance_df["duplicate_attendance_count"] = attendance_df.groupby("Student ID")["Student ID"].transform("count")
    attendance_df["unique_attendance_count"] = attendance_df.drop_duplicates(subset=["Student ID", "Date"]).groupby("Student ID")["Date"].transform("count")
    st.write(f"### Attendance Data - Batch: {selected_batch} | Student ID: {selected_student_id}")
    st.dataframe(attendance_df)

def main():
    st.title("Student Attendance System")
    st.write("Upload the Main Attendance file and optionally an Extra Session Attendance file.")

    tab1, tab2 = st.tabs(["Batch-wise Export", "Real-time Filter View"])

    with tab1:
        attendance_file = st.file_uploader("Upload Main Attendance File", type=["xls", "xlsx"], key="main")
        extra_file = st.file_uploader("Upload Extra Session Attendance File (Optional)", type=["xls", "xlsx"], key="extra")

        if attendance_file:
            attendance_df = load_excel(attendance_file, skiprows=6)
            attendance_df = attendance_df.drop(columns=[col for col in ["Sr. No.", "Center", "Student Signature", "Remark"] if col in attendance_df.columns], errors='ignore')
            attendance_df = attendance_df.rename(columns=lambda x: x.strip())
            attendance_df["Date"] = pd.to_datetime(attendance_df["Date"], errors="coerce")

            extra_df = None
            if extra_file:
                extra_df = load_excel(extra_file, skiprows=4)
                extra_df["Extra Session Attendance Date"] = pd.to_datetime(extra_df["Extra Session Attendance Date"], errors="coerce")
            else:
                extra_df = pd.DataFrame(columns=["Student ID", "Extra Session Attendance Date", "Batch"])

            generate_batch_reports(attendance_df, extra_df)

    with tab2:
        file = st.file_uploader("Upload File to View and Filter", type=["xls", "xlsx"], key="sort")
        if file:
            df = load_excel(file, skiprows=6)
            df = df.drop(columns=[col for col in ["Sr. No.", "Center", "Student Signature", "Remark"] if col in df.columns], errors='ignore')
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            display_filtered_attendance(df)

    st.markdown("---")
    st.markdown("<p style='text-align: center;'>Developed with ❤️ by Syed Aun Muhammad</p>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
