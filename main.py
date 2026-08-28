import streamlit as st
import pandas as pd

# Page Configuration
st.set_page_config(page_title="Student Attendance Portal", layout="wide")
st.title("📊 Student Attendance Dashboard")

@st.cache_data
def load_data(file):
    df = pd.read_csv(file)
    # Parse DD/MM/YYYY dates for chronological ordering
    df['ParsedDate'] = pd.to_datetime(df['AttendanceDate'], format='%d/%m/%Y', errors='coerce')
    return df

# Upload or fallback to local file
uploaded_file = st.sidebar.file_uploader("Upload Student Attendance CSV", type=["csv"])
file_to_load = uploaded_file if uploaded_file is not None else "StudentAttendanceReport (4).csv"

try:
    df = load_data(file_to_load)
    
    # Sidebar quick filters
    st.sidebar.header("Global Filters")
    courses = ["All"] + sorted(df['CourseName'].dropna().unique().tolist())
    selected_course = st.sidebar.selectbox("Filter by Course", courses)
    
    if selected_course != "All":
        df = df[df['CourseName'] == selected_course]

    tab1, tab2 = st.tabs([
        "📅 Student-wise Attendance (Ordered by Date)", 
        "🔢 Unique Attendance Count per Student"
    ])

    # -------------------------------------------------------------
    # TAB 1: Student-wise Attendance Grouped & Ordered by Date
    # -------------------------------------------------------------
    with tab1:
        st.subheader("Student Attendance History")
        
        # Student selection dropdown for detailed drill-down
        student_list = df[['StudentID', 'StudentName']].drop_duplicates().sort_values('StudentID')
        student_options = ["All Students"] + [
            f"{row['StudentID']} - {row['StudentName']}" for _, row in student_list.iterrows()
        ]
        
        selected_student = st.selectbox("Select a Student to Filter (or view all)", student_options)

        # Sort strictly by StudentID and chronological Date
        sorted_df = df.sort_values(by=['StudentID', 'ParsedDate'], ascending=[True, True])
        
        display_columns = [
            'StudentID', 'StudentName', 'AttendanceDate', 'TimeSlotDesc', 
            'CourseName', 'ModuleName', 'SessionName', 'EmployeeName'
        ]

        if selected_student != "All Students":
            target_id = selected_student.split(" - ")[0]
            filtered_df = sorted_df[sorted_df['StudentID'] == target_id]
            st.write(f"Showing **{len(filtered_df)}** records for **{selected_student}**:")
            st.dataframe(filtered_df[display_columns], use_container_width=True, hide_index=True)
        else:
            st.write(f"Displaying **{len(sorted_df):,}** records grouped by StudentID and ordered by date:")
            st.dataframe(sorted_df[display_columns], use_container_width=True, hide_index=True)

    # -------------------------------------------------------------
    # TAB 2: Unique Day Attendance Count per Student
    # -------------------------------------------------------------
    with tab2:
        st.subheader("Attendance Aggregation (Deduplicated by Date)")
        st.caption("Counts unique calendar attendance dates as single presence (ignoring same-day multiple sessions).")

        summary_df = (
            df.groupby(['StudentID', 'StudentName'])
            .agg(
                Unique_Days_Attended=('AttendanceDate', 'nunique'),
                Total_Session_Records=('AttendanceDate', 'count')
            )
            .reset_index()
            .sort_values(by='Unique_Days_Attended', ascending=False)
        )

        # Quick metric indicators
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Unique Students", len(summary_df))
        col2.metric("Max Days Attended", summary_df['Unique_Days_Attended'].max())
        col3.metric("Avg Days Attended", f"{summary_df['Unique_Days_Attended'].mean():.1f}")

        # Search bar within summary table
        search_query = st.text_input("🔍 Search by Student ID or Name", "")
        if search_query:
            summary_df = summary_df[
                summary_df['StudentID'].str.contains(search_query, case=False, na=False) |
                summary_df['StudentName'].str.contains(search_query, case=False, na=False)
            ]

        st.dataframe(
            summary_df.rename(columns={
                'StudentID': 'Student ID',
                'StudentName': 'Student Name',
                'Unique_Days_Attended': 'Distinct Days Present',
                'Total_Session_Records': 'Total Recorded Sessions'
            }),
            use_container_width=True,
            hide_index=True
        )

        # Download button
        csv_download = summary_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Unique Attendance Summary (CSV)",
            data=csv_download,
            file_name="Student_Unique_Attendance_Summary.csv",
            mime="text/csv"
        )

except FileNotFoundError:
    st.error("Could not find the dataset. Please upload `StudentAttendanceReport (4).csv` using the sidebar uploader.")
