import streamlit as st
import pandas as pd
from datetime import datetime

def load_excel(file):
    # Load the Excel file and return a dictionary of sheet names and their data
    excel_data = pd.ExcelFile(file)
    sheet_names = excel_data.sheet_names
    data = {sheet: excel_data.parse(sheet) for sheet in sheet_names}
    return data, sheet_names

def compare_projects(projects_new_data, ucf_awards_data):
    # Extract the relevant dataframes (assuming 'CUI Active(CMMC L2)' and 'Sheet1' contain the needed data)
    projects_new_df = projects_new_data.get("CUI Active(CMMC L2)", pd.DataFrame())
    ucf_awards_df = ucf_awards_data.get("Sheet1", pd.DataFrame())

    # Strip any leading/trailing whitespace from column names
    projects_new_df.columns = projects_new_df.columns.str.strip()
    ucf_awards_df.columns = ucf_awards_df.columns.str.strip()

    # Find the column that contains "HURON" in its name
    huron_column = [col for col in projects_new_df.columns if 'HURON' in col.upper()]
    
    if not huron_column:
        raise KeyError("No column containing 'HURON' found in Projects NEW.")
    
    huron_column_name = huron_column[0]  # Take the first match

    # Extract the first 11 characters from the "Award" column in UCF Awards to match the "HURON Award ID"
    ucf_awards_df["Award ID"] = ucf_awards_df["Award"].str[:11]

    # Compare Projects NEW vs UCF Awards by HURON Award ID
    projects_new_not_in_ucf = projects_new_df[~projects_new_df[huron_column_name].isin(ucf_awards_df["Award ID"])]
    ucf_not_in_projects_new = ucf_awards_df[~ucf_awards_df["Award ID"].isin(projects_new_df[huron_column_name])]

    # Convert "End Date" and "Official End Date" columns to datetime format, force errors to NaT (Not a Time)
    projects_new_df["End Date"] = pd.to_datetime(projects_new_df["End Date"], errors='coerce', format='%m/%d/%Y')
    ucf_awards_df["Official End Date"] = pd.to_datetime(ucf_awards_df["Official End Date"], errors='coerce', format='%m/%d/%Y')

    # Get today's date
    today = datetime.today()

    # Identify expired projects based on "End Date" for Projects NEW (only if date is valid)
    expired_projects_df = projects_new_data.get("EXPIRED", pd.DataFrame())
    
    # Handle possible invalid or missing "End Date" by ensuring the column is converted to datetime properly
    expired_projects_df["End Date"] = pd.to_datetime(expired_projects_df["End Date"], errors='coerce', format='%m/%d/%Y')
    
    # Make sure "End Date" is not NaT before comparison
    expired_projects = expired_projects_df[expired_projects_df["End Date"].notna() & (expired_projects_df["End Date"] < today)]

    # Prepare the summary report
    report = {
        "Projects in Projects NEW not in UCF Awards": projects_new_not_in_ucf,
        "Projects in UCF Awards not in Projects NEW": ucf_not_in_projects_new,
        "Expired Projects in Projects NEW": expired_projects
    }
    
    return report

def display_summary_report(report):
    st.write("### Summary Report")
    for title, df in report.items():
        st.write(f"#### {title}")
        if df.empty:
            st.write("No projects found.")
        else:
            st.dataframe(df)

def main():
    st.title("Project Comparison Tool")
    
    st.sidebar.header("Upload Files")
    projects_new_file = st.sidebar.file_uploader("Upload 'Projects NEW' Excel file", type=["xlsx"])
    ucf_awards_file = st.sidebar.file_uploader("Upload 'UCF Awards' Excel file", type=["xlsx"])
    
    if projects_new_file and ucf_awards_file:
        st.write("Files uploaded successfully!")
        
        # Load both files and get the sheet names
        projects_new_data, projects_new_sheets = load_excel(projects_new_file)
        ucf_awards_data, ucf_awards_sheets = load_excel(ucf_awards_file)
        
        # Show available sheet names for the user to choose from
        st.sidebar.write(f"Sheets in 'Projects NEW': {', '.join(projects_new_sheets)}")
        st.sidebar.write(f"Sheets in 'UCF Awards': {', '.join(ucf_awards_sheets)}")

        # Compare and generate the report
        report = compare_projects(projects_new_data, ucf_awards_data)
        
        if report:
            display_summary_report(report)
    else:
        st.write("Please upload both Excel files to start the comparison.")
    
if __name__ == "__main__":
    main()

