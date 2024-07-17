import streamlit as st
import pandas as pd
import re
from io import BytesIO
def load_excel_new_reports(file, header_row=19):
    """Loads an Excel file starting from the specified header row."""
    if file is not None:
        return pd.read_excel(file, header=header_row).fillna('2000-01-01')
def load_excel_old_reports(file, header_row=19):
    """Loads an Excel file starting from the specified header row."""
    if file is not None:
        return pd.read_excel(file).fillna('2000-01-01')
def load_csv(file):
    if file is not None:
        return pd.read_csv(file)
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer._save()
    processed_data = output.getvalue()
    return processed_data
def get_rows_not_in_first(df1, df2):
    """Returns rows from df2 not found in df1 based on all columns."""
    return pd.concat([df1, df2, df2]).drop_duplicates(keep=False)
def find_differences(old_df, new_entry_df):
    merged_df = pd.merge(old_df, new_entry_df, on='Employee ID', suffixes=('_df1', '_df2'))
    all_results=[]
    # Checking for differences
    for index, row in merged_df.iterrows():
        differences = []
        for col in old_df.columns[1:]:  # Skip 'Employee ID' since it's the key for merge
            if row[col + '_df1'] != row[col + '_df2']:
                differences.append(f"{col} differs: {row[col + '_df1']} vs {row[col + '_df2']}")
        if differences:
            text=f"Employee name {row['Worker_df1']} with ID {row['Employee ID']} has differences:  ", '; '.join(differences)
            all_results.append(str(text))
    return all_results
def find_new_entries(df_old,df_new):
    df_old['Pay Through Date']=df_old['Pay Through Date'].astype('datetime64[ns]')

    df_new['Pay Through Date']=df_new['Pay Through Date'].astype('datetime64[ns]')
    merged_df = pd.merge(df_old, df_new, how='outer', indicator=True)
    new_entries = merged_df[merged_df['_merge'] == 'right_only'].drop(columns=['_merge'])
    return new_entries
def extract_Employee_ID_from_Trello(trello_df):
    employee_id_list=[]
    for description in trello_df["Card Description"]:
            # Using regular expressions to find all 6-digit numbers
        if len(re.findall(r'\b\d{6}\b', str(description)))>0:
            numbers = re.findall(r'\b\d{6}\b', str(description))
            employee_id_list.extend(numbers)
        else:
             employee_id_list.extend('0')
    trello_df["Card Description"]=employee_id_list
    trello_df["Card Description"]=trello_df["Card Description"].astype(int)
    return trello_df
def find_new_entries_comparing_to_trello(trello_df,new_entries):
    trello_df.rename(columns = {'Card Description':'Employee ID'}, inplace = True)
    merged_df = pd.merge(trello_df, new_entries, on='Employee ID',how='outer', indicator=True)
    new_entries = merged_df[merged_df['_merge'] == 'right_only'].drop(columns=['Card ID','Card URL', 'Card Name', 'Labels',
       'Members', 'Due Date', 'Attachment Count', 'Attachment Links',
       'Checklist Item Total Count', 'Checklist Item Completed Count',
       'Vote Count', 'Comment Count', 'Last Activity Date', 'List ID',
       'List Name', 'Board ID', 'Board Name', 'Archived', 'Start Date',
       'Due Complete', 'To DO Date', 'Priority', 'Status','_merge']).fillna(0)
    return new_entries
# Streamlit app
st.title("Termination Excel Comparison Tool")
# File uploaders
file1 = st.file_uploader("Upload old Termination Excel report file", type=['xlsx'])
file2 = st.file_uploader("Upload new Termintaion Excel report file", type=['xlsx'])
trello_report = st.file_uploader("Upload Trello report file", type=['csv'])
if file1 and file2 and trello_report:
    # Load and display the data
    df1 = load_excel_new_reports(file1)
    df2 = load_excel_new_reports(file2)
    trello_df=load_csv(trello_report)
    if df1 is not None and df2 is not None and trello_df is not None:
        #show previews
        st.write("Preview of Old Report Table:")
        st.dataframe(df1)
        st.write("Preview of New Report Table:")
        st.dataframe(df2)
        st.write("Preview of Trello Report:")
        st.dataframe(trello_df)
        # Find new entries
        diff_df = find_new_entries(df1, df2)
        st.write("New Entries in New Report Table :")
        st.dataframe(diff_df)
        # Find differences of new entries and old report
        differences = find_differences(df1, diff_df)
        st.write("Differences of Old and New report:")
        s = ''
        for i in differences:
            s += "- " + i + "\n"
        st.markdown(s)
        #check for duplicates in trello and new entries
        trello_df=extract_Employee_ID_from_Trello(trello_df)
        #remove IDs found in trello
        diff_df=find_new_entries_comparing_to_trello(trello_df,diff_df)
        st.write("Preview of New entries with removed duplicated rows:")
        st.dataframe(diff_df)
        # Download link
        excel_file = to_excel(diff_df)
        st.download_button(
                label="Download Excel",
                data=excel_file,
                file_name='Upload List.xlsx',
            )
