import streamlit as st
import pandas as pd
from io import BytesIO
def load_excel(file, header_row=19):
    """Loads an Excel file starting from the specified header row."""
    if file is not None:
        return pd.read_excel(file, header=header_row).fillna('0001-01-01')

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
    merged_df = pd.merge(df_old, df_new, how='outer', indicator=True)
    new_entries = merged_df[merged_df['_merge'] == 'right_only'].drop(columns=['_merge'])
    return new_entries

# Streamlit app
st.title("Termination Excel Comparison Tool")

# File uploaders
file1 = st.file_uploader("Upload old Termination Excel report file", type=['xlsx'])
file2 = st.file_uploader("Upload new Termintaion Excel report file", type=['xlsx'])

if file1 and file2:
    # Load and display the data
    df1 = load_excel(file1)
    df2 = load_excel(file2)

    if df1 is not None and df2 is not None:
        st.write("Preview of First Table:")
        st.dataframe(df1)
        st.write("Preview of Second Table:")
        st.dataframe(df2)

        # Compute differences
        diff_df = find_new_entries(df1, df2)
        st.write("New Entries in New Report Table :")
        st.dataframe(diff_df)


        # Find differences
        differences = find_differences(df1, diff_df)
        st.write("Differences :")
        s = ''
        for i in differences:
            s += "- " + i + "\n"
        st.markdown(s)
      

        # Download link
        
        excel_file = to_excel(diff_df)
        st.download_button(
                label="Download Excel",
                data=excel_file,
                file_name='sample_data.xlsx',
            )
        
