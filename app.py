import pandas as pd
import streamlit as st
import io
from thefuzz import fuzz

def is_similar(text, keywords, threshold=85):
    """Check string similarity with fuzzy matching"""
    text = str(text).lower()
    
    # Direct match first
    if any(keyword.lower() in text for keyword in keywords):
        return True
    
    # Fuzzy matching
    for keyword in keywords:
        keyword = keyword.lower()
        if any(ratio >= threshold for ratio in [
            fuzz.ratio(text, keyword),
            fuzz.partial_ratio(text, keyword),
            fuzz.token_sort_ratio(text, keyword)
        ]):
            return True
    return False

def categorize_description(description):
    """Kategorisasi deskripsi dengan fuzzy matching dan prioritas yang tepat"""
    description = str(description).lower()
    
    # Define categories and their keywords
    categories = {
        'GALON': ['aqua', 'galon', 'isi ulang', 'air minum', 'gallon'],
        'BERAS': ['beras'],
        'MINI TRAINING': ['mini training', 'training', 'pelatihan'],
        'JUMSIH': ['jumsih', 'jumat bersih', "jum'at", 'jum at', 'bersih'],
        'SYUKURAN': ['syukuran', 'syukur'],
        'LAINNYA': []
    }
    
    # Count matched categories
    matches = {}
    for category, keywords in categories.items():
        if is_similar(description, keywords):
            matches[category] = True
    
    # If multiple categories match, return LAINNYA
    if len(matches) > 1:
        return 'LAINNYA'
    # If one category matches, return that category
    elif len(matches) == 1:
        return list(matches.keys())[0]
    # If no matches, return LAINNYA
    else:
        return 'LAINNYA'

def process_data(file):
    # Read the Excel file
    df = pd.read_excel(file)
    
    # Convert Transaction Date to datetime
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'], errors='coerce')
    df['TRANS. DATE'] = df['TRANS. DATE'].dt.strftime('%d/%m/%Y')
    
    # Sort the data in ascending order by TRANS. DATE
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'], format='%d/%m/%Y')
    df = df.sort_values('TRANS. DATE')
    
    # Add CATEGORY column
    df['CATEGORY'] = df['DESCRIPTION'].apply(categorize_description)
    
    # Create month-year column with custom formatting
    df['MONTH-YEAR'] = df['TRANS. DATE'].dt.strftime('%b, %Y')
    
    # Prepare summary and category-specific dataframes
    summary_df = df.groupby('MONTH-YEAR')['DEBIT'].sum().reset_index()
    
    # Category-specific monthly pivot
    category_pivot = df.groupby(['MONTH-YEAR', 'CATEGORY'])['DEBIT'].sum().unstack(fill_value=0).reset_index()
    
    # Separate dataframes for each category
    category_dfs = {}
    for category in df['CATEGORY'].unique():
        category_dfs[category] = df[df['CATEGORY'] == category]
    
    return {
        'original': df,
        'summary': summary_df,
        'category_pivot': category_pivot,
        'category_dfs': category_dfs
    }

def export_to_excel(processed_data):
    # Create a BytesIO object to save the Excel file
    output = io.BytesIO()
    
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write Summary sheet
        processed_data['summary'].to_excel(writer, sheet_name='SUMMARY', index=False)
        
        # Write Category Pivot sheet
        processed_data['category_pivot'].to_excel(writer, sheet_name='CATEGORY PIVOT', index=False)
        
        # Write individual category sheets
        for category, df in processed_data['category_dfs'].items():
            df.to_excel(writer, sheet_name=category, index=False)
        
        # Write original data sheet
        processed_data['original'].to_excel(writer, sheet_name='ORIGINAL DATA', index=False)
    
    # Seek to the beginning of the BytesIO object
    output.seek(0)
    
    return output

def main():
    st.title('Kategorisasi Data Konsumsi-55300000')
    st.write("""File ini berisikan data transaksi by account konsumsi/553000000 (Beras, Air Minum, Air Galon, Kopi, Gula, Teh, Syukuran Kantor, Mini Training, dll)""")
    st.write("""Rapihkan header dan footer, dan untuk header. cek terlebih dahulu karena pasti ada karakter spesial""")
    st.write("""Untuk data header seperti berikut: | VOUCHER NO. | TRANS. DATE | DESCRIPTION | DEBIT |""")
    
    # Add custom keywords option
    st.subheader("Custom Category Keywords (Optional)")
    
    with st.expander("Customize Category Keywords"):
        custom_keywords = {}
        for category in ['GALON', 'BERAS', 'MINI TRAINING', 'JUMSIH', 'LAINNYA']:
            custom_input = st.text_input(
                f"Keywords for {category} (comma separated):",
                value="",
                key=f"custom_{category}"
            )
            if custom_input:
                custom_keywords[category] = [kw.strip() for kw in custom_input.split(',')]
    
    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Process the file
        processed_data = process_data(uploaded_file)
        
        # Display summary
        st.subheader('Summary')
        st.dataframe(processed_data['summary'])
        
        # Display category pivot
        st.subheader('Category Pivot')
        st.dataframe(processed_data['category_pivot'])
        
        # Export to Excel button
        if st.button('Export to Excel'):
            excel_file = export_to_excel(processed_data)
            
            st.download_button(
                label="Download Excel File",
                data=excel_file,
                file_name='categorized_data.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

if __name__ == '__main__':
    main()
