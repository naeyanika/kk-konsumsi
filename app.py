import pandas as pd
import streamlit as st
import io
import re
from fuzzywuzzy import fuzz

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

def extract_rice_quantity(description):
    """Extract rice quantity (kg) from description"""
    description = str(description).lower()
    
    # Find all occurrences of numbers followed by "kg" near "beras"
    if "beras" in description:
        # Pattern: one or more digits, optional spaces, then "kg" or similar
        matches = re.findall(r'(\d+)(?:\s*)(?:kg|kilogram|kilo)', description)
        if matches:
            # Return the sum of all rice quantities
            return sum(int(match) for match in matches)
    
    return 0  # Return 0 if no rice or no quantity found

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
    
    # Check for "beras" first - it has highest priority
    if is_similar(description, ['beras'], threshold=85):
        return 'BERAS'
    
    # Count matched categories
    matches = {}
    for category, keywords in categories.items():
        if category != 'BERAS' and is_similar(description, keywords):
            matches[category] = True
    
    # If multiple categories match or none match, return LAINNYA
    if len(matches) != 1:
        return 'LAINNYA'
    # If one category matches, return that category
    else:
        return list(matches.keys())[0]

def process_data(file):
    # Read the Excel file
    df = pd.read_excel(file)
    
    # Convert Transaction Date to datetime
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'], errors='coerce')
    
    # Sort the data in ascending order by TRANS. DATE
    df = df.sort_values('TRANS. DATE')
    
    # Add CATEGORY column
    df['CATEGORY'] = df['DESCRIPTION'].apply(categorize_description)
    
    # Create month-year column with custom formatting
    df['MONTH-YEAR'] = df['TRANS. DATE'].dt.strftime('%b, %Y')
    
    # Add RICE KG column to track rice quantities
    df['RICE KG'] = df['DESCRIPTION'].apply(extract_rice_quantity)
    
    # Prepare summary and category-specific dataframes
    summary_df = df.groupby('MONTH-YEAR')['DEBIT'].sum().reset_index()
    
    # Category-specific monthly pivot
    category_pivot = df.groupby(['MONTH-YEAR', 'CATEGORY'])['DEBIT'].sum().unstack(fill_value=0).reset_index()
    
    # Create a rice summary sheet that includes both BERAS category and rice in LAINNYA
    rice_summary = df.groupby('MONTH-YEAR').apply(
        lambda x: pd.Series({
            'BERAS_CATEGORY_KG': x[x['CATEGORY'] == 'BERAS']['RICE KG'].sum(),
            'LAINNYA_WITH_RICE_KG': x[(x['CATEGORY'] == 'LAINNYA') & (x['RICE KG'] > 0)]['RICE KG'].sum(),
            'TOTAL_RICE_KG': x['RICE KG'].sum()
        })
    ).reset_index()
    
    # Separate dataframes for each category
    category_dfs = {}
    for category in df['CATEGORY'].unique():
        category_dfs[category] = df[df['CATEGORY'] == category]
    
    # Create a dataframe for LAINNYA that contains rice
    category_dfs['LAINNYA_WITH_RICE'] = df[(df['CATEGORY'] == 'LAINNYA') & (df['RICE KG'] > 0)]
    
    return {
        'original': df,
        'summary': summary_df,
        'category_pivot': category_pivot,
        'rice_summary': rice_summary,
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
        
        # Write Rice Summary sheet
        processed_data['rice_summary'].to_excel(writer, sheet_name='JUMLAH BERAS', index=False)
        
        # Write individual category sheets
        for category, df in processed_data['category_dfs'].items():
            sheet_name = category
            # Limit sheet name length to 31 characters (Excel limitation)
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
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
        
        # Display rice summary
        st.subheader('Jumlah Beras per Bulan')
        st.dataframe(processed_data['rice_summary'])
        
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
