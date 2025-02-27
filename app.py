import pandas as pd
import streamlit as st
import io
import re
from thefuzz import fuzz

def is_similar(text, keywords, threshold=85):
    """Check string similarity with fuzzy matching"""
    text = str(text).lower()
    
    if any(keyword.lower() in text for keyword in keywords):
        return True
    
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
    """Extract rice quantity (kg) specifically after the word 'beras'"""
    description = str(description).lower()
    
    matches = re.findall(r'beras\s*(\d+)\s*(?:kg|kilogram|kilo|k\.?g\.?)', description)
    
    if not matches:
        kg_patterns = re.findall(r'(?:(?<=beras)\s*(\d+)\s*kg)|(?:(\d+)\s*kg\s*beras)', description)
        matches = [match for group in kg_patterns for match in group if match]
    
    if matches:
        return sum(int(match) for match in matches)
    
    return 0

def categorize_description(description):
    """Kategorisasi deskripsi dengan fuzzy matching dan prioritas yang tepat"""
    description = str(description).lower()
    
    categories = {
        'GALON': ['aqua', 'galon', 'isi ulang', 'air minum', 'gallon'],
        'BERAS': ['beras'],
        'MINI TRAINING': ['mini training', 'training', 'pelatihan'],
        'JUMSIH': ['jumsih', 'jumat bersih', "jum'at", 'jum at', 'bersih'],
        'SYUKURAN': ['syukuran', 'syukur'],
        'LAINNYA': []
    }
    
    if is_similar(description, ['beras'], threshold=85):
        return 'BERAS'
    
    matches = {}
    for category, keywords in categories.items():
        if category != 'BERAS' and is_similar(description, keywords):
            matches[category] = True
    
    if len(matches) != 1:
        return 'LAINNYA'
    else:
        return list(matches.keys())[0]

def process_data(file):
    df = pd.read_excel(file)
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'], errors='coerce')   
    df = df.sort_values('TRANS. DATE')
    df['CATEGORY'] = df['DESCRIPTION'].apply(categorize_description)
    df['MONTH-YEAR'] = df['TRANS. DATE'].dt.strftime('%b, %Y')
    
    df['RICE KG'] = df['DESCRIPTION'].apply(extract_rice_quantity)
    summary_df = df.groupby('MONTH-YEAR')['DEBIT'].sum().reset_index()
    category_pivot = df.groupby(['MONTH-YEAR', 'CATEGORY'])['DEBIT'].sum().unstack(fill_value=0).reset_index()
    
    rice_summary = df.groupby('MONTH-YEAR').apply(
        lambda x: pd.Series({
            'BERAS_CATEGORY_KG': x[x['CATEGORY'] == 'BERAS']['RICE KG'].sum(),
            'LAINNYA_WITH_RICE_KG': x[(x['CATEGORY'] == 'LAINNYA') & (x['RICE KG'] > 0)]['RICE KG'].sum(),
            'TOTAL_RICE_KG': x['RICE KG'].sum()
        })
    ).reset_index()
    
    category_dfs = {}
    for category in df['CATEGORY'].unique():
        category_dfs[category] = df[df['CATEGORY'] == category]
    
    category_dfs['LAINNYA_WITH_RICE'] = df[(df['CATEGORY'] == 'LAINNYA') & (df['RICE KG'] > 0)]
    
    return {
        'original': df,
        'summary': summary_df,
        'category_pivot': category_pivot,
        'rice_summary': rice_summary,
        'category_dfs': category_dfs
    }

def export_to_excel(processed_data):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        processed_data['summary'].to_excel(writer, sheet_name='SUMMARY', index=False)
        processed_data['rice_summary'].to_excel(writer, sheet_name='JUMLAH BERAS', index=False)
        
        for category, df in processed_data['category_dfs'].items():
            sheet_name = category
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        processed_data['original'].to_excel(writer, sheet_name='ORIGINAL DATA', index=False)
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
