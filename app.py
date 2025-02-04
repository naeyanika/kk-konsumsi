import pandas as pd
import streamlit as st
import io
import locale

def categorize_description(description):
    description = description.lower()
    
    # Exact match for single keywords
    if 'aqua' in description and description.strip() == 'aqua':
        return 'GALON'
    elif 'galon' in description and description.strip() == 'galon':
        return 'GALON'
    elif 'isi ulang' in description and description.strip() == 'isi ulang':
        return 'GALON'
    elif 'air' in description and description.strip() == 'air':
        return 'GALON'
    
    # Exact match for beras
    elif 'beras' in description and description.strip() == 'beras':
        return 'BERAS'
    
    # Exact match for mini training
    elif 'mini training' in description and description.strip() == 'mini training':
        return 'MINI TRAINING'
    
    # Exact match for jumsih
    elif 'jumsih' in description and description.strip() == 'jumsih':
        return 'JUMSIH'
    
    # Everything else goes to Lainnya
    else:
        return 'LAINNYA'

def process_data(file):
    # Read the Excel file
    df = pd.read_excel(file)
    
    # Convert Transaction Date to datetime
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'])
    
    # Add CATEGORY column
    df['CATEGORY'] = df['DESCRIPTION'].apply(categorize_description)
    
    # Create month-year column with custom formatting
    df['MONTH-YEAR'] = df['TRANS. DATE'].dt.strftime('%B, %Y')
    
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
    st.title('Excel Data Categorization and Export Tool')
    
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
