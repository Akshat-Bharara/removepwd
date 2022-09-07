import msoffcrypto
import io
import pandas as pd
import streamlit as st
from pyxlsb import open_workbook as open_xlsb

def to_excel(df):
    output = decrypted
    writer = pd.ExcelWriter(output)
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data


st.title('Remove Excel Passwords 🔑')

uploaded_file = st.file_uploader("Choose the excel file",type="xlsx")

pwd = st.text_input('Enter the password: ')

remove = st.button("Remove password")

if uploaded_file is not None and remove:
    
    name = uploaded_file.name
    st.write("Name of file: ",name)

    decrypted = io.BytesIO()

    with open(name, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=pwd)  # Use password
        file.decrypt(decrypted)

    df = pd.read_excel(decrypted)

    file_container = st.expander("Check your uploaded excel",expanded=True)
    st.write(df)
    
    
    df_xlsx = to_excel(df)
    st.download_button(label='📥 Download Excel',
                                    data=df_xlsx ,
                                    file_name= name)
