import msoffcrypto
import io
import pandas as pd
import streamlit as st

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

    file_container = st.expander("Check your uploaded .csv",expanded=True)
    st.write(df)


    st.download_button(
     label="Download excel file",
     data=df,
     file_name=name+".xlsx",
     mime='excel',
    )

    st.success("Password has been removed") 

