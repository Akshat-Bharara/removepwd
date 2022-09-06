import msoffcrypto
import io
import pandas as pd
import os
import streamlit as st

st.title('Remove Excel Passwords 🔑')

uploaded_file = st.file_uploader("Choose the excel file",type="xlsx")

pwd = st.text_input('Enter the password: ')

remove = st.button("Remove password")

if uploaded_file is not None and remove:
    
    name = uploaded_file.name
    st.write("Name of file: ",name)

    try:
        decrypted = io.BytesIO()

        with open("C:\\Delete\\"+name, "rb") as f:
            file = msoffcrypto.OfficeFile(f)
            file.load_key(password=pwd)  # Use password
            file.decrypt(decrypted)

        df = pd.read_excel(decrypted)

        file_container = st.expander("Check your uploaded .csv")
        st.write(df)

        os.remove(name)

        df.to_excel(name,index=False) 

        with open("C:\\Delete\\"+name, "rb") as f:
            pass

        st.success("Password has been removed") 

    except:
        st.warning('Password has already been removed')
