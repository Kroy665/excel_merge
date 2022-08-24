import streamlit as st
import pandas as pd
import openpyxl


uploaded_file = st.file_uploader("Choose a file", type=["xlsx"])
if uploaded_file is not None:
    dfxl = pd.ExcelFile(uploaded_file)
    sheets = dfxl.sheet_names
    sheetsLength = len(sheets)
    i=0
    for sheet in sheets:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        df.to_excel(f'temp{i}.xlsx')
        i+=1


if st.button("Merge"):
    writer = pd.ExcelWriter("my_excel_file"+'.xlsx')
    for i in range(sheetsLength):                 # loop through Excel files
        df = pd.read_excel(f'temp{i}.xlsx',engine='openpyxl',sheet_name=f'Sheet1')
        df_total = pd.read_excel(f'merged{i}.xlsx',engine='openpyxl',sheet_name=f'Sheet1')
        df_total = df_total.append(df)

        mergeWrite = pd.ExcelWriter(f'merged{i}.xlsx')
        df_total.to_excel(mergeWrite,sheet_name='Sheet1')
        mergeWrite.save()

        df_total.to_excel(writer, sheet_name=f'Sheet{i+1}')
        writer.save()
        # df_total.to_excel(f'merged{i}.xlsx')
    st.success("Merged successfully")
    # create a download button for my_excel_file.xlsx
    with open("my_excel_file.xlsx", "rb") as file:
        st.download_button(
            label="Download Excel File",
            data=file,
            file_name='my_excel_file.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )