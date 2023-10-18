import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime

st.set_page_config(
    page_title="DATA analytics App",
    page_icon="🛶",
    layout="wide"
)

st.title("DEMO DATA ANALYTIC APP")


today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)

start_date = st.sidebar.date_input("Start date",yesterday)
end_date =st.sidebar.date_input("end date",today)


report,uploadfile = st.tabs(["REPORT","UPLOAD FILE"])

with report:
    df = pd.read_excel("./db/file.xlsx")
    df['stopDate'] = pd.to_datetime(df['StopTime'])  
    mask = (df['stopDate'] >= str(start_date)) & (df['stopDate'] < str(end_date+datetime.timedelta(days=1)))
    df_process = df.loc[mask]
 
    df_plot = df_process[["stopDate","Band V bright"]]

    fig, ax = plt.subplots()
    ax.plot(df_plot["stopDate"], df_plot["Band V bright"])
    st.pyplot(fig)

    st.dataframe(df_plot)


with uploadfile:
    with st.form("upload_file"):
                        uploaded_file = st.file_uploader("Choose a file")
                        if uploaded_file is not None:
                            df_upload_file = pd.read_excel(uploaded_file)
                            df_upload_file.to_excel("./db/file.xlsx")
                            
                        submitted = st.form_submit_button("Submit")
                        if submitted:
                            if uploaded_file is not None:
                                st.dataframe(df_upload_file,width=1000)
