import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px


#icon
#https://streamlit-emoji-shortcodes-streamlit-app-gwckff.streamlit.app/
st.set_page_config(
    page_title="DATA analytics App",
    page_icon="🛶",
    layout="wide"
)

st.title("TITLE")
st.header("header")
st.subheader("subheader")
st.write("Hello world")
st.caption("caption")




data = pd.DataFrame(data=np.random.rand(4,4)*100000,
index=["Q1","Q2","Q3","Q4"],
columns=["East","West","North","South"])
data.index.name = "Quaters"
data.columns.name = "Region"
fig = data.plot.bar().get_figure()
st.pyplot(fig)


df = px.data.gapminder()

fig = px.scatter(
    df.query("year==2007"),
    x="gdpPercap",
    y="lifeExp",
    size="pop",
    color="continent",
    hover_name="country",
    log_x=True,
    size_max=60,
)
st.plotly_chart(fig, theme="streamlit", use_container_width=True)


if st.button("Submit"):
    st.write("Click submit")

def convert_df(df):
    return df.to_csv().encode('utf-8')

csv = convert_df(df)

st.download_button(
    label="download data csv",
    data=csv,
    file_name='df.csv',
    mime='text/csv',
)


df = pd.DataFrame(np.random.randn(50, 20), columns=("col %d" % i for i in range(20)))
st.dataframe(df,width=1500)


col1,col2,col3 = st.columns(3)

with col1:
    st.header("col1")
    st.area_chart(df)

with col2:
    st.header("col2")
    st.line_chart(df)
    
with col3:
    st.header("col3")
    st.scatter_chart(df)

col1,col2 = st.columns(2)

with col1:
    st.header("col1")
    st.area_chart(df)

with col2:
    st.header("col2")
    st.line_chart(df)

with st.expander("See explanation1"):
    st.dataframe(df,width=1500)

contact_lists = ["Email","Home phone","Mobile phone"]

selectbox_value = st.selectbox("select contacted",
                    contact_lists,
                    index=None,
                    placeholder="select contact method")
st.write(selectbox_value)


st.sidebar.title("DATA ANALYTICS APP")
with st.sidebar.form("my_form"):
    multi_value = st.multiselect("select contacted",
                        contact_lists,
                        placeholder="select contact method"
    )

    submit = st.form_submit_button("SUBMIT")
    if submit:
        st.write(multi_value)

tab1,tab2 = st.tabs(["tab1","tab2"])

with tab1:
    st.header("tab1")
with tab2:
    st.header("tab2")
    with st.form("upload_file"):
                        uploaded_file = st.file_uploader("Choose a file")
                        if uploaded_file is not None:
                            df_upload_file = pd.read_excel(uploaded_file)

                        submitted = st.form_submit_button("Submit")
                        if submitted:
                            if uploaded_file is not None:
                                st.dataframe(df_upload_file,width=1000)