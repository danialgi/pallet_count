import streamlit as st
import pandas as pd
import plotly.express as px
import webbrowser as wb
import openpyxl
import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from natsort import natsort_keygen

st.set_page_config(page_title="Pallet Count", page_icon="🚚", layout="wide")

st.title("🚚 Genuine Inside (M) Sdn. Bhd. - Pallet Count🧇")
st.markdown("##")

st.header("WMS File")
wms_file = st.file_uploader(".xls file",type=['xls'])
df_wms = pd.read_html(wms_file)
df_wms=df_wms[0]
df_wms.columns = range(df_wms.shape[1])
df_wms = df_wms[[5, 7, 10]].copy()
st.write("UPLOAD SUCCESS")
#df_wms

st.markdown("#")
st.header("Pallet File")
pallet_file = st.file_uploader(".xlsx file",type=['xlsx'])
df_pallet= pd.read_excel(pallet_file, sheet_name="Final")
df_pallet = df_pallet[['Row Labels', 'Qty per pallet','Qty per drawer']].copy()
st.write("UPLOAD SUCCESS")
#df_pallet

st.write("______________________________________________________________________________________")
num_rows = len(df_wms.index)
dfq={}

for i in range(num_rows):
    cell_value = df_wms.iat[i, 0]
    dfq[i] = df_pallet[df_pallet['Row Labels'] == cell_value]
    if dfq[i].empty:
        dfq[i] = pd.DataFrame({'Row Labels': [None], 'Qty per pallet': [None],  'Qty per drawer': [None]})


dfq_con = pd.concat(dfq.values(), ignore_index=True)
#dfq_con

wms_pallet= pd.concat([df_wms, dfq_con], axis=1)
wms_pallet = wms_pallet.drop('Row Labels', axis=1)
wms_pallet["Pallet Count"]=wms_pallet[10]/wms_pallet["Qty per pallet"]
wms_pallet["Drawer Count"]=wms_pallet[10]/wms_pallet["Qty per drawer"]
wms_pallet["Pallet Count"] = np.ceil(wms_pallet["Pallet Count"])
wms_pallet["Drawer Count"] = np.ceil(wms_pallet["Drawer Count"])
wms_pallet.rename(columns={5: 'Product', 7: 'Product Name', 10: 'Total'}, inplace=True)
wms_pallet

from io import BytesIO

# Function to write DataFrames to an Excel file in memory
def dfs_to_excel(df_list, sheet_list):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for dataframe, sheet in zip(df_list, sheet_list):
            dataframe.to_excel(writer, sheet_name=sheet, index=False)
    output.seek(0)
    return output

df_list = [wms_pallet]
sheet_list = ['Sheet1']

# Convert DataFrames to Excel in memory
excel_file = dfs_to_excel(df_list, sheet_list)

# Streamlit download button
st.download_button(
    label="Download Excel file",
    data=excel_file,
    file_name="Kimma_Billing.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
