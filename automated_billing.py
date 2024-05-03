import streamlit as st
import pandas as pd
import plotly.express as px
import webbrowser as wb
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import io
import calendar
from datetime import datetime

today_date = datetime.now().strftime('%Y-%m-%d')

st.set_page_config(page_title="GI Billing", page_icon="🚚", layout="wide")
st.write("🚚 Genuine Inside (M) Sdn. Bhd.")
st.title("Automated Billing🧾")
st.markdown("##")

partner_option=['Zucca', 'ViewnetMono', 'HP & Seagate', 'Harman Kardon', 'Rakuten Kobo', 'Earth Home', 'iRobot', 'Power Root', 'Paseo', 'Galaxy Sports', 'NekoTech']
partner_option.sort()
partner = st.selectbox("Partner: ", partner_option)
st.write("_________________________________________________________________________________________________")


def oc_data():
    cart_file = st.file_uploader("Open Cart file",type=['xlsx'])
    df_cart = pd.read_excel(cart_file)
    df_cart = df_cart.drop([0, 1, 2, 3])
    df_cart.columns = df_cart.iloc[0]
    df_cart = df_cart[1:]
    df_cart
    df_cart = df_cart.ffill()
    st.markdown("##")
    return df_cart

def exclude_status(exclude, df, status):
    if exclude == False:
        df = df[df['Order Status'] == "Complete"]
    else:
        df = df[~df['Order Status'].isin(status)]
        st.write(f"Exclude: {status}")

    return df

def matching(df_cart):
    wms_file = st.file_uploader("WMS file",type=['xls'])
    df_wms = pd.read_html(wms_file)
    df_wms = df_wms[0]
    df_wms = df_wms[df_wms['Status'] == "COMPLETED"]
    df_wms
    st.write("_____________________________________________________________________")

    order_id = pd.concat([df_cart['Order ID'], df_wms['Order No.']])
    order_id = order_id.drop_duplicates(keep='first')
    order_id=order_id.reset_index()
    order_id = order_id.drop(['index'], axis=1)

    num_rows = len(order_id)
    df_cart_i={}
    df_wms_i={}

    for i in range(num_rows):
        cell_value = order_id.iat[i, 0]
        df_wms_i[i] = df_wms[df_wms['Order No.'] == cell_value]

    df_wms_i = pd.concat(df_wms_i)
    df_wms_i = df_wms_i.reset_index()
    df_wms_i = df_wms_i.drop(['level_0', 'level_1'], axis=1)
    st.write("WMS Matched: ", df_wms_i)
    st.markdown("##")
    return df_wms_i

def revenue(df, column, percent):
    df[column] = df[column].replace('-', 0)
    total = df[column].sum()
    revenue = total*percent/100
    st.write(f"{percent}%")
    st.write("Column: ", column)
    st.write("Revenue: ", revenue)
    return revenue

def rate_card(df):
    rate_card = df['Order ID'].nunique()
    st.write("(Rate Card) Orders: ", rate_card)
    return rate_card

def on_demand(df):
    on_demand=1

data = oc_data()

if partner == 'Zucca':
    status=["Canceled","Canceled Reversal", "Refunded", "Returned", "Pending"]
    column='Total'
    percent=6
    data = exclude_status(True, data, status)
    total=revenue(data, column, percent)

if partner == 'ViewnetMono' or partner =='HP & Seagate' or partner == 'Harman Kardon' or partner == 'Rakuten Kobo':
    status=["Canceled","Canceled Reversal", "Refunded", "Returned", "Pending"]
    column='Order Income (RM)'
    percent=1.5
    data = exclude_status(True, data, status)
    total=revenue(data, column, percent)

if partner == 'Earth Home' or partner == 'iRobot' or partner == 'Power Root' or partner == 'Paseo' or partner == 'Galaxy Sports':
    status=["Pending"]
    data = exclude_status(True, data, status)
    orders=rate_card(data)

if partner == 'NekoTech':
    status=["Canceled","Canceled Reversal", "Pending"]
    column='Order Income (RM)'
    percent=1.5
    data = exclude_status(True, data, status)
    total=revenue(data, column, percent)

if partner == 'Kimma':
    data = exclude_status(False, data, status=[])
    data = matching(data)
