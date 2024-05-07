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

partner_option=['Zucca',
    'ViewnetMono',
    'HP & Seagate',
    'Harman Kardon',
    'Rakuten Kobo',
    'Earth Home',
    'iRobot',
    'Power Root',
    'Paseo',
    'Galaxy Sports',
    'NekoTech',
    #'Asia Century Supplies Sdn Bhd',
    'CMC Plus Plt',
    'CommBax Sdn Bhd',
    'Healthy Passion Wellnes Sdn Bhd (Marna)',
    'Kimma Sdn Bhd',
    'Kimma Sdn Bhd - outlet',
    'Healthy World Lifestyle Sdn Bhd (Ogawa)',
    #'Akademi Sempoa & Mental -Aritmetik Ucmas',
    'Mejorcare Sdn Bhd',
    #'Is Distributions Sdn Bhd',
    #'Grow Beyond Consulting Sdn Bhd',
    #'Dou Dou Trading',
    #'Jacko Agriculture Resources Sdn. Bhd.',
    #'Beast Kingdom (Malaysia) Sdn Bhd',
    'OBA Creative Sdn Bhd',
    'Nanjing Quka Pet Products Co Ltd',
    'Homelection (M) Sdn Bhd',
    'Leapro Fashion',
    'EEPRO MALAYSIA SDN BHD',
    'Twinings',
    'Connell (Nour by Nature)',
    'VICTOR SPORTS',
    'South Ocean',
    ]
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
    df_cart = df_cart[df_cart['Delivery Method'] == "By BRP Warehouse"]
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
#############################
    id_oc = df_cart [['Order ID']].copy()
    id_oc = id_oc.drop_duplicates(keep='first')

    id_wms = df_wms [['Order No.']].copy()
    id_wms = id_wms.drop_duplicates(keep='first')

    only_oc = id_oc[~id_oc['Order ID'].isin(id_wms['Order No.'])]
    only_oc=only_oc.reset_index()
    only_oc = only_oc.drop(['index'], axis=1)

    only_wms = id_wms[~id_wms['Order No.'].isin(id_oc['Order ID'])]
    only_wms=only_wms.reset_index()
    only_wms = only_wms.drop(['index'], axis=1)
################################
    order_id = df_cart [['Order ID']].copy()
    #order_id = pd.concat([df_cart['Order ID'], df_wms['Order No.']])
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
    st.markdown("__________________________________________________________________________")
    return df_wms_i

def revenue(df, column, percent):
    df[column] = df[column].replace('-', 0)
    total = df[column].sum()
    revenue = total*percent/100
    st.write(f"{percent}%", column)
    st.write("Revenue: ", revenue)
    return revenue

def rate_card(df):
    rate_card = df['Order ID'].nunique()
    st.write("(Rate Card) Orders: ", rate_card)
    return rate_card

def on_demand(df, name, column):
    id_unique = df[column].nunique()
    st.write(f"{name}", id_unique)
    return id_unique

def formula_match(df, column_df, sheet, column_formula):
    formula_file = 'Formula.xlsx'
    df_formula = pd.read_excel(formula_file, sheet_name=sheet, engine='openpyxl')
    df_formula = df_formula.drop_duplicates(subset=[column_formula], keep='first')

    df_rows=len(df)
    df_formula_i={}

    for i in range(df_rows):
        cell_value = df.at[i, column_df]
        df_formula_i[i] = df_formula[df_formula[column_formula] == cell_value]
        if df_formula_i[i].empty:
            df_formula_i[i] = pd.DataFrame({column_df: [None]})

    df_formula_i = pd.concat(df_formula_i)
    df_formula_i .reset_index(inplace=True)
    df_formula_i  = df_formula_i .drop(['level_0','level_1'], axis=1)

    df_concat= pd.concat([df, df_formula_i], axis=1, ignore_index=True)
    df_concat

    df_empty = df_concat[df_concat[36].isnull()]
    df_empty  = df_empty [[17]].copy()
    df_empty = df_empty.drop_duplicates(keep='first')
    st.write("*MISSING FORMULA*", df_empty)
    st.markdown("##")

    #df_concat_unique = df_concat[12].nunique()
    #df_concat_unique
    return df_concat

def cal_weight(df):
    df["Weight"]=df[21]*df[37]

    sum_weights_by_product = df.groupby(12, as_index=False).agg({'Weight': 'sum'})
    df = pd.merge(df, sum_weights_by_product, on=12, suffixes=('', '_sum'))
    df = df.drop_duplicates(subset=[12], keep='first')

    df_0kg = df[df['Weight_sum'].between(0, 2.999)]
    name="0-3kg"
    column=12
    df_0kg_rows=on_demand(df_0kg, name, column)

    df_3kg = df[df['Weight_sum'].between(3, 4.999)]
    name="3-5kg"
    df_3kg_rows=on_demand(df_3kg, name, column)

    df_5kg = df[df['Weight_sum'].between(5, 10.999)]
    name="5-10kg"
    df_5kg_rows=on_demand(df_5kg, name, column)

    df_10kg = df[df['Weight_sum'].between(10, 14.999)]
    name="10-15kg"
    df_10kg_rows=on_demand(df_10kg, name, column)

    df_15kg = df[df['Weight_sum']>=15]
    name="above 15kg"
    df_15kg_rows=on_demand(df_15kg, name, column)


data = oc_data()

if partner == 'Zucca':
    status=["Canceled","Canceled Reversal", "Refunded", "Returned", "Pending"]
    data = exclude_status(True, data, status)

    column='Total'
    percent=6
    total=revenue(data, column, percent)

if partner == 'ViewnetMono' or partner =='HP & Seagate' or partner == 'Harman Kardon' or partner == 'Rakuten Kobo':
    status=["Canceled","Canceled Reversal", "Refunded", "Returned", "Pending"]
    data = exclude_status(True, data, status)

    column='Order Income (RM)'
    percent=1.5
    total=revenue(data, column, percent)

if partner == 'Earth Home' or partner == 'iRobot' or partner == 'Power Root' or partner == 'Paseo' or partner == 'CMC Plus Plt' or partner == 'Twinings' or partner == 'Connell (Nour by Nature)' or partner == 'Mejorcare Sdn Bhd':
    status=["Pending"]
    data = exclude_status(True, data, status)

    orders=rate_card(data)

if partner == 'Galaxy Sports' or partner == 'VICTOR SPORTS':
    status=["Pending"]
    data1 = exclude_status(True, data, status)

    orders=rate_card(data1)

    cashback_status=["Canceled","Canceled Reversal", "Pending"]
    data2 = exclude_status(True, data, cashback_status)

    data2 = data2[data2['Category'].str.contains('Badminton Rackets')]
    cashback=data2['Quantity'].sum()
    st.write("Badminton Rackets:", cashback)



if partner == 'NekoTech':
    status=["Canceled","Canceled Reversal", "Pending"]
    data = exclude_status(True, data, status)

    column='Order Income (RM)'
    percent=1.5
    total=revenue(data, column, percent)

if partner == 'Kimma Sdn Bhd':
    data = exclude_status(False, data, status=[])
    data = matching(data)

    column_df='Item Description'
    sheet='Kimma weight'
    column_formula='Log'
    data = formula_match(data, column_df, sheet, column_formula)

    seborin = data[data[17].str.contains('SEBORIN', na=False)]
    seborin_INDEX=seborin.index
    data.drop(seborin_INDEX, inplace=True)
    name="Seborin"
    column=12
    seborin_rows=on_demand(seborin, name, column)

    single = data.drop_duplicates(subset=[12], keep=False)
    single = single[single[21] == 1]
    single_INDEX=single.index
    data.drop(single_INDEX, inplace=True)
    name="Single"
    single_rows=on_demand(single, name, column)

    cal_weight(data)

if partner == 'Kimma Sdn Bhd - outlet':
    data = exclude_status(False, data, status=[])
    data = matching(data)

    column_df='Item Description'
    sheet='Kimma weight'
    column_formula='Log'
    data = formula_match(data, column_df, sheet, column_formula)

    cal_weight(data)

if partner =='OBA Creative Sdn Bhd' or partner =='Nanjing Quka Pet Products Co Ltd' or partner =='Homelection (M) Sdn Bhd' or partner =='CommBax Sdn Bhd':
    status=[]
    data = exclude_status(False, data, status)

    data=matching(data)

    name="On Demand"
    column='Order No.'
    rows=on_demand(data, name, column)

if partner == 'Healthy Passion Wellnes Sdn Bhd (Marna)':
    status=[]
    data = exclude_status(False, data, status)

    data=matching(data)

    selfcollect = data[data['Truck No.'].str.contains('SELFCOLLECT', na=False)]
    name="Self Collect"
    column='Order No.'
    rows1=on_demand(selfcollect, name, column)

    selfcollect_INDEX=selfcollect.index
    data.drop(selfcollect_INDEX, inplace=True)

    name="Web/MP"
    rows2=on_demand(data, name, column)

if partner == 'Healthy World Lifestyle Sdn Bhd (Ogawa)':
    status=[]
    data = exclude_status(False, data, status)

    data=matching(data)

    column_df='Item No.'
    sheet='Ogawa SKU Mar24'
    column_formula='Log'
    data = formula_match(data, column_df, sheet, column_formula)
    total=data[37].sum()
    st.write("Handling: RM", total)

    data['return_rm'] = data[37].map({2.5: 3, 4: 8, 7: 10})
    total_return=data['return_rm'].sum()
    st.write("Return: RM", total_return)

if partner == 'Leapro Fashion':
    status=["Canceled","Canceled Reversal", "Refunded", "Returned", "Pending"]
    data1 = exclude_status(False, data, status)

    MP = data1[data1['Order Source'] != 'Web']
    Web = data1[data1['Order Source'] == 'Web']

    "MARKETPLACE"
    column_MP='Order Income (RM)'
    percent_MP=6
    total_MP=revenue(MP, column_MP, percent_MP)
    st.markdown("#")

    "WEB"
    column_Web='Cost Price'
    percent_Web=3
    total_Web=revenue(Web, column_Web, percent_Web)
    st.markdown("#")

    cancel_status=["Canceled","Canceled Reversal", "Refunded", "Returned"]
    data2=data[data['Order Status'].isin(cancel_status)]

    "RETURN"
    column2='Total'
    percent2=3
    total2=revenue(data2, column2, percent2)
    st.markdown("#")

if partner == 'EEPRO MALAYSIA SDN BHD':
    status=["Canceled","Canceled Reversal", "Refunded", "Returned", "Pending"]
    data1 = exclude_status(False, data, status)

    "WEB/MP"
    column1='Total'
    percent1=6
    total1=revenue(data1, column1, percent1)
    st.markdown("#")

    cancel_status=["Canceled","Canceled Reversal", "Refunded", "Returned"]
    data2=data[data['Order Status'].isin(cancel_status)]

    "RETURN"
    column2='Total'
    percent2=3
    total2=revenue(data2, column2, percent2)
    st.markdown("#")

if partner == 'South Ocean':
    status=["Pending"]
    data1 = exclude_status(False, data, status)

    "WEB/MP"
    column1='Total'
    percent1=3
    total1=revenue(data1, column1, percent1)
    st.markdown("#")

    cancel_status=["Canceled","Canceled Reversal", "Refunded", "Returned"]
    data2=data[data['Order Status'].isin(cancel_status)]

    "RETURN"
    column2='Total'
    percent2=3
    total2=revenue(data2, column2, percent2)
    st.markdown("#")
