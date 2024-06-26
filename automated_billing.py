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
import datetime

today_date = datetime.date.today()

st.set_page_config(page_title="GI Billing", page_icon="🚚", layout="wide")
st.write("🚚 Genuine Inside (M) Sdn. Bhd.")
st.title("Automated Billing🧾")
st.markdown("##")

partner_option=['Zucca Commerce Sdn. Bhd.',
    'Mono Digital Sdn Bhd (ViewnetMono)',
    'HP & Seagate',
    'Harman Kardon',
    'Rakuten Kobo',
    'Earth Home',
    'iRobot',
    'Power Root',
    'Paseo',
    'Galaxy Sports',
    'NekoTech',
    'Asia Century Supplies Sdn Bhd',
    'CMC Plus Plt',
    'CommBax Sdn Bhd',
    'Healthy Passion Wellnes Sdn Bhd (Marna)',
    'Kimma Sdn Bhd',
    #'Kimma Sdn Bhd - outlet',
    'Healthy World Lifestyle Sdn Bhd (Ogawa)',
    #'Akademi Sempoa & Mental -Aritmetik Ucmas',
    'Mejorcare Sdn Bhd',
    'Is Distributions Sdn Bhd',
    'Grow Beyond Consulting Sdn Bhd',
    'Dou Dou Trading',
    'Jacko Agriculture Resources Sdn. Bhd.',
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

def oc_data(name):
    cart_file = st.file_uploader(name,type=['xlsx'])
    df_cart = pd.read_excel(cart_file)
    if partner == 'Mejorcare Sdn Bhd':
        df_cart = df_cart.drop([0, 1, 2, 3, 4])
    else:
        df_cart = df_cart.drop([0, 1, 2, 3])
    df_cart.columns = df_cart.iloc[0]
    df_cart = df_cart[1:]
    df_cart['Order ID'] = df_cart['Order ID'].ffill()
    df_cart['Date Added'] = df_cart['Date Added'].ffill()
    if 'Order Source' in df_cart.columns:
        df_cart['Order Source'] = df_cart['Order Source'].ffill()
    df_cart['Order Status'] = df_cart['Order Status'].ffill()
    df_cart['Delivery Method'] = df_cart['Delivery Method'].ffill()
    if 'Tracking' in df_cart.columns:
        df_cart['Tracking'] = df_cart['Tracking'].ffill()
    if 'Tracking No.' in df_cart.columns:
        df_cart['Tracking No.'] = df_cart['Tracking No.'].ffill()
    df_cart['Courier'] = df_cart['Courier'].ffill()
    df_cart = df_cart[df_cart['Delivery Method'] == "By BRP Warehouse"]
    st.markdown("##")
    return df_cart

def exclude_status(df, status):
    df = df[~df['Order Status'].isin(status)]
    st.write(f"Exclude: {status}")

    return df

def matching(df_cart):
    if partner == 'Is Distributions Sdn Bhd' or partner == 'Grow Beyond Consulting Sdn Bhd':
        wms_file = st.file_uploader("WMS file",type=['csv'])
        df_wms = pd.read_csv(wms_file)
    else:
        wms_file = st.file_uploader("WMS file",type=['xls'])
        df_wms = pd.read_html(wms_file)
        df_wms = df_wms[0]

    df_wms = df_wms[df_wms['Status'] == "COMPLETED"]

    order_id = df_cart [['Order ID']].copy()
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
    formula_file = 'Formular BRP Billing 2023 (2).xlsx'
    df_formula = pd.read_excel(formula_file, sheet_name=sheet, engine='openpyxl')
    df_formula = df_formula.drop_duplicates(subset=[column_formula], keep='first')

    df_rows=len(df)
    df_formula_i={}

    for i in range(df_rows):
        cell_value = df.at[i, column_df]
        if str(cell_value).isdecimal():
            df_formula_i[i] = df_formula[df_formula[column_formula]==int(cell_value)]
        else:
            df_formula_i[i] = df_formula[df_formula[column_formula]==(cell_value)]
        #df_formula_i[i] = df_formula[df_formula[column_formula]==(cell_value)]
        if df_formula_i[i].empty:
            df_formula_i[i] = pd.DataFrame({column_formula: [None]})

    df_formula_i = pd.concat(df_formula_i)
    df_formula_i .reset_index(inplace=True)
    df_formula_i  = df_formula_i .drop(['level_0','level_1'], axis=1)

    df_concat= pd.concat([df, df_formula_i], axis=1, ignore_index=True)
    df_columns = df.shape[1]

    df_empty = df_concat[df_concat[df_columns+1].isnull()]
    if column_df == 'Model':
        df_empty  = df_empty [[6,7]].copy()
        df_empty.rename(columns={6: 'Item Name', 7: 'Item Code'}, inplace=True)
    else:
        df_empty  = df_empty [[16,17]].copy()
        df_empty.rename(columns={16: 'Item Code', 17: 'Item Name'}, inplace=True)
    df_empty = df_empty.drop_duplicates(keep='first')
    st.write("*MISSING FORMULA*", df_empty)
    st.markdown("##")

    #df_concat_unique = df_concat[12].nunique()
    #df_concat_unique
    return df_concat

def cal_weight(df, order_column, weight_column):
    sum_weights_by_product = df.groupby(order_column, as_index=False).agg({weight_column: 'sum'})
    df = pd.merge(df, sum_weights_by_product, on=order_column, suffixes=('', '_sum'))
    df = df.drop_duplicates(subset=[order_column], keep='first')

    df_0kg = df[df[f'{weight_column}_sum'].between(0, 2.999)]
    name="0-3kg"
    df_0kg_rows=on_demand(df_0kg, name, order_column)

    df_3kg = df[df[f'{weight_column}_sum'].between(3, 4.999)]
    name="3-5kg"
    df_3kg_rows=on_demand(df_3kg, name, order_column)

    df_5kg = df[df[f'{weight_column}_sum'].between(5, 9.999)]
    name="5-10kg"
    df_5kg_rows=on_demand(df_5kg, name, order_column)

    df_10kg = df[df[f'{weight_column}_sum'].between(10, 14.999)]
    name="10-15kg"
    df_10kg_rows=on_demand(df_10kg, name, order_column)

    df_15kg = df[df[f'{weight_column}_sum']>=15]
    name="above 15kg"
    df_15kg_rows=on_demand(df_15kg, name, order_column)


data = oc_data('Open cart file:')

if partner == 'Zucca Commerce Sdn. Bhd.':
    data2 = oc_data('2nd Open cart file:')

    status=["Canceled","Canceled Reversal", "Refunded", "Returned", "Pending"]
    data = exclude_status(data, status)
    data2 = exclude_status(data2, status)

    data_concat = pd.concat([data, data2])
    column='Total'
    percent=6
    total=revenue(data_concat, column, percent)

if partner == 'Mono Digital Sdn Bhd (ViewnetMono)' or partner =='HP & Seagate' or partner == 'Harman Kardon' or partner == 'Rakuten Kobo':
    status=["Canceled","Canceled Reversal", "Refunded", "Returned", "Pending"]
    data = exclude_status( data, status)

    column='Order Income (RM)'
    percent=1.5
    total=revenue(data, column, percent)

if partner == 'Earth Home' or partner == 'iRobot' or partner == 'Power Root' or partner == 'Paseo' or partner == 'CMC Plus Plt' or partner == 'Twinings' or partner == 'Connell (Nour by Nature)' or partner == 'Mejorcare Sdn Bhd':
    status=["Pending"]
    data = exclude_status(data, status)

    orders=rate_card(data)

if partner == 'Galaxy Sports' or partner == 'VICTOR SPORTS':
    status=["Pending"]
    data1 = exclude_status(data, status)

    orders=rate_card(data1)

    cashback_status=["Canceled","Canceled Reversal", "Pending"]
    data2 = exclude_status(data, cashback_status)

    data2 = data2[data2['Category'].str.contains('Badminton Rackets')]
    cashback=data2['Quantity'].sum()
    st.write("Badminton Rackets:", cashback)

if partner == 'NekoTech':
    status=["Canceled","Canceled Reversal", "Pending"]
    data = exclude_status(data, status)

    column='Order Income (RM)'
    percent=1.5
    total=revenue(data, column, percent)

if partner == 'Kimma Sdn Bhd':
    kimma_outlet=['00-HQ', '01-Taman Bukit Maluri', '02-Kepong Baru', '05-Taman Megah SS24', '06-Metro Perdana', '07-Petaling Jaya SS2/64', 'HD 10-Desa Aman Puri', 'HD 12-Selayang Jaya',  'HD 17 (F)-Muar Johor', 'HD 18-E-curve', 'HD 20-Bandar Sri Damansara', 'HD 21-Amcorp Mall','HD 22-Nu Sentral', 'HD 23-Mid Valley', 'HD 28-Time Square', 'HD 31-One Utama', 'HD 33-AEON Rawang', 'HD 34-Melawati Mall']

    data = matching(data)

    column_df='Item Description'
    sheet='Kimma weight'
    column_formula='Log'
    data = formula_match(data, column_df, sheet, column_formula)

    #OUTLET#
    outlet = data[data[15].isin(kimma_outlet)]
    outlet_INDEX=outlet.index
    data.drop(outlet_INDEX, inplace=True)
    #OUTLET#

    seborin = data[data[17].str.contains('SKF SEBORIN AKTIV HAIR TONIC 300ML : 12 CTN', na=False)]
    seb_row=seborin.shape[0]
    seborin_i={}
    for i in range(seb_row):
        cell_value = seborin.iat[i, 12]
        seborin_i[i] = data[data[12] == cell_value]
    seborin_i = pd.concat(seborin_i)
    seborin_i = seborin_i.drop_duplicates(subset=[12], keep=False)
    seborin_i = seborin_i.reset_index()
    seborin_i.set_index('level_1', inplace=True)
    seborin_i = seborin_i.drop(columns=['level_0'])
    seborin_i_INDEX=seborin.index
    data.drop(seborin_i_INDEX, inplace=True)
    name="Seborin"
    column=12
    seborin_rows=on_demand(seborin_i, name, column)

    single = data.drop_duplicates(subset=[12], keep=False)
    single = single[single[21] == 1]
    single_INDEX=single.index
    data.drop(single_INDEX, inplace=True)
    name="Single"
    single_rows=on_demand(single, name, column)

    data["Weight"]=data[21]*data[37]
    order_column=12
    weight_column='Weight'
    cal_weight(data, order_column, weight_column)

    "_________________________________________"
    "OUTLET"
    outlet["Weight"]=outlet[21]*outlet[37]
    order_column=12
    weight_column='Weight'
    cal_weight(outlet, order_column, weight_column)

if partner =='OBA Creative Sdn Bhd' or partner =='Nanjing Quka Pet Products Co Ltd' or partner =='Homelection (M) Sdn Bhd' or partner =='CommBax Sdn Bhd':
    data=matching(data)

    name="On Demand"
    column='Order No.'
    rows=on_demand(data, name, column)

if partner == 'Healthy Passion Wellnes Sdn Bhd (Marna)':
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
    data1=matching(data1)

    column_df1='Item No.'
    sheet1='Ogawa SKU Mar24'
    column_formula1='Log'
    data1 = formula_match(data1, column_df1, sheet1, column_formula1)
    total=data1[37].sum()
    st.write("Handling: RM", total)
    "#"

    cancel_status=['Returned']
    data2=data[data['Order Status'].isin(cancel_status)]

    column_df2='Model'
    sheet2='Ogawa SKU Mar24'
    column_formula2='Log'
    data2.reset_index(inplace=True)
    data2  = data2.drop(['index'], axis=1)
    data2 = formula_match(data2, column_df2, sheet2, column_formula2)
    data2['return_rm'] = data2[29].map({2.5: 3, 4: 8, 7: 10})
    total_return=data2['return_rm'].sum()
    st.write("Return: RM", total_return)


if partner == 'Leapro Fashion':
    status=["Canceled","Canceled Reversal", "Refunded", "Returned", "Pending"]
    data1 = exclude_status(data, status)

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
    data1 = exclude_status(data, status)

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
    data1 = exclude_status(data, status)

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

if partner == 'Is Distributions Sdn Bhd':
    data = matching(data)

    order_column='Order No.'
    weight_column='Box Weight'
    cal_weight(data, order_column, weight_column)

if partner == 'Grow Beyond Consulting Sdn Bhd':
    data = matching(data)

    selfcollect = data[data['Courier'] == 'Self collect']
    name="Self Collect"
    column='Order No.'
    rows1=on_demand(selfcollect, name, column)

    selfcollect_INDEX=selfcollect.index
    data.drop(selfcollect_INDEX, inplace=True)

    order_column='Order No.'
    weight_column='Box Weight'
    cal_weight(data, order_column, weight_column)

if partner == 'Dou Dou Trading':
    data1= matching(data1)

    order_column='Order No.'
    weight_column='Order Qty'
    cal_weight(data1, order_column, weight_column)

    cancel_status=['Returned']
    data2=data[data['Order Status'].isin(cancel_status)]
    name="Return"
    column='Order ID'
    rows1=on_demand(data2, name, column)

if partner == 'Jacko Agriculture Resources Sdn. Bhd.':
    data = matching(data)

    order_column='Order No.'
    weight_column='Order Qty'
    cal_weight(data, order_column, weight_column)

if partner == 'Asia Century Supplies Sdn Bhd':
    data = matching(data)

    column_df = 'Item No.'
    sheet = 'ACS SKU'
    column_formula = 'Log'
    data = formula_match(data, column_df, sheet, column_formula)
    data['Total']=data[21]*data[37]

    'Accessories'
    accessories = data[data[39] == 'Accesories']
    column='Total'
    percent=4
    total1=revenue(accessories, column, percent)
    '#'

    'Consumables'
    consumables = data[data[39] == 'Consumables']
    column='Total'
    percent=0.5
    total2=revenue(consumables, column, percent)
