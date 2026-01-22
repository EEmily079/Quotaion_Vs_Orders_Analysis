import pandas as pd     
import plotly
from sqlalchemy import false
import streamlit as st
import plotly.express as px
df = pd.read_excel('SQ_ALLDB_Cleaned(P).xlsx')
st.set_page_config(page_title='Successful Sales Quotation Dashboard', page_icon=':bar_chart:', layout = 'wide')
st.sidebar.header ("Please Filter Here")



#Change datatype of 'Posting Date' to datetime and pull Month name
df['Posting Date'] = pd.to_datetime(df['Posting Date'], format='mixed')
df['Month'] = df['Posting Date'].dt.month_name()

# Create month order list and filter available months in the data
month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
available_months = [m for m in month_order if m in df['Month'].unique()]

#sidebar filters for Month
By_Month = st.sidebar.multiselect(
    'By Month',
    available_months,
    default=available_months
)

#Sort Country List by alphabetical order
country_list = df['GroupName'].dropna().astype(str).unique().tolist()
country_list.sort()
#Sidebar filters for End County 
By_End_County = st.sidebar.multiselect(
    'By End County',
    country_list,
    default=country_list
)

#Sort Vendor List by alphabetical order
vendor_list = df['Customer/Vendor Name'].dropna().astype(str).unique().tolist()
vendor_list.sort()
#Sidebar filters for Vendor Group Name
By_Vendor_Group_Name = st.sidebar.multiselect(
    'By Vendor Group Name',
    vendor_list,
    default= vendor_list
)

# Set default values if no selection is made
By_Month = By_Month or df["Month"].unique().tolist()
By_End_County = By_End_County or df["GroupName"].unique().tolist()
By_Vendor_Group_Name = By_Vendor_Group_Name or df["Customer/Vendor Name"].unique().tolist()

# Create selection based on sidebar filters
selection = (
    df["Month"].isin(By_Month)
    & df["GroupName"].isin(By_End_County)
    & df["Customer/Vendor Name"].isin(By_Vendor_Group_Name)
)
df_select = df.loc[selection]


st.header('Sales Analysis Dashboard')
st.markdown('---')

#partition for Cards Total 5 columns 
a1,a2,a3,a4,a5 = st.columns(5)
#variables for Cards
total_sq = df_select['SQ TOTAL'].sum()
no_of_sq = df_select['SQ NUM'].nunique()
total_so = df_select['SO TOTAL'].sum()
no_of_so = df_select['SO NUM'].nunique()
hit_rate = round((total_so/total_sq)*100, 2)

with a1:
    st.subheader('SQ Count')
    st.subheader(f'{no_of_sq}')

with a2:
     st.subheader('SQ Total (SGD)')
     st.subheader(f'{total_sq:,.2f}')
   
with a3:
    st.subheader('SO Count')
    st.subheader(f'{no_of_so}')

with a4:
    st.subheader('SO Total (SGD)')
    st.subheader(f'{total_so:,.2f}')

with a5:
    st.subheader('Hit Rate%')
    st.subheader(f'{hit_rate}%')

st.markdown('---')

#partition for 2 charts 
b1,b2 = st.columns(2)
#filter total sales by month by selected filters, default is all months(already defined in df_select) 
total_sales_by_month = df_select.groupby('Month',as_index=False)['SO TOTAL'].sum()
# Month in the correct order 
month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
total_sales_by_month['Month'] = pd.Categorical(total_sales_by_month['Month'], categories=month_order, ordered=True)
total_sales_by_month = total_sales_by_month.sort_values('Month')

# Line chart for Total Sales by Month 
fig1 = px.line(
    total_sales_by_month,
    x = total_sales_by_month['Month'],
    y = total_sales_by_month['SO TOTAL'],
    markers=True,
    labels={'x':'Month','y':'Total Sales (SGD)'},
    title='Total Sales by Month'
)
b1.plotly_chart(fig1,use_container_width=True)

#Filter total sales by group by selected filters, default is all groups(already defined in df_select)
sales_by_group = df_select.groupby(['GroupName'])['SO TOTAL'].sum().sort_values(ascending=False).reset_index()
#bar chart for Total Sales by group
fig2 = px.bar(
    sales_by_group,
    x= sales_by_group['GroupName'],
    y= sales_by_group['SO TOTAL'],
    title='Total Sales by Country'
)
b2.plotly_chart(fig2,use_container_width=True)

st.markdown('---')

#partition for 2 Pie charts 
c1,c2 = st.columns(2)

#filter total sales by customer by selected filters, default is all customers(already defined in df_select)
sales_by_cus = df_select.groupby(['Customer/Vendor Name'])['SO TOTAL'].sum().sort_values(ascending=False).reset_index()
sales_by_cus.rename(columns={'Customer/Vendor Name':'Customer Name'},inplace=True)
top_5_cus = sales_by_cus.head(5)
#Pie chart for Top 5 Leading Customer
fig3 = px.pie(
    top_5_cus,
    values='SO TOTAL',
    names='Customer Name',
    title='Top 5 Leading Customer'
)
c1.plotly_chart(fig3,width=1000)

#filter total sales by group by selected filters, default is all groups(already defined in df_select)
sales_by_group = df_select.groupby(['GroupName'])['SO TOTAL'].sum().sort_values(ascending=False).reset_index()
#sales_by_group.rename(columns={'GroupName':'Country'},inplace=True)
top_5_group = sales_by_group.head(5)
#Pie chart for Top 5 Leading Group
fig4 = px.pie(
    top_5_group,
    values='SO TOTAL',
    names='GroupName',
    title='Top 5 Leading Group'
)
c2.plotly_chart(fig4,width=1000)

st.markdown('---')

#whole section for SQ vs SO line chart
sq_so_success = df_select.groupby("GroupName")[["SQ TOTAL", "SO TOTAL"]].sum().reset_index()
#want to show hit rate as hover data so calculate hite rate and keep it in the new column 
sq_so_success["Hit Rate %"] = (sq_so_success["SO TOTAL"] / sq_so_success["SQ TOTAL"] * 100).round(2)

#2 lines chart for SQ vs SO by Group
fig5 = px.line(
    sq_so_success, 
    x="GroupName",  
    y=["SQ TOTAL", "SO TOTAL"], #this is 2 lines chart showing SQ TOTAL and SO TOTAL
    markers= True, 
    hover_data={"Hit Rate %": True}, #add hit rate as hover data
    title="Sales Quotation vs Sales Order by Group (Showing hit rate % on hover)")

st.plotly_chart(fig5,use_container_width=True)


st.markdown('---')

#Matrix table section
st.subheader('Dynamic Table')
st.caption('This table updates on your selected filters')
st.dataframe(df_select[['SQ NUM','Posting Date','GroupName','Customer/Vendor Name','SQ TOTAL','SO NUM','SO TOTAL']], use_container_width=True)