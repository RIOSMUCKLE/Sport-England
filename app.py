import streamlit as st
import os
from dotenv import load_dotenv
load_dotenv()
import pandas as pd
import pyodbc
import psycopg2
import sqlalchemy as sa
import matplotlib.pyplot as plt
from datetime import datetime as dt, timedelta, timezone, date
from dateutil.relativedelta import relativedelta
from PIL import Image
from st_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode, DataReturnMode, ColumnsAutoSizeMode
import io
import xlsxwriter

#Page Setting

st.set_page_config(
    page_title="FA Reporting",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': '''
                    ## Muckle FA Reporting Dashboard
                    
                    The Data & Digital Team will be developing this dashboard based on your feedback and requests.
	            '''
    }
    )

# Initialise session state variables
# Store the initial value of widgets in session state
if "start_date" not in st.session_state:
    st.session_state.start_date = dt.now().date().replace(month=4, day=1)

if "end_date" not in st.session_state:
    st.session_state.end_date = dt.today().date()

if "timescale" not in st.session_state:
    st.session_state.timescale = "This Financial Year"

if "days_in_period" not in st.session_state:
    start = st.session_state.start_date
    end = st.session_state.end_date
    delta = end - start
    days = delta.days
    st.session_state.days_in_period = days + 1

if "customRange" not in st.session_state:
    st.session_state.customRange = True

# Initialise session state variables for narrative using free search text bar

#with open('style.css') as f:
#    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

ad_user = os.environ['AD_USER']
ad_pass = os.environ['AD_PASS']
ad_server = os.environ['AD_SERVER']
ad_database = os.environ['AD_DATABASE']

@st.cache_resource
def init_aderant_connection():
    return pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER={%s};DATABASE={%s};UID={%s};PWD={%s}' % (ad_server, ad_database, ad_user, ad_pass) )
    # return sa.create_engine('mssql+pyodbc://{%s}:{%s}@{%s}/{%s}' % (ad_user, ad_pass, ad_server, ad_database) )

conn = init_aderant_connection()

# Perform query.
# Uses st.experimental_memo to only rerun when the query changes or after 10 min.
@st.cache_data(ttl=600)
def run_aderant_query(query):
    return pd.read_sql(query, conn)

sport_england_query = """
        
SELECT 


		t.TRAN_DATE as [Transaction Date],
        c.CLIENT_NAME + ' (' + c.CLIENT_CODE + ')' as [Client_Name],
		m.MATTER_NAME + ' (' + m.MATTER_CODE + ')' as [Matter_Name],
		p.EMPLOYEE_NAME as [Employee Name],
		t.DEPT as [Department],
		t.ACTION_CODE [Action Code],
		sum(t.tobill_hrs) as [ToBill Hours],
		sum(t.tobill_amt) as [ToBill Amount],
		te.TXT1 as [Narrative]

FROM [dbo].[HBM_CLIENT] as c 
	INNER JOIN [dbo].[HBM_MATTER] as m on c.CLIENT_UNO = m.CLIENT_UNO
	inner join [dbo].[HBL_MATT_TYPE] as mt on m.MATT_TYPE_CODE = mt.MATT_TYPE_CODE
	inner join [dbo].[TAT_TIME] as t on m.MATTER_UNO = t.MATTER_UNO
	inner join [dbo].[HBM_PERSNL] as p on t.TK_EMPL_UNO = p.EMPL_UNO 
	inner join [dbo].[TAT_TEXT] as te on t.NAR_TEXT_ID = te.TEXT_ID


where c.CLIENT_CODE = '44260'-- and m.MATTER_CODE = '5'

group by

	
	t.TRAN_DATE,
	c.CLIENT_NAME, 
	m.MATTER_NAME,
	p.EMPLOYEE_NAME,
	t.dept,
	t.ACTION_CODE,
	c.CLIENT_CODE, 
	m.MATTER_CODE, 
	te.TXT1
	

order by t.TRAN_DATE

"""

def changeTimescale():
    if st.session_state.timescale == "Custom Range":
        st.session_state.customRange = False
    else:
        st.session_state.customRange = True

    if st.session_state.timescale == "Today":
        st.session_state.start_date = dt.now().date()
        st.session_state.end_date = dt.now().date()
    
    if st.session_state.timescale == "Yesterday":
        st.session_state.start_date = dt.now() - relativedelta(days=1)
        st.session_state.end_date = dt.now() - relativedelta(days=1)
    
    if st.session_state.timescale == "This Week":
        st.session_state.start_date = dt.now() + relativedelta(days=-dt.now().weekday())
        st.session_state.end_date =  st.session_state.start_date + relativedelta(days=6)
    
    if st.session_state.timescale == "Last Week":
        st.session_state.start_date = dt.now() + relativedelta(days=-dt.now().weekday(), weeks=-1)
        st.session_state.end_date =  st.session_state.start_date + relativedelta(days=6)

    if st.session_state.timescale == "This Month":
        st.session_state.start_date = dt.now().date().replace(day=1)
        st.session_state.end_date = dt.now().date()

    if st.session_state.timescale == "Last Month":
        st.session_state.end_date = dt.now().replace(day=1) - relativedelta(days=1)
        st.session_state.start_date = st.session_state.end_date.replace(day=1)      

    if st.session_state.timescale == "This Year":
        st.session_state.start_date = dt.now().date().replace(month=1, day=1)
        st.session_state.end_date = dt.now().date()

    if st.session_state.timescale == "Last Year":
        st.session_state.start_date = dt.now().date().replace(month=1, day=1)  - relativedelta(years=1)
        st.session_state.end_date = dt.now().date().replace(month=12, day=31)  - relativedelta(years=1)

    if st.session_state.timescale == "This Financial Year":
        st.session_state.start_date = dt.now().date().replace(month=4, day=1)
        st.session_state.end_date = dt.now().date()

    if st.session_state.timescale == "Last Financial Year":
        st.session_state.start_date = dt.now().date().replace(month=4, day=1) - relativedelta(years=1)
        st.session_state.end_date = dt.now().date().replace(month=3, day=31)
    
    if st.session_state.timescale == "This Quarter":
        st.session_state.start_date = pd.to_datetime(pd.datetime.today() - pd.tseries.offsets.QuarterBegin(startingMonth=4)).date()
        st.session_state.end_date = dt.now().date()

    if st.session_state.timescale == "Last Quarter":
        st.session_state.start_date = pd.to_datetime(pd.datetime.today() - pd.tseries.offsets.QuarterBegin(startingMonth=4)).date()  - relativedelta(months=4)
        st.session_state.end_date = pd.to_datetime(pd.datetime.today() + pd.tseries.offsets.QuarterEnd(startingMonth=4)).date() - relativedelta(months=4)

    start = st.session_state.start_date
    end = st.session_state.end_date
    delta = end - start
    days = delta.days
    st.session_state.days_in_period = days + 1

timescales = st.sidebar.selectbox(
    'Select a date range',
    ('Custom Range', 'This Financial Year', 'Last Financial Year', 'This Quarter', 'Last Quarter', 'This Month', 'Last Month', 'This Week', 'Last Week', 'Today', 'Yesterday', 'This Year', 'Last Year'),
    help="Choose a timescale to update the start and end dates. This will filter all values on the table",
    key="timescale",
    on_change=changeTimescale)
 
startDate = st.sidebar.date_input(label='Start Date: ',
            help="Start Date",
            key="start_date",
            disabled=st.session_state.customRange)

endDate = st.sidebar.date_input(label='End Date: ',
            help="End Date",
            key="end_date",
            disabled=st.session_state.customRange)

sport_england_table = run_aderant_query(sport_england_query)
# sport_england_table['Copy - Transaction Date'] = sport_england_table['Transaction Date']
sport_england_table.set_index('Transaction Date', inplace=True)
sport_england_table = sport_england_table.sort_index().loc[st.session_state.start_date:st.session_state.end_date]
sport_england_table.reset_index(inplace=True)

#sport_england_table_total = sport_england_table.copy()

#sport_england_table_total.loc['Total']= sport_england_table.sum(numeric_only=True)

sport_england_table['ToBill Hours'].fillna(0.0, inplace=True)
sport_england_table['ToBill Amount'].fillna(0.0, inplace=True)
sport_england_table['Narrative'].fillna("", inplace=True)

c = st.container()

with c:
    st.image(Image.open('muckle_logo.png'), use_column_width=False, width=200)

    select_all_matters = st.sidebar.checkbox("Select All Matters")

    #sport_england_table["ToBill Amount"] = sport_england_table["ToBill Amount"].apply(lambda x: "{:.2f}".format((float(x))))

    if select_all_matters:
        matter_filter = st.sidebar.multiselect("Select a Matter", sport_england_table.sort_values(by="Matter_Name").Matter_Name.unique(), sport_england_table.sort_values(by="Matter_Name").Matter_Name.unique())
    else:
        matter_filter = st.sidebar.multiselect("Select a Matter", sport_england_table.sort_values(by="Matter_Name").Matter_Name.unique())    
    
    sport_england_table = sport_england_table[sport_england_table["Matter_Name"].isin(matter_filter)]

    query = st.sidebar.text_input("Filter Narrative")

    if query:
        mask = sport_england_table.map(lambda x: query.lower() in str(x).lower()).any(axis=1)
        sport_england_table=sport_england_table[mask]

    gb = GridOptionsBuilder.from_dataframe(sport_england_table)
    gb.configure_pagination(paginationAutoPageSize=True)
    gb.configure_default_column(editable=True, groupable=True)
    gb.configure_selection(selection_mode="multiple", use_checkbox=True, header_checkbox=True)

    gb.configure_column("ToBill Amount", 
                        type=["numericColumn","numberColumnFilter","customNumericFormat"], 
                        valueFormatter="data['ToBill Amount'].toLocaleString('en-US', {style: 'currency', currency: 'GBP', maximumFractionDigits:2})")
    
    gridOptions = gb.build()
    grid_response = AgGrid(sport_england_table, gridOptions=gridOptions, columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS, enable_enterprise_modules=False)

    columns = st.sidebar.multiselect("Columns:", sport_england_table.columns, default=list(sport_england_table.columns))

    download = sport_england_table[columns]
    sport_england_table.loc['Total'] = sport_england_table.sum(numeric_only=True)
    #download = download.drop('Copy - Transaction Date', axis=1)
    download.loc['Total'] = download.sum(numeric_only=True)

    if not sport_england_table.empty:
        download.loc[:, "ToBill Amount"] = "£" + download["ToBill Amount"].map('{:,.2f}'.format)

    def to_excel(download) -> bytes:
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine="xlsxwriter")
        download.to_excel(writer, sheet_name="Sheet1")
        writer.close()
        processed_data = output.getvalue()
        return processed_data


    st.download_button(
        "Download Excel File",
        data=to_excel(download),
        file_name="FA Report .xlsx",
        mime="application/vnd.ms-excel",
    )

    #total after filters