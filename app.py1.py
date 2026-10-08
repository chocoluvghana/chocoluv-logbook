import streamlit as st
import pandas as pd
import datetime
from sqlalchemy import create_engine
import re

# 1. Page Configuration
st.set_page_config(
    page_title="Chocoluv Ltd Sales Logbook",
    page_icon="🍫",
    layout="wide"
)

# 2. Custom Chocolate Theme Styling
st.markdown("""
    <style>
    .main { background-color: #1a110b; color: #f5efe6; }
    .stTextInput > div > div > input, .stSelectbox > div > div > select, .stNumberInput > div > div > input {
        background-color: #2c1d14; color: #f5efe6; border: 1px solid #5a3d28;
    }
    div[data-testid="stMetric"] {
        background-color: #2c1d14; border: 1px solid #5a3d28; padding: 15px; border-radius: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# 3. Load and Process Product Catalog with Smart Fallbacks
@st.cache_data
def load_catalog():
    try:
        df_cat = pd.read_excel('chocoluv_product_catalog.xlsx', sheet_name='Sheet1')
    except Exception as e:
        st.error(f"Error loading product catalog: {e}")
        return pd.DataFrame()
    
    # Clean weights to grams
    def parse_weight(w):
        if pd.isna(w): return 100.0
        w_str = str(w).strip()
        match = re.search(r'(\d+(?:\.\d+)?)', w_str)
        if match:
            val = float(match.group(1))
            return val * 1000.0 if 'kg' in w_str.lower() else val
        return 100.0

    df_cat['weight_g'] = df_cat['Weight'].apply(parse_weight)
    
    # Fill missing prices with smart defaults or fallbacks
    df_cat['local_price'] = pd.to_numeric(df_cat['Local Price (GHS)'], errors='coerce').fillna(25.0)
    df_cat['export_price'] = pd.to_numeric(df_cat['Export Price (GHS)'], errors='coerce').fillna(50.0)
    
    return df_cat

catalog_df = load_catalog()

# 4. Database Connection Setup
@st.cache_resource
def init_connection():
    try:
        db_url = st.secrets["postgres"]["url"]
    except:
        db_url = "sqlite:///chocoluv_sales.db"
    return create_engine(db_url)

engine = init_connection()

# Ensure table exists
with engine.begin() as conn:
    conn.execute(pd.io.sql.text("""
        CREATE TABLE IF NOT EXISTS chocoluv_sales (
            id SERIAL PRIMARY KEY,
            log_type VARCHAR(50),
            date DATE,
            entered_by VARCHAR(100),
            client VARCHAR(150),
            product VARCHAR(100),
            batch_number VARCHAR(50),
            weight_per_unit DECIMAL(10,2),
            quantity INT,
            unit_cost DECIMAL(10,2),
            exchange_rate DECIMAL(10,4)
        );
    """))

# 5. Sidebar Authentication / Login State
st.sidebar.markdown("### 👤 Staff Portal")
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    username = st.sidebar.text_input("Enter Your Name")
    role = st.sidebar.selectbox("Select Role", ["Sales Officer", "QA Officer"])
    if st.sidebar.button("Log In"):
        if username:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = role
            st.rerun()
        else:
            st.sidebar.warning("Please enter your name.")
    st.stop()
else:
    st.sidebar.success(f"Logged in as: {st.session_state.username} ({st.session_state.role})")
    if st.sidebar.button("Log Out"):
        st.session_state.logged_in = False
        st.rerun()

# 6. Main Header & Navigation Tabs
st.title("🍫 CHOCOLUV LTD SALES LOGBOOK")
st.markdown("Record customer orders, manage batch numbers, exchange rates, and audit financial records.")

tab_entry, tab_local, tab_export, tab_protocol, tab_qa, tab_analytics = st.tabs([
    "📝 New Sales Entry", 
    "📦 Local Sales Table", 
    "🚢 Exports Table", 
    "🎁 Protocol Table", 
    "🔍 QA Desk", 
    "📊 Analytics & Trends"
])

# --- TAB 1: NEW SALES ENTRY ---
with tab_entry:
    st.subheader("Sales Details (Packaging / Entry)")
    with st.form("entry_form", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            log_type = st.selectbox("Target Section", ["Local Sales", "Exports", "Protocol"])
        with col2:
            entry_date = st.date_input("Date", datetime.date.today())
        with col3:
            staff_name = st.text_input("Entered By (Staff Name)", value=st.session_state.username)
            
        col4, col5, col6 = st.columns(3)
        with col4:
            client = st.text_input("Customer Name", placeholder="e.g. Tomreik Hotel")
        with col5:
            product_list = catalog_df['Product Name'].tolist() if not catalog_df.empty else ["Default Bar"]
            selected_product = st.selectbox("Product Name", product_list)
        with col6:
            batch_number = st.text_input("Batch Number", placeholder="e.g. BATCH-2026-001")
            
        # Automatically fetch default weight & price with robust fallbacks
        default_weight = 100.0
        default_cost = 25.0
        if not catalog_df.empty:
            prod_row = catalog_df[catalog_df['Product Name'] == selected_product]
            if not prod_row.empty:
                default_weight = float(prod_row['weight_g'].values[0])
                if log_type == "Exports":
                    default_cost = float(prod_row['export_price'].values[0])
                else:
                    default_cost = float(prod_row['local_price'].values[0])

        col7, col8, col9 = st.columns(3)
        with col7:
            weight_per_unit = st.number_input("Weight per Unit (g)", value=default_weight)
        with col8:
            quantity = st.number_input("Quantity", min_value=1, value=1)
        with col9:
            unit_cost = st.number_input("Unit Cost (GHS)", value=default_cost)
            
        submitted = st.form_submit_button("Save Order Entry")
        if submitted:
            if not client or not batch_number:
                st.error("Please fill in Customer Name and Batch Number.")
            else:
                insert_query = """
                    INSERT INTO chocoluv_sales (log_type, date, entered_by, client, product, batch_number, weight_per_unit, quantity, unit_cost, exchange_rate)
                    VALUES (:l_type, :dt, :eb, :cl, :prod, :batch, :wt, :qty, :cost, :rate)
                """
                with engine.begin() as conn:
                    conn.execute(
                        pd.io.sql.text(insert_query), 
                        {
                            "l_type": log_type, "dt": entry_date, "eb": staff_name,
                            "cl": client, "prod": selected_product, "batch": batch_number,
                            "wt": weight_per_unit, "qty": quantity, "cost": unit_cost, "rate": 11.2000
                        }
                    )
                st.success(f"Successfully recorded order for {client}!")

# Helper function to load and calculate data views
def load_logbook_data(section_name):
    query = f"SELECT * FROM chocoluv_sales WHERE log_type = '{section_name}' ORDER BY date DESC"
    try:
        df = pd.read_sql(query, engine)
    except:
        df = pd.DataFrame()
        
    if not df.empty:
        df['Total Weight (g)'] = df['weight_per_unit'] * df['quantity']
        df['Total Weight (Metric Tonnes)'] = df['Total Weight (g)'] / 1_000_000
        df['Total Amount (GHS)'] = df['unit_cost'] * df['quantity']
        df['Total Amount (USD)'] = df['Total Amount (GHS)'] / df['exchange_rate']
    return df

# --- TABS 2, 3, 4: TABLES (Local, Exports, Protocol) ---
for tab_obj, sec_title in zip([tab_local, tab_export, tab_protocol], ["Local Sales", "Exports", "Protocol"]):
    with tab_obj:
        st.subheader(f"📋 {sec_title} Logbook")
        df_sec = load_logbook_data(sec_title)
        if not df_sec.empty:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Rows / Orders", len(df_sec))
            m2.metric("Total Weight", f"{df_sec['Total Weight (Metric Tonnes)'].sum():.6f} MT")
            m3.metric("Total Money (GHS)", f"GH₵ {df_sec['Total Amount (GHS)'].sum():,.2f}")
            m4.metric("Total Money (USD)", f"${df_sec['Total Amount (USD)'].sum():,.2f}")
            
            st.dataframe(df_sec, use_container_width=True)
            
            csv_data = df_sec.to_csv(index=False).encode('utf-8')
            st.download_button(f"Download {sec_title} CSV Report", csv_data, f"{sec_title.lower().replace(' ', '_')}_report.csv", "text/csv")
        else:
            st.info(f"No records found for {sec_title} yet. Add entries using the 'New Sales Entry' tab.")

# --- TAB 5: QA DESK ---
with tab_qa:
    st.subheader("🔍 QA Desk & Exchange Rate Management")
    st.markdown("Apply monthly exchange rates or edit records to calculate USD valuations accurately.")
    
    qa_target = st.selectbox("Select Target Section to Update Exchange Rate", ["Local Sales", "Exports", "Protocol"], key="qa_target_sec")
    new_rate = st.number_input("New Dollar Exchange Rate (GHS/$)", value=11.2000, format="%.4f")
    
    if st.button("Apply Exchange Rate to Entire Section"):
        update_query = "UPDATE chocoluv_sales SET exchange_rate = :rate WHERE log_type = :l_type"
        with engine.begin() as conn:
            conn.execute(pd.io.sql.text(update_query), {"rate": new_rate, "l_type": qa_target})
        st.success(f"Successfully updated exchange rate to {new_rate} for all {qa_target} records!")

# --- TAB 6: ANALYTICS & TRENDS ---
with tab_analytics:
    st.subheader("📊 Sales & Production Performance Dashboard")
    
    try:
        df_all = pd.read_sql("SELECT * FROM chocoluv_sales", engine)
    except:
        df_all = pd.DataFrame()
        
    if not df_all.empty:
        df_all['Total Weight (g)'] = df_all['weight_per_unit'] * df_all['quantity']
        df_all['Total Weight (Metric Tonnes)'] = df_all['Total Weight (g)'] / 1_000_000
        df_all['Total Amount (GHS)'] = df_all['unit_cost'] * df_all['quantity']
        df_all['Total Amount (USD)'] = df_all['Total Amount (GHS)'] / df_all['exchange_rate']
        
        a1, a2, a3, a4 = st.columns(4)
        a1.metric("Total Company Orders", len(df_all))
        a2.metric("Total Company Weight", f"{df_all['Total Weight (Metric Tonnes)'].sum():.6f} MT")
        a3.metric("Total Revenue (GHS)", f"GH₵ {df_all['Total Amount (GHS)'].sum():,.2f}")
        a4.metric("Total Revenue (USD)", f"${df_all['Total Amount (USD)'].sum():,.2f}")
        
        st.markdown("---")
        st.markdown("### Revenue by Logbook Category")
        cat_grouped = df_all.groupby('log_type')['Total Amount (GHS)'].sum().reset_index()
        st.bar_chart(cat_grouped.set_index('log_type'))
        
        st.markdown("### Top Distributed Products by Weight (MT)")
        prod_grouped = df_all.groupby('product')['Total Weight (Metric Tonnes)'].sum().reset_index()
        st.bar_chart(prod_grouped.set_index('product'))
    else:
        st.info("No data available yet for analytics. Start adding orders in the entry tab!")