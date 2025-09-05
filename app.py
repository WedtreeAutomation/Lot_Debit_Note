import streamlit as st
import pandas as pd
import xmlrpc.client
import os
from dotenv import load_dotenv
import time
import re
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

# Load environment variables
load_dotenv()

# Page configuration
st.set_page_config(
    page_title="Odoo PO Scanner Pro",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Enhanced Custom CSS for professional styling
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    .main-header {
        font-size: 3.5rem;
        font-weight: 700;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        text-align: center;
        margin-bottom: 2rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
    
    .hero-section {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 2rem;
        border-radius: 20px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        text-align: center;
    }
    
    .hero-subtitle {
        font-size: 1.2rem;
        color: #6c757d;
        margin-top: 1rem;
        font-weight: 400;
    }
    
    .stats-container {
        display: flex;
        gap: 1rem;
        margin: 2rem 0;
    }
    
    .stat-card {
        background: white;
        padding: 1.5rem;
        border-radius: 15px;
        box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        text-align: center;
        flex: 1;
        border-left: 4px solid #667eea;
        transition: transform 0.3s ease;
    }
    
    .stat-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    }
    
    .stat-number {
        font-size: 2rem;
        font-weight: 700;
        color: #667eea;
        margin: 0;
    }
    
    .stat-label {
        font-size: 0.9rem;
        color: #6c757d;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-top: 0.5rem;
    }
    
    .login-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 25px;
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.3);
        color: white;
    }
    
    .success-box {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        color: white;
        padding: 20px;
        border-radius: 15px;
        margin: 15px 0;
        box-shadow: 0 5px 15px rgba(79, 172, 254, 0.3);
        border: none;
    }
    
    .error-box {
        background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%);
        color: #721c24;
        padding: 20px;
        border-radius: 15px;
        margin: 15px 0;
        box-shadow: 0 5px 15px rgba(255, 154, 158, 0.3);
        border: none;
    }
    
    .info-box {
        background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
        color: #0c5460;
        padding: 20px;
        border-radius: 15px;
        margin: 15px 0;
        box-shadow: 0 5px 15px rgba(168, 237, 234, 0.3);
        border: none;
    }
    
    .warning-box {
        background: linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%);
        color: #856404;
        padding: 20px;
        border-radius: 15px;
        margin: 15px 0;
        box-shadow: 0 5px 15px rgba(252, 182, 159, 0.3);
        border: none;
    }
    
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 12px 30px;
        border-radius: 25px;
        cursor: pointer;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
        transition: all 0.3s ease;
        box-shadow: 0 5px 15px rgba(102, 126, 234, 0.3);
    }
    
    .stButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(102, 126, 234, 0.4);
    }
    
    .reset-button {
        background: linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%);
    }
    
    .upload-section {
        background: white;
        padding: 2rem;
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        margin-bottom: 2rem;
        border: 2px dashed #667eea;
        transition: all 0.3s ease;
    }
    
    .upload-section:hover {
        border-color: #764ba2;
        transform: translateY(-2px);
    }
    
    .results-section {
        background: white;
        padding: 2rem;
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        margin-bottom: 2rem;
    }
    
    .sidebar-section {
        background: rgba(255,255,255,0.1);
        padding: 1rem;
        border-radius: 15px;
        margin-bottom: 1rem;
        backdrop-filter: blur(10px);
    }
    
    .metric-card {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 1rem;
        border-radius: 15px;
        color: white;
        margin: 0.5rem 0;
        text-align: center;
    }
    
    .process-indicator {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px;
        margin: 1rem 0;
    }
    
    .step-indicator {
        width: 30px;
        height: 30px;
        border-radius: 50%;
        background: #e9ecef;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: bold;
        color: #6c757d;
    }
    
    .step-indicator.active {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    
    .step-indicator.completed {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        color: white;
    }
    
    /* Hide Streamlit branding */
    .stApp > header[data-testid="stHeader"] {
        background: transparent;
    }
    
    .stApp > div[data-testid="stToolbar"] {
        display: none;
    }
    
    footer {
        display: none;
    }
    
    /* Custom scrollbar */
    ::-webkit-scrollbar {
        width: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: #f1f1f1;
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, #764ba2 0%, #667eea 100%);
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'odoo_connected' not in st.session_state:
    st.session_state.odoo_connected = False
if 'processing' not in st.session_state:
    st.session_state.processing = False
if 'results' not in st.session_state:
    st.session_state.results = None
if 'uploaded_data' not in st.session_state:
    st.session_state.uploaded_data = None
if 'barcode_col_name' not in st.session_state:
    st.session_state.barcode_col_name = None
if 'po_col_name' not in st.session_state:
    st.session_state.po_col_name = None
if 'lot_processing_done' not in st.session_state:
    st.session_state.lot_processing_done = False
if 'lot_results' not in st.session_state:
    st.session_state.lot_results = []

# Reset function
def reset_application():
    """Reset all session state variables"""
    st.session_state.processing = False
    st.session_state.results = None
    st.session_state.uploaded_data = None
    st.session_state.barcode_col_name = None
    st.session_state.po_col_name = None
    st.session_state.lot_processing_done = False
    st.session_state.lot_results = []

# Odoo connection function
def connect_to_odoo():
    try:
        url = os.getenv("ODOO_URL")
        db = os.getenv("ODOO_DB")
        username = os.getenv("ODOO_USERNAME")
        password = os.getenv("ODOO_PASSWORD")
        
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
        uid = common.authenticate(db, username, password, {})
        
        if uid:
            st.session_state.odoo_uid = uid
            st.session_state.odoo_models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")
            st.session_state.odoo_db = db
            st.session_state.odoo_password = password
            st.session_state.odoo_connected = True
            return True, "Connection successful!"
        else:
            return False, "Authentication failed."
            
    except Exception as e:
        return False, f"Connection error: {str(e)}"

# Function to process a single barcode and PO
def process_product(barcode, po_number):
    try:
        # Find product by barcode
        product_ids = st.session_state.odoo_models.execute_kw(
            st.session_state.odoo_db, 
            st.session_state.odoo_uid, 
            st.session_state.odoo_password,
            'product.product', 'search',
            [[['barcode', '=', barcode]]]
        )

        if not product_ids:
            return {
                'status': 'error',
                'barcode': barcode,
                'po_number': po_number,
                'message': f"No product found with barcode {barcode}"
            }

        product_id = product_ids[0]
        product_data = st.session_state.odoo_models.execute_kw(
            st.session_state.odoo_db, 
            st.session_state.odoo_uid, 
            st.session_state.odoo_password,
            'product.product', 'read',
            [product_ids], {'fields': ['id', 'name', 'default_code', 'barcode']}
        )[0]

        # Find the Purchase Order by number
        po_ids = st.session_state.odoo_models.execute_kw(
            st.session_state.odoo_db, 
            st.session_state.odoo_uid, 
            st.session_state.odoo_password,
            'purchase.order', 'search',
            [[['name', '=', po_number]]]
        )

        if not po_ids:
            return {
                'status': 'error',
                'barcode': barcode,
                'po_number': po_number,
                'message': f"No Purchase Order found with number {po_number}"
            }

        po_id = po_ids[0]

        # Read PO Header
        po_data = st.session_state.odoo_models.execute_kw(
            st.session_state.odoo_db, 
            st.session_state.odoo_uid, 
            st.session_state.odoo_password,
            'purchase.order', 'read',
            [[po_id]], {'fields': ['name', 'partner_id', 'partner_ref', 'date_order', 'amount_total', 'state']}
        )[0]

        vendor_name = po_data['partner_id'][1] if po_data['partner_id'] else "Unknown Vendor"
        vendor_ref = po_data.get('partner_ref', '')

        # Read PO Lines for the given Product
        po_line_ids = st.session_state.odoo_models.execute_kw(
            st.session_state.odoo_db, 
            st.session_state.odoo_uid, 
            st.session_state.odoo_password,
            'purchase.order.line', 'search',
            [[['order_id', '=', po_id], ['product_id', '=', product_id]]]
        )

        if not po_line_ids:
            return {
                'status': 'error',
                'barcode': barcode,
                'po_number': po_number,
                'message': f"Product {product_data['name']} not found in PO {po_number}"
            }

        po_lines = st.session_state.odoo_models.execute_kw(
            st.session_state.odoo_db, 
            st.session_state.odoo_uid, 
            st.session_state.odoo_password,
            'purchase.order.line', 'read',
            [po_line_ids], {'fields': [
                'product_id', 'name', 'product_qty', 'price_unit',
                'price_subtotal', 'date_planned'
            ]}
        )

        # Prepare result
        result = {
            'status': 'success',
            'barcode': barcode,
            'po_number': po_number,
            'product_id': product_data['id'],
            'product_name': product_data['name'],
            'product_ref': product_data.get('default_code', ''),
            'po_name': po_data['name'],
            'vendor_name': vendor_name,
            'vendor_ref': vendor_ref,
            'order_date': po_data['date_order'],
            'total_amount': po_data['amount_total'],
            'po_status': po_data['state'],
            'quantity': po_lines[0]['product_qty'],
            'unit_price': po_lines[0]['price_unit'],
            'subtotal': po_lines[0]['price_subtotal'],
            'scheduled_date': po_lines[0]['date_planned'],
            'sku': product_data.get('default_code', '')
        }

        return result

    except Exception as e:
        return {
            'status': 'error',
            'barcode': barcode,
            'po_number': po_number,
            'message': f"Processing error: {str(e)}"
        }

# Function to find the correct column names in the uploaded file
def find_columns(df):
    barcode_col = None
    po_col = None
    
    # Convert column names to lowercase for easier matching
    lower_columns = [col.lower() for col in df.columns]
    
    # Look for barcode column
    barcode_patterns = ['barcode', 'bar.code', 'bar_code']
    for pattern in barcode_patterns:
        for i, col in enumerate(lower_columns):
            if pattern in col:
                barcode_col = df.columns[i]
                break
        if barcode_col:
            break
    
    # Look for purchase order column
    po_patterns = ['po', 'purchase.order', 'purchase_order', 'ponumber', 'po_number']
    for pattern in po_patterns:
        for i, col in enumerate(lower_columns):
            if pattern in col:
                po_col = df.columns[i]
                break
        if po_col:
            break
    
    return barcode_col, po_col

def count_barcode_po_occurrences(df, barcode_col, po_col):
    # Group by both barcode and PO number to get the count for each combination
    df_standardized = df.copy()
    df_standardized['barcode'] = df_standardized[barcode_col].astype(str)
    df_standardized['po_number'] = df_standardized[po_col].astype(str)
    
    barcode_po_counts = df_standardized.groupby(['barcode', 'po_number']).size().reset_index()
    barcode_po_counts.columns = ['barcode', 'po_number', 'Ref_Quantity']
    return barcode_po_counts

# Function to count overall barcode occurrences (for Out_Quantity)
def count_barcode_occurrences(df, barcode_col):
    # Create a copy of the barcode column with standardized name for merging
    df_standardized = df.copy()
    df_standardized['barcode'] = df_standardized[barcode_col].astype(str)
    
    barcode_counts = df_standardized['barcode'].value_counts().reset_index()
    barcode_counts.columns = ['barcode', 'Out_Quantity']
    return barcode_counts

def get_odoo_connection(url, db, username, password):
    try:
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
        uid = common.authenticate(db, username, password, {})
        if not uid:
            raise Exception("Authentication failed for user")
        models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")
        return uid, models
    except Exception as e:
        raise Exception(f"Odoo connection failed: {e}")

# Function to update and apply lot to zero
def update_and_apply_lot_to_zero(lot_name, expected_qty):
    try:
        CONFIG = {
            'url': os.getenv("ODOO_URL_PROCESS_LOTS"),
            'db': os.getenv("ODOO_DB_PROCESS_LOTS"),
            'username': os.getenv("ODOO_USERNAME_PROCESS_LOTS"),
            'password': os.getenv("ODOO_PASSWORD_PROCESS_LOTS"),
            'damage_location_name': "Damge/Stock",
            'company_name': "Wedtree eStore Private Limited - HO"
        }

        BASE = CONFIG['url'].rstrip('/')

        # ✅ Authenticate separately for PROCESS_LOTS user
        uid, models = get_odoo_connection(
            BASE, CONFIG['db'], CONFIG['username'], CONFIG['password']
        )

        # 1) fetch all quants for this lot
        quants = models.execute_kw(
            CONFIG['db'], uid, CONFIG['password'],
            'stock.quant', 'search_read',
            [[('lot_id.name', '=', lot_name)]],
            {'fields': ['id', 'location_id', 'quantity', 'inventory_quantity', 'lot_id', 'company_id']}
        )

        if not quants:
            return f"❌ Lot '{lot_name}' not found in stock.quant."

        # 2) filter only for the right company
        company_name = CONFIG['company_name']
        quants = [q for q in quants if q['company_id'] and q['company_id'][1] == company_name]

        if not quants:
            return f"🚨 Lot '{lot_name}' exists but not under company '{company_name}'. Skipping."

        # 3) keep only quants in Damage/Stock by display name
        loc_name = CONFIG['damage_location_name']
        damage_quants = [q for q in quants if q['location_id'] and q['location_id'][1] == loc_name]

        if not damage_quants:
            locations = sorted({q['location_id'][1] for q in quants if q['location_id']})
            where = ", ".join(locations) if locations else "Unknown"
            return f"🚨 Lot '{lot_name}' exists under '{company_name}' but NOT in '{loc_name}'. Current locations: {where}"

        # 4) compare on-hand quantity with expected
        onhand = sum(q['quantity'] for q in damage_quants)
        if abs(onhand - float(expected_qty)) > 1e-6:
            return f"⚠️ Qty mismatch for lot '{lot_name}' in '{loc_name}' (company={company_name}): On-hand={onhand}, Entered={expected_qty}"

        # 5) set counted quantity to 0 for those quants
        quant_ids = [q['id'] for q in damage_quants]
        for qid in quant_ids:
            models.execute_kw(
                CONFIG['db'], uid, CONFIG['password'],
                'stock.quant', 'write',
                [[qid], {'inventory_quantity': 0.0}]
            )

        # 6) APPLY (same as clicking Apply All in the UI)
        try:
            models.execute_kw(
                CONFIG['db'], uid, CONFIG['password'],
                'stock.quant', 'action_apply_inventory',
                [quant_ids]
            )
            return f"✅ Applied for company '{company_name}': lot '{lot_name}' in '{loc_name}' set to 0. Quants={quant_ids}, On-hand={onhand}"
        except xmlrpc.client.Fault:
            return f"✅ Odoo Successfully applied lot '{lot_name}'"
        except Exception as e:
            return f"⚠️ Unexpected error for lot '{lot_name}': {e}"

    except Exception as e:
        return f"❌ Error processing lot '{lot_name}': {str(e)}"

# Enhanced sidebar with modern design
with st.sidebar:
    st.markdown("<div style='text-align: center; margin-bottom: 2rem;'>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Login Section
    st.markdown("<div class='sidebar-section'>", unsafe_allow_html=True)
    st.markdown("### 🔐 Authentication")
    
    if not st.session_state.logged_in:
        with st.container():
            username = st.text_input("👤 Username", placeholder="Enter your username")
            password = st.text_input("🔑 Password", type="password", placeholder="Enter your password")
            
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("🚀 Login", use_container_width=True):
                    if username == os.getenv("APP_USERNAME") and password == os.getenv("APP_PASSWORD"):
                        st.session_state.logged_in = True
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("❌ Invalid credentials")
    else:
        st.success(f"✅ Welcome, {os.getenv('APP_USERNAME')}!")
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.odoo_connected = False
            reset_application()
            st.rerun()
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Odoo Connection Section
    if st.session_state.logged_in:
        st.markdown("<div class='sidebar-section'>", unsafe_allow_html=True)
        st.markdown("### 🔗 Odoo Connection")
        
        if not st.session_state.odoo_connected:
            if st.button("🌐 Connect to Odoo", use_container_width=True):
                with st.spinner("🔄 Establishing connection..."):
                    success, message = connect_to_odoo()
                    if success:
                        st.success("🎉 " + message)
                    else:
                        st.error("⚠️ " + message)
        else:
            st.success("🟢 Connected to Odoo")
            st.markdown("<div class='metric-card'>", unsafe_allow_html=True)
            st.markdown("**Status:** Online & Ready")
            st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("---")
        st.header("🔗 Related Applications")
        st.markdown("""
        - [Inventory Debit Note](https://inventory-debit-note.streamlit.app/)
        - [Lot Debit Note](https://lot-debit-note.streamlit.app/)
        - [Lot Credit Note](https://lot-credit-note.streamlit.app/)
        """)
        
        # Process Indicator
        if st.session_state.logged_in and st.session_state.odoo_connected:
            st.markdown("### 📊 Process Status")
            
            steps = [
                ("Login", st.session_state.logged_in),
                ("Connect", st.session_state.odoo_connected),
                ("Upload", st.session_state.uploaded_data is not None),
                ("Process", st.session_state.results is not None),
                ("Lots", st.session_state.lot_processing_done)
            ]
            
            for i, (step_name, completed) in enumerate(steps):
                if completed:
                    st.markdown(f"✅ **{step_name}** - Complete")
                elif i == len([s for s in steps[:i+1] if s[1]]):
                    st.markdown(f"🔄 **{step_name}** - In Progress")
                else:
                    st.markdown(f"⏳ **{step_name}** - Pending")
        
        # Statistics Section
        if st.session_state.results:
            st.markdown("### 📈 Statistics")
            success_count = len([r for r in st.session_state.results if r['status'] == 'success'])
            error_count = len([r for r in st.session_state.results if r['status'] == 'error'])
            
            st.markdown(f"""
            <div class='metric-card'>
                <div style='font-size: 1.5rem; font-weight: bold;'>{success_count}</div>
                <div>Successful</div>
            </div>
            """, unsafe_allow_html=True)
            
            if error_count > 0:
                st.markdown(f"""
                <div class='metric-card' style='background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%);'>
                    <div style='font-size: 1.5rem; font-weight: bold;'>{error_count}</div>
                    <div>Errors</div>
                </div>
                """, unsafe_allow_html=True)
        
        # Reset Button
        st.markdown("---")
        st.markdown("### 🔄 Reset Application")
        if st.button("🗑️ Reset All Data", use_container_width=True, help="Clear all processed data and start fresh"):
            reset_application()
            st.success("✨ Application reset successfully!")
            time.sleep(1)
            st.rerun()

# Main Content Area
# Hero Section
st.markdown("""
<div class='hero-section'>
    <h1 class='main-header'>Odoo Purchase Order Scanner Pro</h1>
    <p class='hero-subtitle'>🚀 Streamline your purchase order processing with advanced analytics and automation</p>
</div>
""", unsafe_allow_html=True)

if not st.session_state.logged_in:
    st.markdown("""
    <div class='info-box'>
        <h3>👋 Welcome to Odoo PO Scanner Pro!</h3>
        <p>Please authenticate using the sidebar to access the powerful features of our purchase order processing system.</p>
        <ul>
            <li>🔍 Advanced barcode scanning</li>
            <li>📊 Real-time analytics</li>
            <li>🔄 Automated lot processing</li>
            <li>📈 Comprehensive reporting</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
elif not st.session_state.odoo_connected:
    st.markdown("""
    <div class='warning-box'>
        <h3>🔗 Connect to Odoo</h3>
        <p>Please establish a connection to your Odoo instance using the sidebar to begin processing purchase orders.</p>
    </div>
    """, unsafe_allow_html=True)
    
else:
    # File Upload Section
    st.markdown("## 📤 Upload Excel File")
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file", 
        type=['xlsx', 'xls'],
        help="Upload an Excel file containing barcode and purchase order data"
    )
    st.markdown("</div>", unsafe_allow_html=True)
    
    if uploaded_file is not None:
        try:
            # Read Excel file with dtype specification to preserve leading zeros
            df = pd.read_excel(uploaded_file, dtype=str)
            st.session_state.uploaded_data = df
            
            # Find the correct columns
            barcode_col, po_col = find_columns(df)
            st.session_state.barcode_col_name = barcode_col
            st.session_state.po_col_name = po_col
            
            if barcode_col and po_col:
                st.markdown("""
                    <div class='success-box'>
                        <h4>✅ File Successfully Uploaded!</h4>
                        <p><strong>Detected Columns:</strong> '{barcode_col}' and '{po_col}'</p>
                        <p><strong>Total Rows:</strong> {row_count:,}</p>
                    </div>
                    """.format(barcode_col=barcode_col, po_col=po_col, row_count=len(df)), unsafe_allow_html=True)
                
                # Display file preview
                with st.expander("📋 File Preview", expanded=True):
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Records", len(df))
                    with col2:
                        st.metric("Unique Barcodes", df[barcode_col].nunique())
                    with col3:
                        st.metric("Unique POs", df[po_col].nunique())
                    
                    st.markdown("**Sample Data:**")
                    st.dataframe(df.head(), use_container_width=True)
                
                # Count barcode occurrences
                out_quantity_counts = count_barcode_occurrences(df, barcode_col)
                ref_quantity_counts = count_barcode_po_occurrences(df, barcode_col, po_col)
                
                # Processing Button
                st.markdown("<div style='text-align: center; margin: 2rem 0;'>", unsafe_allow_html=True)
                if st.button("🚀 Process Data", type="primary", use_container_width=False):
                    st.session_state.processing = True
                    
                    # Create processing container
                    processing_container = st.container()
                    with processing_container:
                        st.markdown("""
                        <div class='info-box'>
                            <h4>🔄 Processing Your Data</h4>
                            <p>Please wait while we process your purchase orders...</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        results = []
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        # Create metrics containers for real-time updates
                        col1, col2, col3 = st.columns(3)
                        processed_metric = col1.empty()
                        success_metric = col2.empty()
                        error_metric = col3.empty()
                        
                        success_count = 0
                        error_count = 0
                        
                        for i, row in df.iterrows():
                            # Ensure barcode and PO are treated as strings to preserve leading zeros
                            barcode = str(row[barcode_col]).strip()
                            po_number = str(row[po_col]).strip()
                            
                            status_text.text(f"🔍 Processing: {barcode} | PO: {po_number}")
                            result = process_product(barcode, po_number)
                            results.append(result)
                            
                            # Update counters
                            if result['status'] == 'success':
                                success_count += 1
                            else:
                                error_count += 1
                            
                            # Update metrics
                            processed_metric.metric("Processed", f"{i+1}/{len(df)}")
                            success_metric.metric("✅ Success", success_count)
                            error_metric.metric("❌ Errors", error_count)
                            
                            progress_bar.progress((i + 1) / len(df))
                        
                        st.session_state.results = results
                        st.session_state.processing = False
                        
                        # Success message with confetti
                        st.markdown("""
                        <div class='success-box'>
                            <h4>🎉 Processing Complete!</h4>
                            <p>Successfully processed {total} records with {success} successful matches and {errors} errors.</p>
                        </div>
                        """.format(total=len(df), success=success_count, errors=error_count), unsafe_allow_html=True)
                
                st.markdown("</div>", unsafe_allow_html=True)
                
            else:
                st.markdown("""
                <div class='error-box'>
                    <h4>❌ Column Detection Failed</h4>
                    <p>Could not find required columns in the Excel file.</p>
                    <p><strong>Looking for:</strong> columns containing 'barcode' and 'po' (case insensitive)</p>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("**Available columns in your file:**")
                st.code(", ".join(list(df.columns)))
                
        except Exception as e:
            st.markdown(f"""
            <div class='error-box'>
                <h4>❌ File Reading Error</h4>
                <p>Error reading file: {str(e)}</p>
            </div>
            """, unsafe_allow_html=True)
    
    # Display Results Section
    if st.session_state.results:
        st.markdown("<div class='results-section'>", unsafe_allow_html=True)
        
        # Create results dataframe
        results_df = pd.DataFrame(st.session_state.results)
        
        # Separate success and error results
        success_df = results_df[results_df['status'] == 'success'].copy()
        error_df = results_df[results_df['status'] == 'error'].copy()
        
        # Statistics Dashboard
        st.markdown("## 📊 Processing Results Dashboard")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown("""
            <div class='stat-card'>
                <div class='stat-number'>{}</div>
                <div class='stat-label'>Total Processed</div>
            </div>
            """.format(len(results_df)), unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class='stat-card'>
                <div class='stat-number' style='color: #28a745;'>{}</div>
                <div class='stat-label'>Successful</div>
            </div>
            """.format(len(success_df)), unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
            <div class='stat-card'>
                <div class='stat-number' style='color: #dc3545;'>{}</div>
                <div class='stat-label'>Errors</div>
            </div>
            """.format(len(error_df)), unsafe_allow_html=True)
        
        with col4:
            success_rate = (len(success_df) / len(results_df) * 100) if len(results_df) > 0 else 0
            st.markdown("""
            <div class='stat-card'>
                <div class='stat-number' style='color: #17a2b8;'>{:.1f}%</div>
                <div class='stat-label'>Success Rate</div>
            </div>
            """.format(success_rate), unsafe_allow_html=True)
        
        # Visualization Section
        if len(success_df) > 0:
            st.markdown("### 📈 Analytics & Insights")
            
            # Create tabs for different visualizations
            tab1, tab2 = st.tabs(["📊 Vendor Analysis", "💰 Financial Overview"])
            
            with tab1:
                if not success_df.empty:
                    vendor_data = success_df['vendor_name'].value_counts().head(10)
                    fig = px.bar(
                        x=vendor_data.values, 
                        y=vendor_data.index, 
                        orientation='h',
                        title="Top 10 Vendors by Order Count",
                        labels={'x': 'Number of Orders', 'y': 'Vendor'},
                        color=vendor_data.values,
                        color_continuous_scale="Viridis"
                    )
                    fig.update_layout(height=400, showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
            
            with tab2:
                if not success_df.empty:
                    # Convert to numeric for calculation
                    success_df['total_amount_num'] = pd.to_numeric(success_df['total_amount'], errors='coerce')
                    financial_data = success_df.groupby('vendor_name')['total_amount_num'].sum().sort_values(ascending=False).head(10)
                    
                    fig = px.pie(
                        values=financial_data.values,
                        names=financial_data.index,
                        title="Financial Distribution by Top 10 Vendors",
                        color_discrete_sequence=px.colors.qualitative.Set3
                    )
                    fig.update_traces(textposition='inside', textinfo='percent+label')
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Financial metrics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Value", f"{success_df['total_amount_num'].sum():,.2f}")
                    with col2:
                        st.metric("Average Order Value", f"{success_df['total_amount_num'].mean():,.2f}")
                    with col3:
                        st.metric("Largest Order", f"{success_df['total_amount_num'].max():,.2f}")
                    
        # Success Results Table
        if not success_df.empty:
            st.markdown("### ✅ Successful Processing Results")
            
            # Add both Out_Quantity and Ref_Quantity to success results
            if st.session_state.uploaded_data is not None and st.session_state.barcode_col_name and st.session_state.po_col_name:
                # Ensure barcode and po_number are strings in success_df for merging
                success_df['barcode'] = success_df['barcode'].astype(str)
                success_df['po_number'] = success_df['po_number'].astype(str)
                
                # Get Out_Quantity (overall barcode count)
                out_quantity_counts = count_barcode_occurrences(st.session_state.uploaded_data, st.session_state.barcode_col_name)
                
                # Get Ref_Quantity (barcode count with same PO number)
                ref_quantity_counts = count_barcode_po_occurrences(
                    st.session_state.uploaded_data, 
                    st.session_state.barcode_col_name, 
                    st.session_state.po_col_name
                )
                
                # Convert to string for merging
                out_quantity_counts['barcode'] = out_quantity_counts['barcode'].astype(str)
                ref_quantity_counts['barcode'] = ref_quantity_counts['barcode'].astype(str)
                ref_quantity_counts['po_number'] = ref_quantity_counts['po_number'].astype(str)
                
                # Merge both quantities
                success_df = success_df.merge(out_quantity_counts, on='barcode', how='left')
                success_df = success_df.merge(ref_quantity_counts, on=['barcode', 'po_number'], how='left')
            
            # Prepare download data with both quantities
            download_data = success_df[[
                'barcode', 'Out_Quantity', 'Ref_Quantity', 'po_number', 'product_name', 'product_ref', 
                'po_name', 'vendor_name', 'vendor_ref', 'order_date',
                'total_amount', 'po_status', 'quantity', 'unit_price',
                'subtotal', 'scheduled_date', 'sku'
            ]].rename(columns={
                'quantity': 'PO_Quantity',
                'unit_price': 'Unit_Price',
                'subtotal': 'Subtotal',
                'scheduled_date': 'Scheduled_Date'
            })
            
            # Display with enhanced styling
            with st.expander("📋 Detailed Results Table", expanded=False):
                st.dataframe(
                    download_data,
                    use_container_width=True,
                    hide_index=True
                )
            
            # Download Buttons
            st.markdown("### 📥 Download Options")
            col1, col2 = st.columns(2)
            
            with col1:
                # Download button for Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    download_data.to_excel(writer, index=False, sheet_name='PO_Results')
                    # Add barcode counts as additional sheets
                    if st.session_state.uploaded_data is not None and st.session_state.barcode_col_name and st.session_state.po_col_name:
                        out_quantity_counts = count_barcode_occurrences(st.session_state.uploaded_data, st.session_state.barcode_col_name)
                        ref_quantity_counts = count_barcode_po_occurrences(st.session_state.uploaded_data, st.session_state.barcode_col_name, st.session_state.po_col_name)
                        
                        out_quantity_counts.to_excel(writer, index=False, sheet_name='Out_Quantity_Counts')
                        ref_quantity_counts.to_excel(writer, index=False, sheet_name='Ref_Quantity_Counts')
                
                output.seek(0)
                
                st.download_button(
                    label="📊 Download Results (Excel)",
                    data=output,
                    file_name=f"odoo_po_results_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                # Download CSV option
                csv_data = download_data.to_csv(index=False)
                st.download_button(
                    label="📄 Download Results (CSV)",
                    data=csv_data,
                    file_name=f"odoo_po_results_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            # Lot Processing Section
            st.markdown("---")
            st.markdown("### 🔧 Advanced Lot/Serial Processing")
            
            st.markdown("""
            <div class='info-box'>
                <h4>🎯 Lot Processing Information</h4>
                <p>This section will process each unique barcode with its reference quantity and apply lot adjustments to zero in your Odoo inventory system.</p>
            </div>
            """, unsafe_allow_html=True)
            
            if not st.session_state.lot_processing_done:
                if st.button("🔧 Process Lots/Serials to Zero", type="primary", use_container_width=True):
                    st.session_state.lot_processing_done = True
                    
                    with st.spinner("🔄 Processing lots and serials..."):
                        lot_results = []
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        # Get unique barcodes with their Ref_Quantity (same barcode + PO)
                        unique_barcodes = success_df[['barcode', 'Ref_Quantity']].drop_duplicates()
                        
                        for i, (_, row) in enumerate(unique_barcodes.iterrows()):
                            status_text.text(f"🏷️ Processing lot {i+1} of {len(unique_barcodes)}: {row['barcode']}")
                            result = update_and_apply_lot_to_zero(row['barcode'], row['Ref_Quantity'])
                            lot_results.append({
                                'barcode': row['barcode'],
                                'quantity': row['Ref_Quantity'],
                                'result': result
                            })
                            progress_bar.progress((i + 1) / len(unique_barcodes))
                        
                        st.session_state.lot_results = lot_results
                    
                    st.markdown("""
                    <div class='success-box'>
                        <h4>🎉 Lot Processing Complete!</h4>
                        <p>All lots have been successfully processed and applied to your Odoo system.</p>
                    </div>
                    """, unsafe_allow_html=True)
            
            # Display lot processing results
            if st.session_state.lot_results:
                st.markdown("### 📋 Lot Processing Results")
                
                lot_results_df = pd.DataFrame(st.session_state.lot_results)
            
                
                # Display results table
                with st.expander("📊 Detailed Lot Processing Results", expanded=True):
                    st.dataframe(
                        lot_results_df,
                        use_container_width=True,
                        hide_index=True
                    )
                
                # Download lot results
                col1, col2 = st.columns(2)
                with col1:
                    lot_output = BytesIO()
                    with pd.ExcelWriter(lot_output, engine='openpyxl') as writer:
                        lot_results_df.to_excel(writer, index=False, sheet_name='Lot_Results')
                    
                    lot_output.seek(0)
                    
                    st.download_button(
                        label="📊 Download Lot Results (Excel)",
                        data=lot_output,
                        file_name=f"lot_processing_results_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col2:
                    lot_csv = lot_results_df.to_csv(index=False)
                    st.download_button(
                        label="📄 Download Lot Results (CSV)",
                        data=lot_csv,
                        file_name=f"lot_processing_results_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
        
        # Display errors
        if not error_df.empty:
            st.markdown("### ❌ Processing Errors")
            
            st.markdown("""
            <div class='error-box'>
                <h4>⚠️ Items Requiring Attention</h4>
                <p>The following items encountered errors during processing. Please review and resolve manually.</p>
            </div>
            """, unsafe_allow_html=True)
            
            with st.expander("🔍 View Error Details", expanded=True):
                error_display = error_df[['barcode', 'po_number', 'message']].copy()
                st.dataframe(
                    error_display,
                    use_container_width=True,
                    hide_index=True
                )
            
            # Download error report
            error_csv = error_display.to_csv(index=False)
            st.download_button(
                label="📋 Download Error Report",
                data=error_csv,
                file_name=f"processing_errors_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
        
        st.markdown("</div>", unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #6c757d; padding: 2rem;'>
    <h4>🚀 Odoo PO Scanner Pro</h4>
    <p>Powered by Streamlit | Built for efficiency and accuracy</p>
    <p style='font-size: 0.8rem;'>© 2025 - Advanced Purchase Order Processing System</p>
</div>
""", unsafe_allow_html=True)
