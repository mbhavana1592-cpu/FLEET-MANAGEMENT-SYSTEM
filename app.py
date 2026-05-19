# ================== IMPORT LIBRARIES ==================
import pyodbc
import pandas as pd
import streamlit as st
import warnings
warnings.filterwarnings("ignore")
# ================== USER CONFIG ==================
USERS = {
    "admin": {"password": "admin123", "role": "Admin"},
    "user": {"password": "user123", "role": "User"}
}

# ================== CONFIRMATION HELPER FUNCTION ==================
def confirm_action(action_key: str, message: str):
    """
    Displays a confirmation dialog using Streamlit session state.

    Parameters:
        action_key (str): Unique key for the confirmation action.
        message (str): Confirmation message to display.

    Returns:
        bool: True if the user confirms the action, otherwise False.
    """

    # Initialize session state for this action if not present
    if f"{action_key}_trigger" not in st.session_state:
        st.session_state[f"{action_key}_trigger"] = False
    if f"{action_key}_confirmed" not in st.session_state:
        st.session_state[f"{action_key}_confirmed"] = False

    # First click: show confirmation message
    if not st.session_state[f"{action_key}_trigger"]:
        if st.button("Submit", key=f"{action_key}_submit"):
            st.session_state[f"{action_key}_trigger"] = True
            st.rerun()
        return False

    # Confirmation UI
    st.warning(message)
    col1, col2 = st.columns(2)

    with col1:
        if st.button("✅ Confirm", key=f"{action_key}_confirm"):
            st.session_state[f"{action_key}_confirmed"] = True
            st.session_state[f"{action_key}_trigger"] = False
            return True

    with col2:
        if st.button("❌ Cancel", key=f"{action_key}_cancel"):
            st.session_state[f"{action_key}_trigger"] = False
            st.session_state[f"{action_key}_confirmed"] = False
            st.info("Action cancelled.")
            st.rerun()

    return False
# ================== PAGE CONFIG ==================
st.set_page_config(
    page_title="Fleet Management System",
    layout="wide",
    page_icon="🚛"
)

# ================== DATABASE CONFIGURATION ==================
DB_PATH = r"C:\Users\Administrator\Desktop\project fms\FMS db.accdb"

@st.cache_resource
def get_connection():
    """Create and cache a connection to the MS Access database."""
    try:
        conn = pyodbc.connect(
            rf'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={DB_PATH};'
        )
        return conn
    except Exception as e:
        st.error(f"❌ Database connection failed: {e}")
        st.stop()
# ================== SESSION STATE ==================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = ""

if "role" not in st.session_state:
    st.session_state.role = ""

if "page" not in st.session_state:
    st.session_state.page = "login"
# ================== SESSION STATE ==================
if "page" not in st.session_state:
    st.session_state.page = "home"

def navigate(page_name):
    st.session_state.page = page_name
    st.rerun()

# ================== CUSTOM CSS ==================
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #dfe9f3, #ffffff);
}
.main-title {
    text-align: center;
    font-size: 40px;
    font-weight: bold;
    color: #2E86C1;
    margin-bottom: 30px;
}
.card {
    background: linear-gradient(135deg, #ffffff, #f0f4f8);
    padding: 30px;
    border-radius: 15px;
    text-align: center;
    box-shadow: 0px 4px 15px rgba(0,0,0,0.1);
    transition: transform 0.3s, box-shadow 0.3s;
    cursor: pointer;
}
.card:hover {
    transform: scale(1.05);
    box-shadow: 0px 8px 25px rgba(0,0,0,0.2);
}
.icon {
    font-size: 50px;
    margin-bottom: 10px;
}
.label {
    font-size: 18px;
    font-weight: bold;
    color: #34495E;
}
</style>
""", unsafe_allow_html=True)

# ================== DASHBOARD ANALYTICS ==================
def dashboard_analytics():
    st.markdown("## 📊 Dashboard Analytics")

    conn = get_connection()
    if conn is None:
        st.error("❌ Unable to connect to the database.")
        st.stop()

    # ---------------- Helper Function ----------------
    def count_distinct(table, column):
        query = f"""
            SELECT COUNT(*) AS cnt
            FROM (SELECT DISTINCT {column} FROM {table}) AS Temp
        """
        return pd.read_sql(query, conn)['cnt'][0]

    # ---------------- KPIs ----------------
    col1, col2, col3, col4, col5 = st.columns(5)
    total_materials = pd.read_sql(
        "SELECT COUNT(*) AS cnt FROM Material_Master", conn
    )['cnt'][0]

    total_pos = count_distinct("Purchase_Order", "PO_ID")

    total_vehicles = pd.read_sql(
        "SELECT COUNT(*) AS cnt FROM Vehicle_Master", conn
    )['cnt'][0]

    total_maint_orders = count_distinct("Maintenance_Log", "Main_ID")

    col1.metric("📦 Materials", total_materials)
    col2.metric("🧾 Purchase Orders", total_pos)
    col3.metric("🚗 Vehicles", total_vehicles)
    col4.metric("🛠️ Maintenance Orders", total_maint_orders)

    # ---------------- Additional KPIs ----------------
    col6, col7 = st.columns(2)

    active_vehicles = pd.read_sql(
        "SELECT COUNT(*) AS cnt FROM Vehicle_Master WHERE Status='Active'",
        conn
    )['cnt'][0]

    total_maint_cost = pd.read_sql(
        "SELECT SUM(Cost) AS total FROM Maintenance_Log", conn
    )['total'][0] or 0

    col5.metric("✅ Active Vehicles", active_vehicles)
    col6.metric("💰 Total Maintenance Cost (₹)", f"{total_maint_cost:,.2f}")

    # ---------------- Vehicles by Plant ----------------
    st.subheader("🚗 Vehicles by Plant")
    df_plant = pd.read_sql("""
        SELECT p.PlantDesc, COUNT(v.Vehicle_ID) AS Vehicle_Count
        FROM Vehicle_Master AS v
        LEFT JOIN Plant_Master AS p ON v.Plant = p.Plant
        GROUP BY p.PlantDesc
        ORDER BY COUNT(v.Vehicle_ID) DESC
    """, conn)

    if not df_plant.empty:
        st.bar_chart(df_plant.set_index("PlantDesc"))
    else:
        st.info("No vehicle data available.")

    # ---------------- Vehicles by Type ----------------
    st.subheader("🚚 Vehicles by Type")
    df_type = pd.read_sql("""
        SELECT Vehicle_Type, COUNT(*) AS Count
        FROM Vehicle_Master
        GROUP BY Vehicle_Type
        ORDER BY COUNT(*) DESC
    """, conn)

    if not df_type.empty:
        st.bar_chart(df_type.set_index("Vehicle_Type"))
    else:
        st.info("No vehicle type data available.")

    # ---------------- Maintenance Cost by Vehicle ----------------
    st.subheader("🛠️ Maintenance Cost by Vehicle")
    df_maint = pd.read_sql("""
        SELECT Vehicle_ID, SUM(Cost) AS Total_Cost
        FROM Maintenance_Log
        GROUP BY Vehicle_ID
        ORDER BY SUM(Cost) DESC
    """, conn)

    if not df_maint.empty:
        st.bar_chart(df_maint.set_index("Vehicle_ID"))
    else:
        st.info("No maintenance data available.")

    # ---------------- Purchase Order Value by Vendor ----------------
    st.subheader("🧾 Purchase Order Value by Vendor")
    df_vendor = pd.read_sql("""
        SELECT Vendor_Name, SUM(Total_Price) AS Total_PO_Value
        FROM Purchase_Order
        GROUP BY Vendor_Name
        ORDER BY SUM(Total_Price) DESC
    """, conn)

    if not df_vendor.empty:
        st.bar_chart(df_vendor.set_index("Vendor_Name"))
    else:
        st.info("No purchase order data available.")

    # ---------------- Recent Purchase Orders ----------------
    st.subheader("📄 Recent Purchase Orders")
    df_recent_po = pd.read_sql("""
        SELECT TOP 5 PO_ID, Vendor_Name, Order_Date, Total_Price
        FROM Purchase_Order
        ORDER BY Order_Date DESC
    """, conn)

    st.dataframe(df_recent_po, use_container_width=True)

    # ---------------- Recent Maintenance Records ----------------
    st.subheader("🛠️ Recent Maintenance Activities")
    df_recent_maint = pd.read_sql("""
        SELECT TOP 5 Main_ID, Vehicle_ID, Maintenance_Type,
               Maint_Start_Date, Cost
        FROM Maintenance_Log
        ORDER BY Maint_Start_Date DESC
    """, conn)

    st.dataframe(df_recent_maint, use_container_width=True)
    # ---------------- UTILISATION ANALYTICS ----------------
    st.subheader("⏱️ Vehicle Utilisation Overview")

    try:
        df_util = pd.read_sql("SELECT * FROM Utilisation", conn)
    except:
        df_util = pd.DataFrame()

    if not df_util.empty:

        df_util["Util_Date"] = pd.to_datetime(df_util["Util_Date"])

        col1, col2 = st.columns(2)

        with col1:
            selected_vehicle = st.selectbox(
                "Filter by Vehicle",
                ["All"] + df_util["Vehicle_ID"].unique().tolist(),
                key="dash_util_vehicle"
            )

        with col2:
            selected_date = st.date_input(
                "Filter by Date",
                key="dash_util_date"
            )

        df_filtered = df_util.copy()

        if selected_vehicle != "All":
            df_filtered = df_filtered[df_filtered["Vehicle_ID"] == selected_vehicle]

        if selected_date:
            df_filtered = df_filtered[
                df_filtered["Util_Date"] == pd.to_datetime(selected_date)
                ]

        if not df_filtered.empty:

            # ---------------- SUMMARY ----------------
            st.subheader("📊 Utilisation Summary")

            summary = df_filtered[["Worked_hrs", "Idle_hrs", "BD_hrs"]].sum()
            st.bar_chart(summary)

            # ---------------- TREND ----------------
            st.subheader("📈 Daily Trend")

            trend = df_filtered.sort_values("Util_Date")
            trend = trend.set_index("Util_Date")[["Worked_hrs", "Idle_hrs", "BD_hrs"]]

            st.line_chart(trend)

            # ---------------- TABLE ----------------
            st.write(summary)

        else:
            st.info("No utilisation data for selected filters.")

    else:
        st.info("No utilisation records available.")
def login_page():
    st.markdown("""
        <h1 style='text-align:center;color:#1F4E79;'>🚛 Fleet Management System</h1>
        <h3 style='text-align:center;'>🔐 Login</h3>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1,2,1])

    with col2:
        username = st.text_input("👤 Username")
        password = st.text_input("🔑 Password", type="password")

        if st.button("Login"):
            if username in USERS and USERS[username]["password"] == password:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.role = USERS[username]["role"]
                st.session_state.page = "home"
                st.success(f"✅ Welcome {username} ({st.session_state.role})")
                st.rerun()
            else:
                st.error("❌ Invalid Username or Password")
# ================== HOME PAGE ==================
def home():
    st.markdown(
        '<div class="main-title">🚛 Fleet Management System</div>',
        unsafe_allow_html=True
    )

    # ✅ FIXED ORDER (as you requested)
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)

    # ================== MATERIAL ==================
    with col1:
        if st.button("📦", key="mat"):
            navigate("material")
        st.markdown(
            '<div class="card"><div class="icon">📦</div><div class="label">Material Master</div></div>',
            unsafe_allow_html=True
        )

    # ================== PURCHASE ORDER ==================
    with col2:
        if st.button("🧾", key="po"):
            navigate("po")
        st.markdown(
            '<div class="card"><div class="icon">🧾</div><div class="label">Purchase Order</div></div>',
            unsafe_allow_html=True
        )

    # ================== VEHICLE MASTER ==================
    with col3:
        if st.button("🚗", key="veh"):
            navigate("vehicle")
        st.markdown(
            '<div class="card"><div class="icon">🚗</div><div class="label">Vehicle Master</div></div>',
            unsafe_allow_html=True
        )

    # ================== MAINTENANCE ==================
    with col4:
        if st.button("🛠️", key="maint"):
            navigate("maintenance")
        st.markdown(
            '<div class="card"><div class="icon">🛠️</div><div class="label">Maintenance</div></div>',
            unsafe_allow_html=True
        )

    # ================== UTILISATION ==================
    with col5:
        if st.button("⏱️", key="util"):
            navigate("utilisation")
        st.markdown(
            '<div class="card"><div class="icon">⏱️</div><div class="label">Utilisation</div></div>',
            unsafe_allow_html=True
        )

    # ================== DASHBOARD ==================
    with col6:
        if st.button("📊", key="dashboard"):
            navigate("dashboard")
        st.markdown(
            '<div class="card"><div class="icon">📊</div><div class="label">Dashboard Analytics</div></div>',
            unsafe_allow_html=True
        )

    # ================== REPORTS ==================
    with col7:
        if st.button("📄", key="reports"):
            navigate("reports")
        st.markdown(
            '<div class="card"><div class="icon">📄</div><div class="label">Reports</div></div>',
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.info("📌 Select a module to manage fleet operations.")
with st.sidebar:
    st.markdown("## 🚛 FMS Menu")

    # 👤 User Info
    if st.session_state.logged_in:
        st.write(f"👤 {st.session_state.username}")
        st.write(f"🔑 Role: {st.session_state.role}")

    # KEEP ALL YOUR EXISTING BUTTONS BELOW 👇
    if st.button("🏠 Home"):
        navigate("home")

    if st.button("📦 Material"):
        navigate("material")

    if st.button("🧾 Purchase Order"):
        navigate("po")

    if st.button("🚗 Vehicle"):
        navigate("vehicle")

    if st.button("🛠️ Maintenance"):
        navigate("maintenance")

    if st.button("⏱️ Utilisation"):
        navigate("utilisation")

    if st.button("📊 Dashboard"):
        navigate("dashboard")

    if st.button("📄 Reports"):
        navigate("reports")

    st.markdown("---")

    # 🔐 Logout
    if st.session_state.logged_in:
        if st.button("🚪 Logout"):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.role = ""
            st.session_state.page = "login"
            st.rerun()
# ================== MATERIAL MASTER ==================
def material_page():
    st.title("📦 Material Master")

    if st.button("⬅ Back"):
        navigate("home")

    conn = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\Administrator\Desktop\project fms\FMS db.accdb;'
    )
    cursor = conn.cursor()

    st.markdown("## 🛠️ Material Management")

    tab1, tab2, tab3 = st.tabs(["➕ Create", "✏️ Update", "❌ Delete"])

    # -------- CREATE --------
    with tab1:

        def generate_material_id():
            cursor.execute("SELECT MAX(Material) FROM Material_Master")
            result = cursor.fetchone()[0]
            return "100000" if result is None else str(int(result) + 1).zfill(6)

        new_material = generate_material_id()

        st.text_input("Material ID", value=new_material, disabled=True)
        material_desc = st.text_input("Material Description")
        material_type = st.selectbox("Material Type", ["Vehicle", "Spare", "Fuel"])
        base_uom = st.selectbox("Base UOM", ["EA", "NOS", "LTR"])
        standard_price = st.number_input(
            "Standard Price (INR)",
            min_value=0.0,
            max_value=9999999999999.99,
            format="%.2f"
        )

        # -------- VALIDATION --------
        errors = []
        if material_desc.strip() == "":
            errors.append("Material Description is required.")
        if len(material_desc) > 80:
            errors.append("Material Description must not exceed 80 characters.")
        if standard_price <= 0:
            errors.append("Standard Price must be greater than zero.")

        if errors:
            for error in errors:
                st.error(f"❌ {error}")

        # -------- CONFIRMATION & INSERT --------
        if not errors and confirm_action(
                "confirm_create_material",
                f"Are you sure you want to create Material **{new_material}**?"
        ):
            try:
                cursor.execute("""
                    INSERT INTO Material_Master
                    (Material, MaterialDesc, Material_Type, BaseUOM, Standard_Price)
                    VALUES (?, ?, ?, ?, ?)
                """, (
                    new_material,
                    material_desc.strip(),
                    material_type,
                    base_uom,
                    standard_price
                ))
                conn.commit()
                st.success(f"✅ Material {new_material} created successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Failed to create material: {e}")

    # -------- UPDATE --------
    with tab2:

        df = pd.read_sql("SELECT * FROM Material_Master", conn)

        if df.empty:
            st.warning("No materials available")
        else:
            selected_mat = st.selectbox("Select Material", df["Material"], key="update_material")

            row = df[df["Material"] == selected_mat].iloc[0]

            upd_desc = st.text_input("Description", value=row["MaterialDesc"])
            upd_type = st.selectbox("Type", ["Vehicle", "Spare", "Fuel"],
                                    index=["Vehicle", "Spare", "Fuel"].index(row["Material_Type"]))
            upd_uom = st.selectbox("UOM", ["EA", "NOS", "LTR"],
                                   index=["EA", "NOS", "LTR"].index(row["BaseUOM"]))
            upd_price = st.number_input("Price", value=float(row["Standard_Price"]), format="%.2f")

            if st.button("Update Material"):
                if upd_desc.strip() == "":
                    st.error("❌ Description required")
                elif len(upd_desc) > 80:
                    st.error("❌ Max 80 characters")
                else:
                    cursor.execute("""
                        UPDATE Material_Master
                        SET MaterialDesc=?, Material_Type=?, BaseUOM=?, Standard_Price=?
                        WHERE Material=?
                    """, (upd_desc, upd_type, upd_uom, upd_price, selected_mat))
                    conn.commit()
                    st.success("✅ Updated successfully")
                    st.rerun()

    # -------- DELETE --------
    with tab3:

        df = pd.read_sql("SELECT * FROM Material_Master", conn)

        if not df.empty:
            del_mat = st.selectbox("Select Material", df["Material"], key="delete_material")

            if st.button("Delete Material"):
                cursor.execute("DELETE FROM Material_Master WHERE Material=?", (del_mat,))
                conn.commit()
                st.warning("⚠️ Deleted")
                st.rerun()

    # -------- VIEW --------
    st.markdown("### 📋 Material List")
    df = pd.read_sql("SELECT * FROM Material_Master", conn)
    st.write(f"Total Materials: {len(df)}")
    st.dataframe(df)

# ---------------- OTHER PAGES ----------------
def po_page():
    st.title("🧾 Purchase Order")

    # ---------------- SESSION STATE ----------------
    if "po_success" not in st.session_state:
        st.session_state.po_success = ""

    if "po_items" not in st.session_state:
        st.session_state.po_items = []

    if "last_selected_po" not in st.session_state:
        st.session_state.last_selected_po = None

    # 🔥 FIX: INIT UPDATE STATE
    if "po_vendor_update" not in st.session_state:
        st.session_state.po_vendor_update = None

    if "po_qty_update" not in st.session_state:
        st.session_state.po_qty_update = 1

    if "po_price_update" not in st.session_state:
        st.session_state.po_price_update = 0.0

    # ---------------- SHOW MESSAGE ----------------
    if st.session_state.po_success:
        st.success(st.session_state.po_success)
        st.session_state.po_success = ""

    if st.button("⬅ Back", key="po_back"):
        navigate("home")

    # ---------------- DB CONNECTION ----------------
    conn = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\Administrator\Desktop\project fms\FMS db.accdb;'
    )
    cursor = conn.cursor()

    st.markdown("## 🧾 Purchase Order Management")

    # ---------------- VENDOR LIST ----------------
    VENDOR_LIST = [
        "Tata Motors",
        "Ashok Leyland",
        "Bharat Benz",
        "L&T Heavy Equipment",
        "Caterpillar India",
        "Komatsu India"
    ]

    # ---------------- PO NUMBER ----------------
    def generate_po_id():
        cursor.execute("SELECT MAX(PO_ID) FROM Purchase_Order")
        result = cursor.fetchone()[0]

        if result is None:
            return "PO000001"
        else:
            last_num = int(result.replace("PO", ""))
            return f"PO{str(last_num + 1).zfill(6)}"

    tab1, tab2, tab3 = st.tabs(["➕ Create", "✏️ Update", "❌ Delete"])

    # =========================================================
    # -------- CREATE --------
    with tab1:

        def generate_po_id():
            cursor.execute("SELECT MAX(PO_ID) FROM Purchase_Order")
            result = cursor.fetchone()[0]
            if result is None:
                return "PO000001"
            else:
                return f"PO{str(int(result.replace('PO', '')) + 1).zfill(6)}"

        po_id = generate_po_id()

        st.text_input("Purchase Order ID", value=po_id, disabled=True)

        # -------- VENDOR --------
        VENDOR_LIST = [
            "Tata Motors",
            "Ashok Leyland",
            "Bharat Benz",
            "L&T Heavy Equipment",
            "Caterpillar India",
            "Komatsu India"
        ]
        vendor_name = st.selectbox("Vendor Name", VENDOR_LIST)

        # -------- MATERIAL --------
        df_mat = pd.read_sql("SELECT Material, MaterialDesc FROM Material_Master", conn)

        if df_mat.empty:
            st.error("❌ No materials available. Please create a material first.")
            material = None
        else:
            material_dict = {
                f"{row['Material']} - {row['MaterialDesc']}": row['Material']
                for _, row in df_mat.iterrows()
            }

            selected_display = st.selectbox(
                "Select Material",
                list(material_dict.keys())
            )
            material = material_dict[selected_display]

        # -------- QUANTITY & PRICE --------
        quantity = st.number_input(
            "Quantity",
            min_value=1,
            step=1
        )

        unit_price = st.number_input(
            "Unit Price (INR)",
            min_value=0.0,
            max_value=9999999999999.99,
            format="%.2f"
        )

        total_price = quantity * unit_price
        st.success(f"🧮 Total Price: ₹ {total_price:,.2f}")

        # -------- VALIDATION --------
        errors = []

        if material is None:
            errors.append("Material selection is required.")
        if quantity <= 0:
            errors.append("Quantity must be greater than zero.")
        if unit_price <= 0:
            errors.append("Unit Price must be greater than zero.")

        if errors:
            for error in errors:
                st.error(f"❌ {error}")

        # -------- CONFIRMATION & INSERT --------
        if not errors and confirm_action(
                "confirm_create_po",
                f"Are you sure you want to create Purchase Order **{po_id}**?"
        ):
            try:
                po_item_id = f"{po_id}_1"

                cursor.execute("""
                    INSERT INTO Purchase_Order
                    (PO_Item_ID, PO_ID, Vendor_Name, Material,
                     Quantity, Unit_Price, Total_Price, Order_Date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, Date())
                """, (
                    po_item_id,
                    po_id,
                    vendor_name,
                    material,
                    quantity,
                    unit_price,
                    total_price
                ))

                conn.commit()
                st.success(f"✅ Purchase Order {po_id} created successfully!")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Failed to create Purchase Order: {e}")
    # 🔹 UPDATE
    # =========================================================
    with tab2:

        df_po = pd.read_sql("SELECT * FROM Purchase_Order", conn)

        if df_po.empty:
            st.warning("No Purchase Orders available")
        else:
            selected_po = st.selectbox(
                "Select PO Item",
                df_po["PO_Item_ID"],
                key="po_update_select"
            )

            row = df_po[df_po["PO_Item_ID"] == selected_po].iloc[0]

            # 🔥 SYNC SESSION STATE
            if st.session_state.last_selected_po != selected_po:
                st.session_state.last_selected_po = selected_po

                st.session_state.po_vendor_update = row["Vendor_Name"]
                st.session_state.po_qty_update = int(row["Quantity"])
                st.session_state.po_price_update = float(row["Unit_Price"])

            # SAFE INDEX
            vendor_index = (
                VENDOR_LIST.index(st.session_state.po_vendor_update)
                if st.session_state.po_vendor_update in VENDOR_LIST
                else 0
            )

            upd_vendor = st.selectbox(
                "Vendor",
                VENDOR_LIST,
                index=vendor_index,
                key="po_vendor_update"
            )

            df_mat = pd.read_sql("SELECT Material, MaterialDesc FROM Material_Master", conn)

            material_dict = {
                f"{r['Material']} - {r['MaterialDesc']}": r['Material']
                for _, r in df_mat.iterrows()
            }

            selected_display = st.selectbox(
                "Material",
                list(material_dict.keys()),
                key="po_material_update"
            )

            upd_material = material_dict[selected_display]

            upd_qty = st.number_input(
                "Quantity",
                min_value=1,
                key="po_qty_update"
            )

            upd_price = st.number_input(
                "Unit Price",
                min_value=0.0,
                format="%.2f",
                key="po_price_update"
            )

            upd_total = upd_qty * upd_price
            st.success(f"🧮 Total Price: ₹ {upd_total:,.2f}")

            if st.button("Update Purchase Order", key="update_po"):

                cursor.execute("""
                    UPDATE Purchase_Order
                    SET Vendor_Name=?, Material=?, Quantity=?, Unit_Price=?, Total_Price=?
                    WHERE PO_Item_ID=?
                """, (upd_vendor, upd_material, upd_qty, upd_price, upd_total, selected_po))

                conn.commit()

                st.session_state.po_success = "✅ PO Updated"
                st.rerun()

    # =========================================================
    # 🔹 DELETE
    # =========================================================
    with tab3:

        df_po = pd.read_sql("SELECT * FROM Purchase_Order", conn)

        if not df_po.empty:
            del_po = st.selectbox(
                "Select PO Item",
                df_po["PO_Item_ID"],
                key="po_delete_select"
            )

            if st.button("Delete Purchase Order", key="delete_po"):
                cursor.execute("DELETE FROM Purchase_Order WHERE PO_Item_ID=?", (del_po,))
                conn.commit()

                st.session_state.po_success = "⚠️ PO Deleted"
                st.rerun()

    # ---------------- VIEW ----------------
    st.markdown("### 📋 Purchase Orders")

    df_po = pd.read_sql("SELECT * FROM Purchase_Order", conn)
    st.write(f"Total Records: {len(df_po)}")
    st.dataframe(df_po)

    if st.button("⬅ Back", key="po_back_bottom"):
        navigate("home")

def vehicle_page():
    st.title("🚗 Vehicle Master")

    # ---------------- SESSION ----------------
    if "vehicle_success" not in st.session_state:
        st.session_state.vehicle_success = ""

    if st.session_state.vehicle_success:
        st.success(st.session_state.vehicle_success)
        st.session_state.vehicle_success = ""

    if st.button("⬅ Back", key="veh_back"):
        navigate("home")

    # ---------------- DB CONNECTION ----------------
    conn = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\Administrator\Desktop\project fms\FMS db.accdb;'
    )
    cursor = conn.cursor()

    st.markdown("## 🚗 Vehicle Management")

    # ---------------- VEHICLE ID GENERATOR ----------------
    def generate_vehicle_id():
        cursor.execute("SELECT MAX(Vehicle_ID) FROM Vehicle_Master")
        result = cursor.fetchone()[0]

        if result is None:
            return "V000001"
        else:
            last_num = int(result.replace("V", ""))
            return f"V{str(last_num + 1).zfill(6)}"

    tab1, tab2, tab3 = st.tabs(["➕ Create", "✏️ Update", "❌ Delete"])

    # =========================================================
    # -------- CREATE --------
    with tab1:

        # -------- VEHICLE ID GENERATOR --------
        def generate_vehicle_id():
            cursor.execute("SELECT MAX(Vehicle_ID) FROM Vehicle_Master")
            result = cursor.fetchone()[0]
            if result is None:
                return "V000001"
            else:
                return f"V{str(int(result.replace('V', '')) + 1).zfill(6)}"

        vehicle_id = generate_vehicle_id()

        st.text_input("Vehicle ID", value=vehicle_id, disabled=True)

        # -------- MATERIAL --------
        df_mat = pd.read_sql("SELECT Material, MaterialDesc FROM Material_Master", conn)

        if df_mat.empty:
            st.error("❌ No materials available. Please create a material first.")
            material = None
        else:
            material_dict = {
                f"{row['Material']} - {row['MaterialDesc']}": row['Material']
                for _, row in df_mat.iterrows()
            }

            selected_display = st.selectbox(
                "Select Material",
                list(material_dict.keys())
            )
            material = material_dict[selected_display]

        # -------- PURCHASE ORDER (FILTERED BY MATERIAL) --------
        if material:
            df_po = pd.read_sql("""
                SELECT DISTINCT PO_ID, Vendor_Name
                FROM Purchase_Order
                WHERE Material = ?
            """, conn, params=(material,))

            if df_po.empty:
                st.error("❌ No Purchase Orders found for the selected material.")
                po_id = None
                vendor_name = None
            else:
                po_dict = {
                    f"{row['PO_ID']} - {row['Vendor_Name']}": row['PO_ID']
                    for _, row in df_po.iterrows()
                }

                selected_po_display = st.selectbox(
                    "Select Purchase Order",
                    list(po_dict.keys())
                )
                po_id = po_dict[selected_po_display]
                vendor_name = selected_po_display.split(" - ")[1]

                st.info(f"🏢 Vendor: {vendor_name}")
        else:
            po_id = None
            vendor_name = None

        # -------- VEHICLE DETAILS --------
        vehicle_number = st.text_input("Vehicle Number")

        vehicle_type = st.selectbox(
            "Vehicle Type",
            ["Truck", "Dumper", "Excavator", "Dozer", "Other"]
        )

        fuel_type = st.selectbox(
            "Fuel Type",
            ["Diesel", "Petrol", "Electric"]
        )

        commissioning_date = st.date_input("Date of Commissioning")

        # -------- PLANT --------
        df_plant = pd.read_sql("SELECT [Plant], [PlantDesc] FROM [Plant_Master]", conn)

        if df_plant.empty:
            st.error("❌ No plants available. Please create a plant first.")
            plant = None
        else:
            plant_dict = {
                f"{row['Plant']} - {row['PlantDesc']}": row['Plant']
                for _, row in df_plant.iterrows()
            }

            selected_plant_display = st.selectbox(
                "Select Plant",
                list(plant_dict.keys())
            )
            plant = plant_dict[selected_plant_display]

        status = "Active"

        # -------- VALIDATION --------
        errors = []

        if material is None:
            errors.append("Material selection is required.")
        if po_id is None:
            errors.append("Purchase Order selection is required.")
        if vehicle_number.strip() == "":
            errors.append("Vehicle Number is required.")
        if plant is None:
            errors.append("Plant selection is required.")

        if errors:
            for error in errors:
                st.error(f"❌ {error}")

        # -------- CONFIRMATION & INSERT --------
        if not errors and confirm_action(
                "confirm_create_vehicle",
                f"Are you sure you want to create Vehicle **{vehicle_id}**?"
        ):
            try:
                cursor.execute("""
                    INSERT INTO Vehicle_Master
                    (Vehicle_ID, Material, Vehicle_Number, Vehicle_Type,
                     Fuel_Type, Date_of_Commissioning, Plant, Status, PO_ID)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    vehicle_id,
                    material,
                    vehicle_number.strip(),
                    vehicle_type,
                    fuel_type,
                    commissioning_date,
                    plant,
                    status,
                    po_id
                ))

                conn.commit()
                st.success(f"✅ Vehicle {vehicle_id} created successfully!")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Failed to create vehicle: {e}")
    # =========================================================
    # 🔹 UPDATE
    # =========================================================
    with tab2:

        df = pd.read_sql("SELECT * FROM Vehicle_Master", conn)

        if df.empty:
            st.warning("No Vehicles Available")
        else:
            selected_vehicle = st.selectbox(
                "Select Vehicle",
                df["Vehicle_ID"],
                key="veh_update_select"
            )

            row = df[df["Vehicle_ID"] == selected_vehicle].iloc[0]

            # -------- SAFE JOIN FOR PO + VENDOR --------
            cursor.execute("""
                SELECT v.PO_ID, p.Vendor_Name
                FROM Vehicle_Master v
                LEFT JOIN Purchase_Order p ON v.PO_ID = p.PO_ID
                WHERE v.Vehicle_ID=?
            """, (selected_vehicle,))

            result = cursor.fetchone()

            if result:
                po_id, vendor_name = result
                vendor_name = vendor_name if vendor_name else "Unknown / Deleted PO"
            else:
                po_id, vendor_name = "N/A", "N/A"

            st.info(f"🧾 PO: {po_id}")
            st.info(f"🏢 Vendor: {vendor_name}")

            # -------- UPDATE FIELDS --------
            upd_number = st.text_input("Vehicle Number", value=row["Vehicle_Number"], key="veh_upd_num")

            upd_type = st.selectbox(
                "Vehicle Type",
                ["Truck", "Dumper", "Excavator", "Dozer", "Other"],
                index=["Truck", "Dumper", "Excavator", "Dozer", "Other"].index(row["Vehicle_Type"]),
                key="veh_upd_type"
            )

            upd_fuel = st.selectbox(
                "Fuel Type",
                ["Diesel", "Petrol", "Electric"],
                index=["Diesel", "Petrol", "Electric"].index(row["Fuel_Type"]),
                key="veh_upd_fuel"
            )

            upd_date = st.date_input("Date", value=row["Date_of_Commissioning"], key="veh_upd_date")

            # -------- PLANT --------
            df_plant = pd.read_sql("SELECT [Plant], [PlantDesc] FROM [Plant_Master]", conn)

            plant_dict = {
                f"{r['Plant']} - {r['PlantDesc']}": r['Plant']
                for _, r in df_plant.iterrows()
            }

            plant_list = list(plant_dict.keys())

            current_plant_display = None
            for key, val in plant_dict.items():
                if val == row["Plant"]:
                    current_plant_display = key
                    break

            upd_plant_display = st.selectbox(
                "Plant",
                plant_list,
                index=plant_list.index(current_plant_display),
                key="veh_upd_plant"
            )

            upd_plant = plant_dict[upd_plant_display]

            upd_status = st.selectbox(
                "Status",
                ["Active", "Inactive"],
                index=["Active", "Inactive"].index(row["Status"]),
                key="veh_upd_status"
            )

            if st.button("Update Vehicle", key="update_vehicle"):

                if upd_number.strip() == "":
                    st.error("❌ Vehicle Number required")

                else:
                    cursor.execute("""
                        UPDATE Vehicle_Master
                        SET Vehicle_Number=?, Vehicle_Type=?, Fuel_Type=?, Date_of_Commissioning=?, Plant=?, Status=?
                        WHERE Vehicle_ID=?
                    """, (
                        upd_number,
                        upd_type,
                        upd_fuel,
                        upd_date,
                        upd_plant,
                        upd_status,
                        selected_vehicle
                    ))

                    conn.commit()
                    st.session_state.vehicle_success = "✅ Vehicle Updated"
                    st.rerun()

    # =========================================================
    # 🔹 DELETE
    # =========================================================
    with tab3:

        df = pd.read_sql("SELECT * FROM Vehicle_Master", conn)

        if not df.empty:
            del_vehicle = st.selectbox(
                "Select Vehicle",
                df["Vehicle_ID"],
                key="veh_delete_select"
            )

            if st.button("Delete Vehicle", key="delete_vehicle"):
                cursor.execute("DELETE FROM Vehicle_Master WHERE Vehicle_ID=?", (del_vehicle,))
                conn.commit()

                st.session_state.vehicle_success = "⚠️ Vehicle Deleted"
                st.rerun()

    # ---------------- VIEW ----------------
    st.markdown("### 📋 Vehicle List")

    df = pd.read_sql("SELECT * FROM Vehicle_Master", conn)
    st.write(f"Total Vehicles: {len(df)}")
    st.dataframe(df)
    if st.button("⬅ Back"):
        navigate("home")
def maintenance_page():
    st.title("🛠️ Maintenance Log")

    # ---------------- SESSION ----------------
    if "maint_success" not in st.session_state:
        st.session_state.maint_success = ""

    if "maint_items" not in st.session_state:
        st.session_state.maint_items = []

    if st.session_state.maint_success:
        st.success(st.session_state.maint_success)
        st.session_state.maint_success = ""

    if st.button("⬅ Back", key="maint_back"):
        navigate("home")

    # ---------------- DB CONNECTION ----------------
    conn = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\Administrator\Desktop\project fms\FMS db.accdb;'
    )
    cursor = conn.cursor()

    st.markdown("## 🛠️ Maintenance Management")

    # ---------------- MAIN ID GENERATOR ----------------
    def generate_main_id():
        cursor.execute("SELECT MAX(Main_ID) FROM Maintenance_Log")
        result = cursor.fetchone()[0]

        if result is None:
            return "M000001"
        else:
            last = int(result.replace("M", ""))
            return f"M{str(last + 1).zfill(6)}"

    tab1, tab2, tab3 = st.tabs(["➕ Create", "✏️ Update/Delete Items", "❌ Delete Order"])

    # =========================================================
    # 🔹 CREATE
    # =========================================================
    with tab1:

        main_id = generate_main_id()
        st.text_input("Maintenance Order ID", value=main_id, disabled=True)

        # -------- VEHICLE --------
        df_veh = pd.read_sql("SELECT Vehicle_ID FROM Vehicle_Master", conn)
        vehicle = st.selectbox("Select Vehicle", df_veh["Vehicle_ID"], key="maint_vehicle")

        # -------- TYPE --------
        maint_type = st.selectbox(
            "Maintenance Type",
            ["Breakdown", "Preventive", "Inspection"],
            key="maint_type"
        )

        description = st.text_input("Description", key="maint_desc")

        # -------- DATE + TIME --------
        start_date = st.date_input("Start Date", key="maint_start_date")
        start_time = st.time_input("Start Time", key="maint_start_time")

        end_date = st.date_input("End Date", key="maint_end_date")
        end_time = st.time_input("End Time", key="maint_end_time")

        # -------- ODOMETER --------
        odo_start = st.number_input("Odometer Start", min_value=0, key="odo_start")
        odo_end = st.number_input("Odometer End", min_value=0, key="odo_end")

        st.markdown("### ➕ Add Materials")

        # -------- MATERIAL --------
        df_mat = pd.read_sql("""
            SELECT Material, MaterialDesc
            FROM Material_Master
            WHERE Material_Type <> 'Vehicle'
        """, conn)

        mat_dict = {
            f"{r['Material']} - {r['MaterialDesc']}": r['Material']
            for _, r in df_mat.iterrows()
        }

        selected_mat = st.selectbox("Material", list(mat_dict.keys()), key="maint_material")
        material = mat_dict[selected_mat]

        quantity = st.number_input("Quantity", min_value=1, key="maint_qty")

        cost = st.number_input(
            "Cost (INR)",
            min_value=0.0,
            format="%.2f",
            key="maint_cost"
        )

        # ---------------- VALIDATION ----------------
        errors = []

        if vehicle is None:
            errors.append("Vehicle is required")
        if description.strip() == "":
            errors.append("Description is required")
        if quantity <= 0:
            errors.append("Quantity must be greater than 0")
        if cost <= 0:
            errors.append("Cost must be greater than 0")
        if odo_end < odo_start:
            errors.append("Odometer End cannot be less than Start")
        if end_date < start_date:
            errors.append("End Date cannot be before Start Date")

        if errors:
            for e in errors:
                st.error(f"❌ {e}")

        # ---------------- ADD ITEM ----------------
        if st.button("Add Item", key="add_maint_item"):
            st.session_state.maint_items.append({
                "Material": material,
                "Quantity": quantity,
                "Cost": cost
            })

        # ---------------- SHOW ITEMS ----------------
        if st.session_state.maint_items:
            df_items = pd.DataFrame(st.session_state.maint_items)
            st.dataframe(df_items)

            total_cost = df_items["Cost"].sum()
            st.success(f"💰 Total Maintenance Cost: ₹ {total_cost:,.2f}")

        # ================== SUBMIT CLICK ==================
        if "maint_submit_clicked" not in st.session_state:
            st.session_state.maint_submit_clicked = False

        if st.button("Create Maintenance", key="create_maint"):
            st.session_state.maint_submit_clicked = True

        # ================== AFTER SUBMIT ==================
        if st.session_state.maint_submit_clicked:

            # ---------- VALIDATION ----------
            if errors:
                for e in errors:
                    st.error(f"❌ {e}")
                st.session_state.maint_submit_clicked = False

            elif not st.session_state.maint_items:
                st.error("❌ Add at least one material")
                st.session_state.maint_submit_clicked = False

            else:
                # ================== CONFIRM UI ==================
                st.warning(f"Are you sure you want to create Maintenance Order **{main_id}**?")

                col1, col2 = st.columns(2)

                # ✅ CONFIRM BUTTON
                with col1:
                    if st.button("✅ Confirm", key="maint_confirm_btn"):

                        try:
                            for idx, item in enumerate(st.session_state.maint_items, start=1):
                                maint_item_id = f"{main_id}_{idx}"

                                cursor.execute("""
                                    INSERT INTO Maintenance_Log
                                    (Maintenance_ID, Main_ID, Vehicle_ID, Material, Maintenance_Type,
                                     Description, Quantity, Cost, Maint_Start_Date, Maint_Start_Time,
                                     Maint_End_Date, Maint_End_Time, Odometer_Start, Odometer_End)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """, (
                                    maint_item_id,
                                    main_id,
                                    vehicle,
                                    item["Material"],
                                    maint_type,
                                    description,
                                    item["Quantity"],
                                    item["Cost"],
                                    start_date,
                                    start_time,
                                    end_date,
                                    end_time,
                                    odo_start,
                                    odo_end
                                ))

                            conn.commit()

                            st.success(f"✅ Maintenance Created: {main_id}")

                            df_latest = pd.read_sql(
                                "SELECT * FROM Maintenance_Log ORDER BY Main_ID DESC",
                                conn
                            )
                            st.dataframe(df_latest)

                            # ✅ RESET
                            st.session_state.maint_items = []
                            st.session_state.maint_submit_clicked = False

                        except Exception as e:
                            st.error(f"❌ Database Error: {e}")

                # ❌ CANCEL BUTTON (FIXED)
                with col2:
                    if st.button("❌ Cancel", key="maint_cancel_btn"):
                        st.session_state.maint_submit_clicked = False
                        st.info("🚫 Maintenance creation cancelled")

    # 🔹 UPDATE + DELETE ITEM
    # =========================================================
    with tab2:

        df_main = pd.read_sql("SELECT DISTINCT Main_ID FROM Maintenance_Log", conn)

        if df_main.empty:
            st.warning("No Maintenance Records")
        else:
            selected_main = st.selectbox("Select Maintenance Order", df_main["Main_ID"], key="maint_update_main")

            df_items = pd.read_sql(
                "SELECT * FROM Maintenance_Log WHERE Main_ID=?",
                conn,
                params=(selected_main,)
            )

            st.dataframe(df_items)

            selected_item = st.selectbox(
                "Select Item",
                df_items["Maintenance_ID"],
                key="maint_update_item"
            )

            row = df_items[df_items["Maintenance_ID"] == selected_item].iloc[0]

            upd_desc = st.text_input("Description", value=row["Description"])
            upd_qty = st.number_input("Quantity", value=int(row["Quantity"]), min_value=1)
            upd_cost = st.number_input("Cost", value=float(row["Cost"]), min_value=0.0)

            upd_odo_start = st.number_input("Odometer Start", value=int(row["Odometer_Start"]))
            upd_odo_end = st.number_input("Odometer End", value=int(row["Odometer_End"]))

            col1, col2 = st.columns(2)

            # UPDATE
            with col1:
                if st.button("Update Item", key="update_item_btn"):

                    if upd_odo_end < upd_odo_start:
                        st.error("❌ Odometer issue")
                    else:
                        cursor.execute("""
                            UPDATE Maintenance_Log
                            SET Description=?, Quantity=?, Cost=?, 
                                Odometer_Start=?, Odometer_End=?
                            WHERE Maintenance_ID=?
                        """, (
                            upd_desc,
                            upd_qty,
                            upd_cost,
                            upd_odo_start,
                            upd_odo_end,
                            selected_item
                        ))

                        conn.commit()
                        st.session_state.maint_success = "✅ Item Updated"
                        st.rerun()

            # DELETE ITEM
            with col2:
                if st.button("Delete Item", key="delete_item_btn"):

                    cursor.execute(
                        "DELETE FROM Maintenance_Log WHERE Maintenance_ID=?",
                        (selected_item,)
                    )

                    cursor.execute(
                        "SELECT COUNT(*) FROM Maintenance_Log WHERE Main_ID=?",
                        (selected_main,)
                    )

                    remaining = cursor.fetchone()[0]

                    if remaining == 0:
                        st.session_state.maint_success = "⚠️ Last item deleted. Order removed."
                    else:
                        st.session_state.maint_success = "⚠️ Item Deleted"

                    conn.commit()
                    st.rerun()

    # =========================================================
    # 🔹 DELETE FULL ORDER
    # =========================================================
    with tab3:

        df_main = pd.read_sql("SELECT DISTINCT Main_ID FROM Maintenance_Log", conn)

        if not df_main.empty:
            selected_main = st.selectbox("Select Maintenance Order", df_main["Main_ID"], key="maint_delete")

            if st.button("Delete Full Order", key="delete_maint"):

                cursor.execute("DELETE FROM Maintenance_Log WHERE Main_ID=?", (selected_main,))
                conn.commit()

                st.session_state.maint_success = "⚠️ Maintenance Order Deleted"
                st.rerun()


# ================== DASHBOARD PAGE ==================
def dashboard_page():
    st.title("📊 Dashboard Analytics")

    if st.button("⬅ Back", key="dash_back"):
        navigate("home")

    # Call the existing analytics function
    dashboard_analytics()


# ================== UTILISATION PAGE ==================
# ================== UTILISATION PAGE ==================
def utilisation_page():
    st.title("⏱️ Vehicle Utilisation")

    if st.button("⬅ Back"):
        navigate("home")

    # ---------------- SESSION FIX ----------------
    if "util_just_saved" not in st.session_state:
        st.session_state.util_just_saved = False

    if "util_submit_clicked" not in st.session_state:
        st.session_state.util_submit_clicked = False

    # Reset flags properly
    skip_duplicate_check = st.session_state.util_just_saved
    st.session_state.util_just_saved = False

    # ---------------- DB CONNECTION ----------------
    conn = get_connection()
    cursor = conn.cursor()

    st.markdown("## ⏱️ Utilisation Entry")

    # ---------------- CREATE TABLE ----------------
    try:
        cursor.execute("""
            CREATE TABLE Utilisation (
                Util_ID TEXT PRIMARY KEY,
                Util_Date DATE,
                Vehicle_ID TEXT,
                Worked_hrs DOUBLE,
                Idle_hrs DOUBLE,
                BD_hrs DOUBLE
            )
        """)
        conn.commit()
    except:
        pass

    # ---------------- ID GENERATOR ----------------
    def generate_util_id():
        df = pd.read_sql("SELECT Util_ID FROM Utilisation", conn)

        if df.empty:
            return "U000001"

        df["num"] = df["Util_ID"].str.replace("U", "").astype(int)
        next_id = df["num"].max() + 1

        return f"U{next_id:06d}"

    util_id = generate_util_id()
    st.text_input("Utilisation ID", value=util_id, disabled=True)

    # ---------------- INPUT ----------------
    util_date = st.date_input("Date")

    df_veh = pd.read_sql("SELECT Vehicle_ID FROM Vehicle_Master", conn)
    vehicle_id = st.selectbox("Vehicle ID", df_veh["Vehicle_ID"])

    worked = st.number_input("Worked Hours", min_value=0.0, max_value=24.0)

    # ---------------- TIME HANDLING ----------------
    from datetime import datetime, time

    def to_time(val):
        if isinstance(val, str):
            for fmt in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M:%S %p"):
                try:
                    return datetime.strptime(val.strip(), fmt).time()
                except:
                    continue
            return time(0, 0)
        return val

    # ---------------- FETCH MAINTENANCE ----------------
    df = pd.read_sql("""
        SELECT *
        FROM Maintenance_Log
        WHERE Vehicle_ID = ?
    """, conn, params=(vehicle_id,))

    # ---------------- BD CALCULATION ----------------
    intervals = []

    for _, row in df.iterrows():

        start_time = to_time(row["Maint_Start_Time"])
        end_time = to_time(row["Maint_End_Time"])

        start = datetime.combine(row["Maint_Start_Date"], start_time)
        end = datetime.combine(row["Maint_End_Date"], end_time)

        day_start = datetime.combine(util_date, time(0, 0, 0))
        day_end = datetime.combine(util_date, time(23, 59, 59))

        if end < day_start or start > day_end:
            continue

        overlap_start = max(start, day_start)
        overlap_end = min(end, day_end)

        if overlap_start < overlap_end:
            intervals.append((overlap_start, overlap_end))

    # Merge overlapping intervals
    intervals.sort()
    merged = []

    for interval in intervals:
        if not merged or merged[-1][1] < interval[0]:
            merged.append(list(interval))
        else:
            merged[-1][1] = max(merged[-1][1], interval[1])

    bd_hours = sum(
        (end - start).total_seconds() / 3600
        for start, end in merged
    )

    bd = round(min(bd_hours, 24), 2)

    st.info(f"🛠️ Breakdown Hours (Auto): {bd} hrs")

    # ---------------- IDLE ----------------
    idle = round(24 - (worked + bd), 2)

    if idle < 0:
        idle = 0

    st.info(f"⏳ Idle Hours (Auto): {idle} hrs")

    # ---------------- VALIDATION ----------------
    errors = []

    if worked < 0:
        errors.append("Worked hours cannot be negative")

    if worked + bd > 24:
        errors.append("❌ Worked + Breakdown cannot exceed 24 hrs")

    total = worked + bd + idle
    if total > 24.01:
        errors.append("❌ Total hours exceed 24 hrs")

    # ✅ FIX: Duplicate check only after submit click
    if st.session_state.util_submit_clicked and not skip_duplicate_check:
        check = pd.read_sql("""
            SELECT * FROM Utilisation
            WHERE Vehicle_ID=? AND Util_Date=?
        """, conn, params=(vehicle_id, util_date))

        if not check.empty:
            errors.append("❌ Entry already exists for this vehicle & date")

    # Show errors
    if errors:
        for e in errors:
            st.error(e)

    # ---------------- SUBMIT BUTTON ----------------
    if st.button("Submit Utilisation"):
        st.session_state.util_submit_clicked = True

    # ---------------- SAVE ----------------
    if not errors and st.session_state.util_submit_clicked and confirm_action(
        "confirm_util_create",
        f"Save utilisation {util_id}?"
    ):
        try:
            cursor.execute("""
                INSERT INTO Utilisation
                (Util_ID, Util_Date, Vehicle_ID, Worked_hrs, Idle_hrs, BD_hrs)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                util_id,
                util_date,
                vehicle_id,
                worked,
                idle,
                bd
            ))

            conn.commit()

            # FLAGS RESET
            st.session_state.util_just_saved = True
            st.session_state.util_submit_clicked = False

            st.success(f"✅ Saved: {util_id}")
            st.rerun()

        except Exception as e:
            st.error(f"❌ Error: {e}")

    # ---------------- VIEW ----------------
    st.markdown("### 📋 Utilisation Records")

    df = pd.read_sql("SELECT * FROM Utilisation ORDER BY Util_ID DESC", conn)
    st.dataframe(df)
    # ======================================================
    # ======================================================
    # 🔁 COMMON UTILISATION ANALYTICS FUNCTION
    # ======================================================
    def show_utilisation_analytics(conn, key_prefix="util"):

        st.markdown("---")
        st.subheader("⏱️ Vehicle Utilisation Analytics")

        try:
            df_util = pd.read_sql("SELECT * FROM Utilisation", conn)
        except:
            df_util = pd.DataFrame()

        if not df_util.empty:

            df_util["Util_Date"] = pd.to_datetime(df_util["Util_Date"])

            # -------- FILTERS --------
            col1, col2 = st.columns(2)

            with col1:
                selected_vehicle = st.selectbox(
                    "Filter by Vehicle",
                    ["All"] + df_util["Vehicle_ID"].unique().tolist(),
                    key=f"{key_prefix}_vehicle"
                )

            with col2:
                selected_date = st.date_input(
                    "Filter by Date",
                    key=f"{key_prefix}_date"
                )

            df_filtered = df_util.copy()

            if selected_vehicle != "All":
                df_filtered = df_filtered[
                    df_filtered["Vehicle_ID"] == selected_vehicle
                    ]

            if selected_date:
                df_filtered = df_filtered[
                    df_filtered["Util_Date"] == pd.to_datetime(selected_date)
                    ]

            if df_filtered.empty:
                st.warning("No utilisation data available.")
            else:

                # -------- BAR CHART --------
                st.subheader("📊 Worked vs Idle vs Breakdown")
                df_bar = df_filtered[
                    ["Worked_hrs", "Idle_hrs", "BD_hrs"]
                ].sum()

                st.bar_chart(df_bar)

                # -------- LINE CHART --------
                st.subheader("📈 Daily Trend")
                df_line = df_filtered.sort_values("Util_Date")
                df_line = df_line.set_index("Util_Date")[
                    ["Worked_hrs", "Idle_hrs", "BD_hrs"]
                ]

                st.line_chart(df_line)

                # -------- SUMMARY --------
                st.subheader("📌 Summary")
                st.write(df_bar)

        else:
            st.info("No utilisation data available.")
    # ---------------- DELETE ----------------
    st.markdown("### ❌ Delete Utilisation")

    if not df.empty:
        selected_id = st.selectbox("Select Utilisation ID", df["Util_ID"])

        if confirm_action(
            "confirm_delete_util",
            f"Delete {selected_id}?"
        ):
            cursor.execute(
                "DELETE FROM Utilisation WHERE Util_ID=?",
                (selected_id,)
            )
            conn.commit()
            st.success("✅ Deleted successfully")
            st.rerun()

# ================== REPORTS PAGE ==================
def reports_page():
    st.title("📄 Reports")

    # ---------------- BACK BUTTON ----------------
    if st.button("⬅ Back"):
        navigate("home")

    conn = get_connection()

    st.markdown("## 📊 Fleet Management Reports")

    report_type = st.selectbox(
        "Select Report Type",
        [
            "🚗 List of Vehicles",
            "🧾 List of Purchase Orders",
            "⏱️ Utilisation by Vehicle & Date",
            "🛠️ Maintenance Log by Vehicle & Date Range",
            "💰 Cost by Vehicle ID From-To Date"   # ✅ ADDED
        ]
    )

    # ======================================================
    # 🚗 VEHICLES REPORT
    # ======================================================
    if report_type == "🚗 List of Vehicles":

        if st.button("Generate Vehicle Report"):
            df = pd.read_sql("SELECT * FROM Vehicle_Master", conn)

            st.subheader("🚗 Vehicle List Report")
            st.write(f"Total Vehicles: {len(df)}")
            st.dataframe(df, use_container_width=True)

    # ======================================================
    # 🧾 PURCHASE ORDER REPORT
    # ======================================================
    elif report_type == "🧾 List of Purchase Orders":

        if st.button("Generate PO Report"):
            df = pd.read_sql("""
                SELECT PO_ID, Vendor_Name, Material,
                       Quantity, Unit_Price, Total_Price, Order_Date
                FROM Purchase_Order
                ORDER BY Order_Date DESC
            """, conn)

            st.subheader("🧾 Purchase Order Report")
            st.write(f"Total Orders: {len(df)}")
            st.dataframe(df, use_container_width=True)

    # ======================================================
    # ⏱️ UTILISATION REPORT
    # ======================================================
    elif report_type == "⏱️ Utilisation by Vehicle & Date":

        df_veh = pd.read_sql("SELECT DISTINCT Vehicle_ID FROM Utilisation", conn)

        if df_veh.empty:
            st.warning("No utilisation data available.")
            return

        vehicle_id = st.selectbox("Select Vehicle", df_veh["Vehicle_ID"])

        from_date = st.date_input("From Date")
        to_date = st.date_input("To Date")

        if st.button("Generate Utilisation Report"):

            df = pd.read_sql("""
                SELECT *
                FROM Utilisation
                WHERE Vehicle_ID = ?
                AND Util_Date BETWEEN ? AND ?
                ORDER BY Util_Date
            """, conn, params=(vehicle_id, from_date, to_date))

            st.subheader(f"⏱️ Utilisation Report - {vehicle_id}")
            st.dataframe(df, use_container_width=True)

            if not df.empty:
                st.bar_chart(df.set_index("Util_Date")[["Worked_hrs", "Idle_hrs", "BD_hrs"]])

    # ======================================================
    # 🛠️ MAINTENANCE REPORT
    # ======================================================
    elif report_type == "🛠️ Maintenance Log by Vehicle & Date Range":

        df_veh = pd.read_sql("SELECT DISTINCT Vehicle_ID FROM Maintenance_Log", conn)

        if df_veh.empty:
            st.warning("No maintenance data available.")
            return

        vehicle_id = st.selectbox("Select Vehicle", df_veh["Vehicle_ID"])

        from_date = st.date_input("From Date", key="m_from")
        to_date = st.date_input("To Date", key="m_to")

        if st.button("Generate Maintenance Report"):

            df = pd.read_sql("""
                SELECT Main_ID, Vehicle_ID, Material,
                       Maintenance_Type, Description,
                       Quantity, Cost,
                       Maint_Start_Date, Maint_End_Date
                FROM Maintenance_Log
                WHERE Vehicle_ID = ?
                AND Maint_Start_Date BETWEEN ? AND ?
                ORDER BY Maint_Start_Date
            """, conn, params=(vehicle_id, from_date, to_date))

            st.subheader(f"🛠️ Maintenance Report - {vehicle_id}")
            st.dataframe(df, use_container_width=True)

            if not df.empty:
                total_cost = df["Cost"].sum()
                st.success(f"💰 Total Maintenance Cost for {vehicle_id}: ₹ {total_cost:,.2f}")

                st.bar_chart(df.groupby("Maint_Start_Date")["Cost"].sum())

    # ======================================================
    # 💰 COST REPORT (NEW - SEPARATE)
    # ======================================================
    elif report_type == "💰 Cost by Vehicle ID From-To Date":

        df_veh = pd.read_sql("SELECT DISTINCT Vehicle_ID FROM Maintenance_Log", conn)

        if df_veh.empty:
            st.warning("No maintenance data available.")
            return

        vehicle_id = st.selectbox("Select Vehicle", df_veh["Vehicle_ID"], key="c_vehicle")

        from_date = st.date_input("From Date", key="c_from")
        to_date = st.date_input("To Date", key="c_to")

        if st.button("Generate Cost Report"):

            df = pd.read_sql("""
                SELECT *
                FROM Maintenance_Log
                WHERE Vehicle_ID = ?
                AND Maint_Start_Date BETWEEN ? AND ?
            """, conn, params=(vehicle_id, from_date, to_date))

            st.subheader(f"💰 Cost Report - {vehicle_id}")

            if not df.empty:
                total_cost = df["Cost"].sum()
                st.success(f"Total Cost: ₹ {total_cost:,.2f}")
                st.dataframe(df, use_container_width=True)
            else:
                st.warning("No data found for selected range.")
# ================== ROUTING ==================

if not st.session_state.logged_in:
    login_page()

else:
    if st.session_state.page == "home":
        home()

    elif st.session_state.page == "dashboard":
        dashboard_analytics()

    elif st.session_state.page == "material":
        material_page()

    elif st.session_state.page == "po":
        po_page()

    elif st.session_state.page == "vehicle":
        vehicle_page()

    elif st.session_state.page == "maintenance":
        maintenance_page()

    elif st.session_state.page == "utilisation":
        utilisation_page()

    elif st.session_state.page == "reports":
        reports_page()

