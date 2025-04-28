
import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="ME2L Follow-Up Mail Generator", layout="wide")
st.title("📧 ME2L Follow-Up Mail Generator with Delay → Vendor → PO Flow")

uploaded_file = st.file_uploader("📤 Upload ME2L Excel File", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df['Document Date'] = pd.to_datetime(df['Document Date'], errors='coerce')
        df['Still to be delivered (qty)'] = pd.to_numeric(df['Still to be delivered (qty)'], errors='coerce')
        df['Purchasing Document'] = df['Purchasing Document'].astype(str)

        today = pd.to_datetime(datetime.today())
        df['Delay (Days)'] = (today - df['Document Date']).dt.days

        # Step 1: Delay Slider
        st.sidebar.header("⏱️ Step 1: Set Minimum Delay")
        max_delay = int(df['Delay (Days)'].max())
        selected_delay = st.sidebar.slider("Show POs delayed more than (Days)", 0, max_delay, 30)

        delayed_df = df[df['Delay (Days)'] >= selected_delay]

        # 🔢 Summary
        st.subheader("📊 Summary for Selected Delay Filter")
        num_vendors = delayed_df['Name of Supplier'].nunique()
        num_pos = delayed_df['Purchasing Document'].nunique()
        st.markdown(f"**Total Delayed Vendors:** {num_vendors}")
        st.markdown(f"**Total Delayed POs:** {num_pos}")

        # Step 2: Vendor Selection
        st.sidebar.header("🏢 Step 2: Select Vendor")
        vendor_options = sorted(delayed_df['Name of Supplier'].dropna().unique())
        selected_vendor = st.sidebar.selectbox("Select Supplier", vendor_options if vendor_options else ["No matching vendors"])

        # Step 3: PO List based on vendor
        vendor_df = delayed_df[delayed_df['Name of Supplier'] == selected_vendor]
        po_list = sorted(vendor_df['Purchasing Document'].unique())
        st.sidebar.header("📄 Step 3: Select PO")
        selected_po = st.sidebar.selectbox("Select PO Number", po_list if po_list else ["No POs for selected vendor"])

        # Display mail content
        final_df = vendor_df[vendor_df['Purchasing Document'] == selected_po]

        if not final_df.empty:
            doc_date = final_df['Document Date'].iloc[0].strftime('%d-%b-%Y')
            delay_days = int(final_df['Delay (Days)'].max())
            items = final_df[['Short Text', 'Still to be delivered (qty)', 'Order Unit']].dropna()
            item_lines = [f"- {row['Short Text']} — Pending: {int(row['Still to be delivered (qty)'])} {row['Order Unit']}" for idx, row in items.iterrows()]
            items_text = "\n".join(item_lines)

            mail_text = f"""
Subject: Follow-up on Outstanding Delivery for PO {selected_po}

Dear Supplier ({selected_vendor}),

We wish to bring to your attention that the following items under Purchase Order **{selected_po}**, dated **{doc_date}**, are still pending delivery:

{items_text}

It has been over **{delay_days} days** since the PO was issued. We kindly request your immediate attention to expedite the dispatch and share a revised delivery schedule.

Please treat this as a priority.

Appreciate your cooperation.

Best regards,  
Ramkumar  
Sr. GM – Procurement  
Thermopads Pvt. Ltd.
"""
            st.subheader("📬 Generated Follow-Up Mail")
            st.code(mail_text, language='markdown')
        else:
            st.warning("⚠️ No matching data for selected vendor and PO.")

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("👆 Please upload your ME2L Excel file.")
