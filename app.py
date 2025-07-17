import streamlit as st
from PIL import Image
import pandas as pd
import re
from fuzzywuzzy import fuzz
from io import BytesIO
import time
import difflib
import numpy as np

# Streamlit page config
st.set_page_config(page_title="HMV Fair Quote Tool", layout="wide", page_icon="🔧")

# Load logos and align them side-by-side
logo_col1, logo_col2, logo_col3 = st.columns([2, 6, 2])
with logo_col1:
    st.image("logo1.png", width=200)
with logo_col3:
    st.image("logo2.png", width=200)

# Title
st.markdown("""
    <h2 style='text-align:center;'>HMV Fair Quote Validation Tool</h2>
    <hr style='border:2px solid #3498db; border-radius:5px;'>
""", unsafe_allow_html=True)

# Custom CSS for styling
st.markdown("""
    <style>
    :root {
        --primary: #3498db;
        --success: #2ecc71;
        --warning: #f39c12;
        --danger: #e74c3c;
        --dark: #2c3e50;
    }
    
    .conclusion-box {
        padding: 1.5em;
        border-radius: 12px;
        margin: 2em 0;
        border-left: 6px solid;
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        background: linear-gradient(to right, #f8f9fa, #ffffff);
        transition: all 0.3s ease;
    }
    .conclusion-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 10px 20px rgba(0,0,0,0.2);
    }
    .exact-match { border-color: var(--success); }
    .approx-match { border-color: var(--warning); }
    .closest-match { border-color: var(--danger); }
    
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 1.5em;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        text-align: center;
        margin: 1em 0;
        border-top: 4px solid var(--primary);
    }
    
    .metric-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: var(--dark);
    }
    
    .metric-label {
        font-size: 1.1rem;
        color: #7f8c8d;
    }
    
    .diff-positive { color: var(--danger); }
    .diff-negative { color: var(--success); }
    .diff-neutral { color: var(--dark); }
    
    .loading-overlay {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.85);
        z-index: 9999;
        display: flex;
        justify-content: center;
        align-items: center;
        flex-direction: column;
        color: white;
        font-size: 28px;
    }
    
    .spinner {
        border: 10px solid rgba(255,255,255,0.3);
        border-top: 10px solid var(--primary);
        border-radius: 50%;
        width: 80px;
        height: 80px;
        animation: spin 1.5s linear infinite;
        margin-bottom: 30px;
    }
    
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    
    .pulse {
        animation: pulse 2s infinite;
    }
    
    @keyframes pulse {
        0% { transform: scale(1); }
        50% { transform: scale(1.05); }
        100% { transform: scale(1); }
    }
    
    .form-container {
        background: #f8f9fa;
        padding: 2em;
        border-radius: 15px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }
    
    .result-table {
        border-collapse: collapse;
        width: 100%;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    
    .result-table th {
        background: var(--primary);
        color: white;
        padding: 12px 15px;
        text-align: left;
        font-weight: 600;
    }
    
    .result-table td {
        padding: 12px 15px;
        border-bottom: 1px solid #e0e0e0;
    }
    
    .result-table tr:nth-child(even) {
        background-color: #f5f7fa;
    }
    
    .result-table tr:hover {
        background-color: #ebf5fb;
    }
    
    .highlight {
        background-color: #fffacd !important;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# Create a placeholder for loading overlay
loading_placeholder = st.empty()

def show_loading(message):
    loading_placeholder.markdown(f"""
        <div class="loading-overlay">
            <div class="spinner"></div>
            <div class='pulse'>{message}</div>
        </div>
    """, unsafe_allow_html=True)

def hide_loading():
    loading_placeholder.empty()

# Upload file
uploaded_file = st.file_uploader("Upload HMV Excel File (hmv_data.xlsx format):", type=["xlsx"])

if uploaded_file:
    show_loading("🔍 Processing your file...")
    
    df = pd.read_excel(uploaded_file)

    def normalize_text(text):
        if pd.isna(text): return ""
        text = str(text).upper()
        text = re.sub(r'\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b', '', text)
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    df['Normalized Corrective Action'] = df['Corrective Action'].apply(normalize_text)
    df['Normalized Discrepancy'] = df['Description'].apply(lambda x: normalize_text(x.replace("(FOR REFERENCE ONLY)", "")))
    df['Combined Key'] = df['Normalized Discrepancy'] + " | " + df['Normalized Corrective Action']

    clusters = {}
    for key in df['Combined Key'].unique():
        if not key: continue
        for rep in clusters:
            if fuzz.token_set_ratio(key, rep) >= 90:
                clusters[rep].append(key)
                break
        else:
            clusters[key] = [key]

    key_to_rep = {k: r for r, lst in clusters.items() for k in lst}
    df['Cluster Key'] = df['Combined Key'].map(key_to_rep)

    # Average total hours for clusters
    hours = df.groupby('Cluster Key')['Total Hours'].agg(['mean', 'count']).reset_index()
    hours.columns = ['Cluster Key', 'Actual Historic Hours', 'Occurrences']
    df = df.merge(hours, on='Cluster Key', how='left')
    df['Fair Quote (hrs)'] = df['Actual Historic Hours'].round(2)

    hide_loading()

    st.sidebar.header("🔍 Filters")
    all_years = sorted(df['Year'].dropna().unique())
    year_filter = st.sidebar.multiselect("Select Year(s):", all_years, default=all_years)
    card_filter = st.sidebar.text_input("Card Number (partial):")
    
    max_val = int(df['Total Hours'].max()) if df['Total Hours'].max() > 0 else 0
    min_hr, max_hr = st.sidebar.slider(
        "Hour Range", 
        0, 
        max_val, 
        (0, max_val)
    )

    filtered_df = df[df['Year'].isin(year_filter)]
    if card_filter:
        filtered_df = filtered_df[filtered_df['Orig. Card #'].astype(str).str.contains(card_filter, case=False)]
    filtered_df = filtered_df[(filtered_df['Total Hours'] >= min_hr) & (filtered_df['Total Hours'] <= max_hr)]

    st.markdown("### 📝 Enter Maintenance Details")
    with st.form("form", clear_on_submit=False):
        with st.container():
            col1, col2 = st.columns(2)
            with col1:
                # Changed Discrepancy to Description of Non-Routine
                discrepancy_input = st.text_area("Description of Non-Routine", height=120,
                                               placeholder="Describe the issue or discrepancy...")
            with col2:
                corrective_input = st.text_area("Corrective Action", height=120,
                                              placeholder="Describe the corrective action taken...")
            
            supplier_hours = st.number_input("Supplier Quoted Hours", min_value=0.0, step=0.1)
            
            submit = st.form_submit_button("🔍 Analyze Historical Invoices", use_container_width=True)

    if submit and discrepancy_input and corrective_input:
        show_loading("📊 Analyzing historical invoices...")
        
        norm_disc = normalize_text(discrepancy_input.replace("(FOR REFERENCE ONLY)", ""))
        norm_corr = normalize_text(corrective_input)
        combined_input = norm_disc + " | " + norm_corr

        exact = df[df['Combined Key'] == combined_input]

        def semantic_overlap(a, b):
            matcher = difflib.SequenceMatcher(None, a.split(), b.split())
            return matcher.ratio() * 100

        def total_similarity(row):
            d_ov = semantic_overlap(norm_disc, row['Normalized Discrepancy'])
            c_ov = semantic_overlap(norm_corr, row['Normalized Corrective Action'])
            return (d_ov + c_ov) / 2

        df['Overlap'] = df.apply(total_similarity, axis=1)
        approx = df[(df['Overlap'] >= 50) & (df['Combined Key'] != combined_input)]
        top2 = approx.sort_values(by='Overlap', ascending=False).head(2)
        closest = df[df['Overlap'] < 50].sort_values(by='Overlap', ascending=False).head(1)

        def get_conclusion(supplier, fair):
            # Calculate percentage difference
            if fair == 0:
                percent_diff = "N/A"
                diff_class = "diff-neutral"
            else:
                percent_diff = ((supplier - fair) / fair) * 100
                if percent_diff < 0:
                    diff_class = "diff-negative"
                elif abs(percent_diff) <= 5:
                    diff_class = "diff-neutral"
                else:
                    diff_class = "diff-positive"
            
            # Format percentage display
            if fair == 0:
                percent_display = "N/A (no historical data)"
            else:
                sign = "+" if percent_diff >= 0 else ""
                percent_display = f"{sign}{percent_diff:.1f}%"
            
            # Generate conclusion text
            if fair == 0:
                return ("No historical data available - needs manual review", "red", percent_display, diff_class)
            if supplier < fair:
                return ("Fair quote - approve quote", "green", percent_display, diff_class)
            elif abs(supplier - fair) / fair <= 0.05:
                return ("In expected range (±5%) - consider approving", "black", percent_display, diff_class)
            else:
                return ("Beyond expected range - needs BP review", "red", percent_display, diff_class)
        
        hide_loading()

        # Always show conclusion in a consistent box at the top
        conclusion_box = None
        
        if not exact.empty:
            row = exact.iloc[0]
            conclusion, color, percent_diff, diff_class = get_conclusion(supplier_hours, row['Fair Quote (hrs)'])
            
            # Create metrics cards
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{row['Fair Quote (hrs)']:.2f}</div>
                        <div class="metric-label">Fair Quote based on historical data</div>
                    </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{supplier_hours:.2f}</div>
                        <div class="metric-label">Supplier Quoted Hours</div>
                    </div>
                """, unsafe_allow_html=True)
            with col3:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value {diff_class}">{percent_diff}</div>
                        <div class="metric-label">Percentage Difference</div>
                    </div>
                """, unsafe_allow_html=True)
            
            # Enhanced conclusion box
            conclusion_box = st.markdown(f"""
                <div class="conclusion-box exact-match">
                    <div style='display:flex; justify-content:space-between; align-items:center;'>
                        <div>
                            <h3 style='margin:0; color:{color};'>Conclusion: {conclusion}</h3>
                            <p style='margin:0; font-size:1.1rem;'>Match Type: <b>Exact Match</b></p>
                        </div>
                        <div style='text-align:right;'>
                            <p style='margin:0; font-size:1.1rem;'>Historical Occurrences: <b>{row['Occurrences']}</b></p>
                            <p style='margin:0; font-size:1.1rem;'>Average Hours: <b>{row['Actual Historic Hours']:.2f}</b></p>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            st.success("### ✅ Exact Match Found")
            st.dataframe(exact[['Description', 'Corrective Action', 'Actual Historic Hours', 'Fair Quote (hrs)', 'Occurrences']].style.set_properties(**{'white-space': 'pre-wrap'}), use_container_width=True)

        elif not top2.empty:
            row = top2.iloc[0]
            conclusion, color, percent_diff, diff_class = get_conclusion(supplier_hours, row['Fair Quote (hrs)'])
            
            # Create metrics cards
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{row['Fair Quote (hrs)']:.2f}</div>
                        <div class="metric-label">Fair Quote based on historical data</div>
                    </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{supplier_hours:.2f}</div>
                        <div class="metric-label">Supplier Quoted Hours</div>
                    </div>
                """, unsafe_allow_html=True)
            with col3:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value {diff_class}">{percent_diff}</div>
                        <div class="metric-label">Percentage Difference</div>
                    </div>
                """, unsafe_allow_html=True)
            
            # Enhanced conclusion box
            conclusion_box = st.markdown(f"""
                <div class="conclusion-box approx-match">
                    <div style='display:flex; justify-content:space-between; align-items:center;'>
                        <div>
                            <h3 style='margin:0; color:{color};'>Conclusion: {conclusion}</h3>
                            <p style='margin:0; font-size:1.1rem;'>Match Type: <b>Approximate Match</b></p>
                        </div>
                        <div style='text-align:right;'>
                            <p style='margin:0; font-size:1.1rem;'>Historical Occurrences: <b>{row['Occurrences']}</b></p>
                            <p style='margin:0; font-size:1.1rem;'>Average Hours: <b>{row['Actual Historic Hours']:.2f}</b></p>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            st.info("### 🔍 Approximate Matches (Top 2)")

            def highlight_diff(text, ref):
                ref_words = set(ref.split())
                return " ".join([f"<b><span style='color:red'>{w}</span></b>" if w not in ref_words else w for w in text.split()])

            rows = []
            for _, row in top2.iterrows():
                rows.append({
                    'Description': highlight_diff(row['Normalized Discrepancy'], norm_disc),
                    'Corrective Action': highlight_diff(row['Normalized Corrective Action'], norm_corr),
                    'Actual Historic Hours': f"{row['Actual Historic Hours']:.2f}",
                    'Fair Quote (hrs)': f"{row['Fair Quote (hrs)']:.2f}",
                    'Occurrences': row['Occurrences'],
                    'Overlap %': f"{row['Overlap']:.1f}%"
                })
            
            # Create HTML table
            html_table = f"""
            <table class="result-table">
                <thead>
                    <tr>
                        <th>Description</th>
                        <th>Corrective Action</th>
                        <th>Historic Hours</th>
                        <th>Fair Quote (hrs)</th>
                        <th>Occurrences</th>
                        <th>Overlap %</th>
                    </tr>
                </thead>
                <tbody>
            """
            
            for row in rows:
                html_table += f"""
                <tr>
                    <td>{row['Description']}</td>
                    <td>{row['Corrective Action']}</td>
                    <td>{row['Actual Historic Hours']}</td>
                    <td>{row['Fair Quote (hrs)']}</td>
                    <td>{row['Occurrences']}</td>
                    <td>{row['Overlap %']}</td>
                </tr>
                """
            
            html_table += "</tbody></table>"
            st.markdown(html_table, unsafe_allow_html=True)

        elif not closest.empty:
            row = closest.iloc[0]
            conclusion, color, percent_diff, diff_class = get_conclusion(supplier_hours, row['Fair Quote (hrs)'])
            
            # Create metrics cards
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{row['Fair Quote (hrs)']:.2f}</div>
                        <div class="metric-label">Fair Quote based on historical data</div>
                    </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value">{supplier_hours:.2f}</div>
                        <div class="metric-label">Supplier Quoted Hours</div>
                    </div>
                """, unsafe_allow_html=True)
            with col3:
                st.markdown(f"""
                    <div class="metric-card">
                        <div class="metric-value {diff_class}">{percent_diff}</div>
                        <div class="metric-label">Percentage Difference</div>
                    </div>
                """, unsafe_allow_html=True)
            
            # Enhanced conclusion box
            conclusion_box = st.markdown(f"""
                <div class="conclusion-box closest-match">
                    <div style='display:flex; justify-content:space-between; align-items:center;'>
                        <div>
                            <h3 style='margin:0; color:{color};'>Conclusion: {conclusion}</h3>
                            <p style='margin:0; font-size:1.1rem;'>Match Type: <b>Nearest Reference</b></p>
                        </div>
                        <div style='text-align:right;'>
                            <p style='margin:0; font-size:1.1rem;'>Historical Occurrences: <b>{row['Occurrences']}</b></p>
                            <p style='margin:0; font-size:1.1rem;'>Average Hours: <b>{row['Actual Historic Hours']:.2f}</b></p>
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            st.warning("### 📝 No close matches found — showing nearest reference")
            html_ref = f"""
            <table class="result-table">
                <thead>
                    <tr>
                        <th>Description</th>
                        <th>Corrective Action</th>
                        <th>Historic Hours</th>
                        <th>Fair Quote (hrs)</th>
                        <th>Occurrences</th>
                        <th>Overlap %</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>{row['Normalized Discrepancy']}</td>
                        <td>{row['Normalized Corrective Action']}</td>
                        <td>{row['Actual Historic Hours']:.2f}</td>
                        <td>{row['Fair Quote (hrs)']:.2f}</td>
                        <td>{row['Occurrences']}</td>
                        <td>{row['Overlap']:.1f}%</td>
                    </tr>
                </tbody>
            </table>
            """
            st.markdown(html_ref, unsafe_allow_html=True)

else:
    st.info("Please upload the 'hmv_data.xlsx' file to begin.")