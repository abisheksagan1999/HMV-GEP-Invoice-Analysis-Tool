import streamlit as st
from PIL import Image
import pandas as pd
import re
from fuzzywuzzy import fuzz
from io import BytesIO
import time
import difflib

# Streamlit page config
st.set_page_config(page_title="HMV Fair Quote Tool", layout="wide")

# Load logos and align them side-by-side
logo_col1, logo_col2, logo_col3 = st.columns([2, 6, 2])
with logo_col1:
    st.image("logo1.png", width=200)
with logo_col3:
    st.image("logo2.png", width=200)

# Title
st.markdown("""
    <h2 style='text-align:center;'>HMV Fair Quote Validation Tool</h2>
    <hr>
""", unsafe_allow_html=True)

# Custom CSS for styling
st.markdown("""
    <style>
    .conclusion-box {
        padding: 1.5em;
        border-radius: 10px;
        margin: 1.5em 0;
        border-left: 5px solid;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .exact-match { border-color: #00cc00; background-color: #f0fff4; }
    .approx-match { border-color: #ffcc00; background-color: #fffaf0; }
    .closest-match { border-color: #ff6666; background-color: #fff5f5; }
    .loading-overlay {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.7);
        z-index: 9999;
        display: flex;
        justify-content: center;
        align-items: center;
        flex-direction: column;
        color: white;
        font-size: 24px;
    }
    .spinner {
        border: 8px solid #f3f3f3;
        border-top: 8px solid #3498db;
        border-radius: 50%;
        width: 60px;
        height: 60px;
        animation: spin 2s linear infinite;
        margin-bottom: 20px;
    }
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    </style>
""", unsafe_allow_html=True)

# Loading overlay function
def show_loading(message):
    st.markdown(f"""
        <div class="loading-overlay">
            <div class="spinner"></div>
            <div>{message}</div>
        </div>
    """, unsafe_allow_html=True)
    st.experimental_rerun()

# Upload file
uploaded_file = st.file_uploader("Upload HMV Excel File (hmv_data.xlsx format):", type=["xlsx"])

if uploaded_file:
    # Show loading overlay during file processing
    show_loading("Processing your file...")
    
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

    # Remove loading overlay after processing
    st.markdown("<style>.loading-overlay{display:none;}</style>", unsafe_allow_html=True)

    st.sidebar.header("Filters")
    all_years = sorted(df['Year'].dropna().unique())
    year_filter = st.sidebar.multiselect("Select Year(s):", all_years, default=all_years)
    card_filter = st.sidebar.text_input("Card Number (partial):")
    min_hr, max_hr = st.sidebar.slider("Hour Range", 0, int(df['Total Hours'].max()), (0, int(df['Total Hours'].max())))

    filtered_df = df[df['Year'].isin(year_filter)]
    if card_filter:
        filtered_df = filtered_df[filtered_df['Orig. Card #'].astype(str).str.contains(card_filter, case=False)]
    filtered_df = filtered_df[(filtered_df['Total Hours'] >= min_hr) & (filtered_df['Total Hours'] <= max_hr)]

    st.write("### Enter Description and Corrective Action")
    with st.form("form"):
        # Changed Discrepancy to Description of Non-Routine
        discrepancy_input = st.text_area("Description of Non-Routine", height=80)
        corrective_input = st.text_area("Corrective Action", height=80)
        supplier_hours = st.number_input("Supplier Quoted Hours", min_value=0.0, step=0.1)
        submit = st.form_submit_button("Search")

    if submit and discrepancy_input and corrective_input:
        # Show loading overlay during search processing
        show_loading("Searching for matches...")
        
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
            if supplier < fair:
                return ("Fair quote - approve quote", "green")
            elif abs(supplier - fair) / fair <= 0.05:
                return ("In expected range (±5%) - consider approving", "black")
            else:
                return ("Beyond expected range - needs to be reviewed with BP", "red")
        
        # Remove loading overlay after processing
        st.markdown("<style>.loading-overlay{display:none;}</style>", unsafe_allow_html=True)

        # Always show conclusion in a consistent box at the top
        conclusion_box = None
        
        if not exact.empty:
            row = exact.iloc[0]
            conclusion, color = get_conclusion(supplier_hours, row['Fair Quote (hrs)'])
            conclusion_box = st.markdown(f"""
                <div class="conclusion-box exact-match">
                    <h3 style='text-align:center; margin:0; color:{color};'>Conclusion: {conclusion}</h3>
                    <p style='text-align:center; margin:0;'>Match Type: <b>Exact Match</b> | Fair Quote: {row['Fair Quote (hrs)']:.2f} hrs</p>
                </div>
            """, unsafe_allow_html=True)
            st.success("### ✅ Exact Match Found")
            st.dataframe(exact[['Description', 'Corrective Action', 'Actual Historic Hours', 'Fair Quote (hrs)', 'Occurrences']].style.set_properties(**{'white-space': 'pre-wrap'}), use_container_width=True)

        if not top2.empty:
            row = top2.iloc[0]
            if conclusion_box is None:
                conclusion, color = get_conclusion(supplier_hours, row['Fair Quote (hrs)'])
                conclusion_box = st.markdown(f"""
                    <div class="conclusion-box approx-match">
                        <h3 style='text-align:center; margin:0; color:{color};'>Conclusion: {conclusion}</h3>
                        <p style='text-align:center; margin:0;'>Match Type: <b>Approximate Match</b> | Fair Quote: {row['Fair Quote (hrs)']:.2f} hrs</p>
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
            
            # Create HTML table without conclusion column
            html_table = pd.DataFrame(rows).to_html(escape=False, index=False)
            st.markdown(html_table, unsafe_allow_html=True)

        if exact.empty and top2.empty and not closest.empty:
            row = closest.iloc[0]
            if conclusion_box is None:
                conclusion, color = get_conclusion(supplier_hours, row['Fair Quote (hrs)'])
                conclusion_box = st.markdown(f"""
                    <div class="conclusion-box closest-match">
                        <h3 style='text-align:center; margin:0; color:{color};'>Conclusion: {conclusion}</h3>
                        <p style='text-align:center; margin:0;'>Match Type: <b>Nearest Reference</b> | Fair Quote: {row['Fair Quote (hrs)']:.2f} hrs</p>
                    </div>
                """, unsafe_allow_html=True)
            
            st.warning("### 📝 No close matches found — showing nearest reference")
            ref_row = {
                'Description': row['Normalized Discrepancy'],
                'Corrective Action': row['Normalized Corrective Action'],
                'Actual Historic Hours': f"{row['Actual Historic Hours']:.2f}",
                'Fair Quote (hrs)': f"{row['Fair Quote (hrs)']:.2f}",
                'Occurrences': row['Occurrences'],
                'Overlap %': f"{row['Overlap']:.1f}%"
            }
            html_ref = pd.DataFrame([ref_row]).to_html(escape=False, index=False)
            st.markdown(html_ref, unsafe_allow_html=True)

else:
    st.info("Please upload the 'hmv_data.xlsx' file to begin.")