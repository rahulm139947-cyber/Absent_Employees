import pandas as pd
import streamlit as st
import io

# --- Page Setup ---
st.set_page_config(page_title="Absentee Report Generator", layout="centered")

# --- Custom Styling ---
st.markdown("""
    <style>
        .main-title {
            font-size: 32px;
            font-weight: bold;
            color: #2E86C1;
            text-align: center;
            margin-bottom: 10px;
        }
        .subtitle {
            font-size: 18px;
            color: #555;
            text-align: center;
            margin-bottom: 30px;
        }
        .footer {
            font-size: 13px;
            color: gray;
            text-align: center;
            margin-top: 50px;
        }
        .stButton>button {
            background-color: #2E86C1;
            color: white;
            font-weight: bold;
        }
    </style>
""", unsafe_allow_html=True)

# --- Logo (local or URL) ---
st.image("your_logo.png", width=150)  # Replace with your actual logo file or URL

# --- Title and Subtitle ---
st.markdown('<div class="main-title">Absentee Report Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload your attendance file to generate a clean absentee list</div>', unsafe_allow_html=True)

# --- File Upload ---
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")

    try:
        # Step 1: Load Excel
        st.write("📄 Reading Excel file...")
        df = pd.read_excel(uploaded_file, header=2)

        # Step 2: Normalize column names
        st.write("🧹 Cleaning column names...")
        df.columns = df.columns.str.replace('\n', ' ').str.strip().str.lower()

        # Step 3: Identify columns by position
        name_column = df.columns[1]      # Column B
        date_column = df.columns[8]      # Column I
        status_column = df.columns[13]   # Column N
        nationality_column = df.columns[43]  # Column AR

        st.write("🔍 Filtering for 'NOT IN' status and excluding SAUDI/KOREA...")

        # Step 4: Apply filters
        filtered_df = df[
            (df[status_column].astype(str).str.strip().str.upper() == 'NOT IN') &
            (~df[nationality_column].astype(str).str.upper().isin(['SAUDI', 'KOREA']))
        ]

        # Step 5: Prepare output
        output_df = filtered_df[[date_column, name_column]].copy()
        output_df.columns = ['Date', 'Absent Name']

        st.write(f"📊 Found **{len(output_df)}** absentees.")

        # Step 6: Convert to Excel in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False)

        # Step 7: Download button
        st.download_button(
            label="📥 Download Absentee Report",
            data=output.getvalue(),
            file_name="absent_by_date.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")

# --- Footer ---
st.markdown('<div class="footer">© 2025 Your Company • Contact: support@yourcompany.com</div>', unsafe_allow_html=True)

