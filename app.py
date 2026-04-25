import streamlit as st
import tempfile
import os
import time
from bootstrap_dematel import run_pipeline

# ===== UI =====
st.set_page_config(page_title="DEMATEL Tool", layout="wide")

st.title("📊 Bootstrap Z-Fuzzy DEMATEL Tool")

# Upload file
uploaded_file = st.file_uploader("📂 Upload Excel file", type=["xlsx"])

# ===== CONFIG =====
st.sidebar.header("⚙️ Configuration")

B = st.sidebar.number_input("Bootstrap samples (B)", min_value=100, value=2000)
alpha = st.sidebar.slider("Alpha (CI)", 0.01, 0.2, 0.05)
seed = st.sidebar.number_input("Random seed", value=80)

output_name = st.sidebar.text_input("Output name", value="result")

# ===== RUN BUTTON =====
if st.button("🚀 Run Analysis"):

    if uploaded_file is None:
        st.error("❌ Please upload Excel file")
    else:
        # Save temp input file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            input_path = tmp.name

        # Create output folder
        os.makedirs("outputs", exist_ok=True)

        ts = int(time.time())
        output_xls = f"outputs/{output_name}_{ts}.xlsx"
        output_img = f"outputs/{output_name}_{ts}.png"

        # ===== RUN =====
        with st.spinner("⏳ Running Bootstrap... (có thể mất vài phút)"):
            df = run_pipeline(
                input_path,
                output_xls,
                output_img,
                B=B,
                alpha=alpha,
                seed=seed
            )

        st.success("✅ Done!")

        # ===== SHOW RESULT =====
        st.subheader("📋 Result Table")
        st.dataframe(df, use_container_width=True)

        st.subheader("🖼️ IRM Visualization")
        st.image(output_img)

        # ===== DOWNLOAD =====
        with open(output_xls, "rb") as f:
            st.download_button("⬇️ Download Excel", f, file_name=os.path.basename(output_xls))

        with open(output_img, "rb") as f:
            st.download_button("⬇️ Download Image", f, file_name=os.path.basename(output_img))