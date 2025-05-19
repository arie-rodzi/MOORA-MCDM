import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="MOORA Method App", layout="wide")
st.title("MOORA (Multi-Objective Optimization on the basis of Ratio Analysis)")

st.markdown("Upload an Excel file with three sheets:")
st.markdown("1. `DecisionMatrix` — alternatives × criteria values")
st.markdown("2. `Weights` — a single row with weights (must sum to 1)")
st.markdown("3. `Types` — a single row of 'benefit' or 'cost' labels for each criterion")

uploaded_file = st.file_uploader("📤 Upload Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Read input sheets
        dm = pd.read_excel(uploaded_file, sheet_name="DecisionMatrix", index_col=0)
        weights_df = pd.read_excel(uploaded_file, sheet_name="Weights", header=0)
        types_df = pd.read_excel(uploaded_file, sheet_name="Types", header=0)

        weights = pd.to_numeric(weights_df.values[0], errors='coerce')
        types = [str(x).strip().lower() for x in types_df.values[0]]

        # Check for numeric weights
        if np.isnan(weights).any():
            st.error("❌ One or more weights are non-numeric or missing.")
            st.stop()

        # Validate types
        valid_types = {"benefit", "cost"}
        if any(t not in valid_types for t in types):
            st.error("❌ One or more criteria types are not 'benefit' or 'cost'.")
            st.stop()

        # Check length match
        if len(weights) != dm.shape[1] or len(types) != dm.shape[1]:
            st.error("❌ Number of weights/types must match the number of criteria.")
            st.stop()

        # Check if weights sum to 1
        if not np.isclose(np.sum(weights), 1.0):
            st.error("❌ Weights must sum to 1. Current sum: {:.4f}".format(np.sum(weights)))
            st.stop()

        # Step 1: Normalize
        norm_matrix = dm / np.sqrt((dm**2).sum())

        # Step 2: Weighted normalized matrix
        weighted_norm = norm_matrix * weights

        # Step 3: MOORA Score
        benefit_idx = [i for i, t in enumerate(types) if t == 'benefit']
        cost_idx = [i for i, t in enumerate(types) if t == 'cost']

        benefit_scores = weighted_norm.iloc[:, benefit_idx].apply(pd.to_numeric, errors='coerce')
        cost_scores = weighted_norm.iloc[:, cost_idx].apply(pd.to_numeric, errors='coerce')

        scores = benefit_scores.sum(axis=1) - cost_scores.sum(axis=1)

        result_df = pd.DataFrame({
            "Score": scores,
            "Rank": scores.rank(ascending=False).astype(int)
        }, index=dm.index)

        st.success("✅ MOORA Calculation Completed!")
        st.dataframe(result_df.sort_values("Rank"))

        # Export results to Excel
        def to_excel():
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                dm.to_excel(writer, sheet_name="Input_Data")
                pd.DataFrame([weights], columns=dm.columns).to_excel(writer, sheet_name="Weights", index=False)
                pd.DataFrame([types], columns=dm.columns).to_excel(writer, sheet_name="Types", index=False)
                norm_matrix.to_excel(writer, sheet_name="Normalized")
                weighted_norm.to_excel(writer, sheet_name="Weighted")
                result_df.to_excel(writer, sheet_name="MOORA_Result")
            output.seek(0)
            return output

        excel_data = to_excel()
        st.download_button("📥 Download Results as Excel", data=excel_data,
                           file_name="moora_results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ Error: {e}")