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
        weights = pd.read_excel(uploaded_file, sheet_name="Weights", header=None).values[0]
        types = pd.read_excel(uploaded_file, sheet_name="Types", header=None).values[0]

        # Validation
        assert len(weights) == dm.shape[1], "Mismatch: Weights and Criteria"
        assert len(types) == dm.shape[1], "Mismatch: Types and Criteria"
        assert np.isclose(sum(weights), 1.0), "Weights must sum to 1"

        # Step 1: Normalize
        norm_matrix = dm / np.sqrt((dm**2).sum())

        # Step 2: Weighted normalized
        weighted_norm = norm_matrix * weights

        # Step 3: MOORA Score
        benefit_idx = [i for i, t in enumerate(types) if t.lower() == 'benefit']
        cost_idx = [i for i, t in enumerate(types) if t.lower() == 'cost']

        scores = weighted_norm.iloc[:, benefit_idx].sum(axis=1) - weighted_norm.iloc[:, cost_idx].sum(axis=1)

        result_df = pd.DataFrame({
            "Score": scores,
            "Rank": scores.rank(ascending=False).astype(int)
        }, index=dm.index)

        st.success("MOORA Calculation Completed!")
        st.dataframe(result_df.sort_values("Rank"))

        # Export as Excel
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