import streamlit as st
import pandas as pd
from io import BytesIO

# ----------------------------------------------------
# Efficient combination finder (NO sorting)
# ----------------------------------------------------
def find_combinations(df, value_cols, weight_cols, weight_factor, target_length):
    results = []

    def recurse(idx, current_rows, current_length):
        if current_length > target_length:
            return
        if current_length == target_length:
            subset = df.loc[current_rows]

            unweighted_sum = subset[value_cols].sum().sum()
            weighted_sum = unweighted_sum
            if weight_cols:
                weighted_sum += subset[weight_cols].sum().sum() * (weight_factor - 1)

            results.append({
                "Indices": current_rows.copy(),
                "Subset": subset,
                "Unweighted": unweighted_sum,
                "Weighted": weighted_sum
            })
            return

        for i in range(idx, len(df)):
            recurse(i + 1,
                    current_rows + [df.index[i]],
                    current_length + df.iloc[i]["Length"])

    recurse(0, [], 0)
    return results


# ----------------------------------------------------
# Streamlit App
# ----------------------------------------------------
st.title("Combination Finder (Performance Optimized)")

uploaded = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded:
    df = pd.read_excel(uploaded)
    df.columns = df.columns.str.replace("\r", "").str.replace("\n", "").str.strip()

    st.success("Excel loaded!")

    # Choose target length
    target_length = st.selectbox(
        "Select Target Length",
        [400, 200, 150, 100],
        index=1
    )

    # Identify numeric columns
    value_cols = df.columns[2:]

    # Weighting controls
    st.subheader("Optional Weighting")
    weight_cols = st.multiselect("Select columns to weight", list(value_cols))
    weight_factor = st.number_input(
        "Weight factor",
        min_value=1.0,
        max_value=10.0,
        value=1.0,
        step=0.1
    )

    # Clean numeric columns
    for col in value_cols:
        df[col] = pd.to_numeric(
            df[col].astype(str).str.replace(",", ".").str.strip(),
            errors="coerce"
        ).fillna(0)

    # RUN button
    if st.button("Run"):
        st.info("Computing combinations... please wait.")
        with st.spinner("Working..."):

            results = find_combinations(df, value_cols, weight_cols, weight_factor, target_length)

        if not results:
            st.error("No valid combinations found.")
        else:
            st.success(f"Found {len(results)} valid combinations!")

            # -------------------------
            # Build Excel WITHOUT sorting
            # -------------------------
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                unweighted_rows = []
                weighted_rows = []

                for combo_id, r in enumerate(results, start=1):
                    subset = r["Subset"].copy()
                    subset["Combination"] = combo_id
                    subset["Unweighted_Sum"] = r["Unweighted"]
                    unweighted_rows.append(subset)

                    subset2 = r["Subset"].copy()
                    subset2["Combination"] = combo_id
                    subset2["Weighted_Sum"] = r["Weighted"]
                    weighted_rows.append(subset2)

                pd.concat(unweighted_rows).to_excel(writer, sheet_name="Unweighted", index=False)
                pd.concat(weighted_rows).to_excel(writer, sheet_name="Weighted", index=False)

            # Download button
            st.download_button(
                "Download Excel Results",
                output.getvalue(),
                file_name="results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("Please upload an Excel file.")
