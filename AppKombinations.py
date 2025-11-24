import streamlit as st
import pandas as pd
from io import BytesIO

# ----------------------------------------
# Hilfsfunktion: Kombos über Backtracking
# ----------------------------------------
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


# ----------------------------------------
# Streamlit UI
# ----------------------------------------
st.title("Kombinations-Finder (Längen + Gewichtung)")

uploaded = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])

if uploaded:
    df = pd.read_excel(uploaded)
    df.columns = df.columns.str.replace("\r", "").str.replace("\n", "").str.strip()

    st.success("Datei geladen!")

    # -------------------------
    # Länge auswählen
    # -------------------------
    target_length = st.selectbox(
        "Ziel-Länge auswählen",
        [400, 200, 150, 100],
        index=1
    )

    # Werte-Spalten erkennen
    value_cols = df.columns[2:]

    # -------------------------
    # Gewichtete Spalten wählen
    # -------------------------
    st.subheader("Spalten für Gewichtung auswählen")
    weight_cols = st.multiselect(
        "Spalten",
        list(value_cols)
    )

    weight_factor = st.number_input(
        "Gewichtungsfaktor (z. B. 2 = doppelte Gewichtung)",
        min_value=1.0,
        max_value=10.0,
        value=2.0,
        step=0.1
    )

    # Daten in numerisch umwandeln
    for col in value_cols:
        df[col] = pd.to_numeric(
            df[col].astype(str).str.replace(",", ".").str.strip(),
            errors="coerce"
        ).fillna(0)

    # -------------------------
    # RUN BUTTON
    # -------------------------
    if st.button("Berechnung starten"):
        st.info("Berechnung läuft... kann bei vielen Rows dauern!")

        results = find_combinations(df, value_cols, weight_cols, weight_factor, target_length)

        if not results:
            st.error("Keine gültigen Kombinationen gefunden.")
        else:
            st.success(f"{len(results)} Kombinationen gefunden!")

            # -------------------------
            # Sortieren
            # -------------------------
            results.sort(key=lambda x: x["Weighted"], reverse=True)

            # -------------------------
            # Excel erzeugen
            # -------------------------
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                unweighted_rows = []
                weighted_rows = []

                for combo_id, r in enumerate(results, start=1):
                    subset = r["Subset"].copy()
                    subset["Kombination"] = combo_id
                    subset["Unweighted_Sum"] = r["Unweighted"]
                    unweighted_rows.append(subset)

                    subset2 = r["Subset"].copy()
                    subset2["Kombination"] = combo_id
                    subset2["Weighted_Sum"] = r["Weighted"]
                    weighted_rows.append(subset2)

                unweighted_df = pd.concat(unweighted_rows)
                unweighted_df = unweighted_df.sort_values("Unweighted_Sum", ascending=False)
                unweighted_df.to_excel(writer, sheet_name="Unweighted", index=False)

                weighted_df = pd.concat(weighted_rows)
                weighted_df = weighted_df.sort_values("Weighted_Sum", ascending=False)
                weighted_df.to_excel(writer, sheet_name="Weighted", index=False)
            # -------------------------
            # Download
            # -------------------------
            st.download_button(
                label="Ergebnis als Excel herunterladen",
                data=output.getvalue(),
                file_name="kombinationen.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.success("Excel bereit!")


else:
    st.info("Bitte eine Excel-Datei hochladen.")
