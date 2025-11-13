import streamlit as st
import pandas as pd
import numpy as np
import os, io, re, unicodedata

# ============================================================
# SAME CODE AS BEFORE… (all imports, css, helpers unchanged)
# ============================================================

# (Your existing long block remains unchanged until the results section)

# ==============================
# Final Output
# ==============================
    if results:
        df_out = pd.DataFrame(results)

        # ============================================================
        # NEW: Compute Unmatched Numbers column
        # ============================================================
        unmatched_list = []
        for row in results:
            # All numbers found during processing
            file = row["File"]
            sheet = row["Sheet"]
            key_row = row["Key Row"]

            # Retrieve the dataframe again
            df_target = read_all_sheets(file, open(file, "rb").read())[sheet]
            nums_row = row_numbers(df_target, key_row - 1)

            # Remove nominal and tolerance matched values
            ref_nom = row["Reference Nominal"]
            ref_tol = float(row["Reference Tolerance"].replace("+/- ", ""))

            remaining = []
            for val in nums_row:
                if approx_equal(val, ref_nom, eps):
                    continue
                if approx_equal(val, +ref_tol, eps) or approx_equal(val, -ref_tol, eps):
                    continue
                remaining.append(val)

            unmatched_list.append(", ".join([str(round(x, 3)) for x in remaining]) if remaining else "")

        df_out["Unmatched Numbers"] = unmatched_list

        # ============================================================
        # NEW: Color YES/NO columns (Nominal Found, Tolerance Found)
        # ============================================================
        def color_yes_no(val):
            if val == "Yes":
                return "background-color: #b6f3b6; color: black;"   # light green
            else:
                return "background-color: #ffb3b3; color: black;"   # light red

        styled = df_out.style.applymap(color_yes_no, subset=["Nominal Found", "Tolerance Found"])

        st.write("### 📊 Results")
        st.dataframe(styled, use_container_width=True)

        st.download_button(
            "⬇️ Download results (CSV)",
            df_out.to_csv(index=False),
            "cod_comparison_results.csv",
            "text/csv",
        )
    else:
        st.warning("No matches found in uploaded files.")
