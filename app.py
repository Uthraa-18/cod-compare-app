import streamlit as st
import pandas as pd
import numpy as np
import os, io, re, unicodedata

# ==============================
# Style & helpers (UI)
# ==============================
st.set_page_config(page_title="COD Compare", layout="wide")

st.markdown("""
<style>
.section-title {
  font-size: 1.35rem;
  font-weight: 700;
  display: inline-flex;
  align-items: center;
  gap: .5rem;
  margin: .25rem 0 .5rem 0;
}
.info-dot {
  display:inline-block;
  font-size: 0.95rem;
  line-height: 1;
  padding: .1rem .35rem;
  border-radius: 999px;
  border: 1px solid #aaa;
  color: #333;
  cursor: help;
}
.subtle {
  font-size: 0.95rem;
  color: #555;
  margin-top: .25rem;
}
.small-input .stNumberInput > div > div > input {
  font-size: .9rem;
}
.block-container { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)

def header_with_tip(text: str, tip: str):
    st.markdown(
        f"<div class='section-title'>{text}"
        f"<span class='info-dot' title='{tip}'>ⓘ</span></div>",
        unsafe_allow_html=True
    )

# ==============================
# Core regex & utils
# ==============================
RE_PM    = re.compile(r'(?:±|\+/-)\s*(\d+(?:[.,]\d+)?)', re.I)
RE_SIGNED= re.compile(r'^[\+\-]?\s*\d+(?:[.,]\d+)?$')
RE_NUM   = re.compile(r'[-+]?\d+(?:[.,]\d+)?')

def to_float(x):
    try: return float(str(x).replace(",", "."))
    except: return None

def norm(s):
    if s is None or (isinstance(s, float) and pd.isna(s)): return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower().strip().replace("’","'")

def get_ext(name): return os.path.splitext(name)[-1].lower()

def read_all_sheets(name, file_bytes):
    engine = "xlrd" if get_ext(name)==".xls" else "openpyxl"
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
    return {s: pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, engine=engine, header=None) for s in xls.sheet_names}

# ==============================
# COD extraction helpers
# ==============================
def find_codification_value_below(cod_sheets, label="codification", scan_down=30):
    target = norm(label)
    for sname, df in cod_sheets.items():
        R,C = df.shape
        for r in range(R):
            for c in range(C):
                if norm(df.iat[r,c]) == target:
                    for rr in range(r+1, min(R, r+1+scan_down)):
                        if norm(df.iat[rr,c]) != "":
                            return sname, str(df.iat[rr,c]).strip(), r, c
    return None, None, None, None

def find_stacked_anchor_vertical(df, words, max_gap=10):
    R,C = df.shape
    W = [w.lower() for w in words]
    for c in range(C):
        starts = [r for r in range(R) if W[0] in norm(df.iat[r,c])]
        for r0 in starts:
            rcur = r0
            ok = True
            for w in W[1:]:
                found = False
                for rr in range(rcur+1, min(R, rcur+1+max_gap)):
                    if w in norm(df.iat[rr,c]):
                        rcur = rr; found = True; break
                if not found: ok=False; break
            if ok: return rcur, c
    return None, None

def first_number_below(df, start_row, col, right_span=12, down_rows=4):
    R,C = df.shape
    for rr in range(start_row+1, min(R, start_row+1+down_rows)):
        for cc in range(col, min(C, col+right_span+1)):
            s = "" if pd.isna(df.iat[rr,cc]) else str(df.iat[rr,cc])
            m = RE_PM.search(s)
            if m:
                v = to_float(m.group(1))
                if v is not None: return v, rr, cc
            if norm(s) in {"±","+/-"}:
                for cc2 in range(cc+1, min(C, cc+1+3)):
                    s2 = "" if pd.isna(df.iat[rr,cc2]) else str(df.iat[rr,cc2])
                    m2 = RE_NUM.search(s2)
                    if m2:
                        v = to_float(m2.group(0))
                        if v is not None: return v, rr, cc2
            m3 = RE_NUM.search(s)
            if m3:
                v = to_float(m3.group(0))
                if v is not None: return v, rr, cc
    return None, None, None

def two_signed_values_below_same_column(df, start_row, col, max_rows=8):
    R,_ = df.shape
    vals = []
    for rr in range(start_row+1, min(R, start_row+1+max_rows)):
        s = "" if pd.isna(df.iat[rr,col]) else str(df.iat[rr,col]).strip()
        if RE_SIGNED.match(s):
            x = to_float(s); 
            if x is not None: vals.append(x)
        elif norm(s) in {"±","+/-"} and col+1 < df.shape[1]:
            s2 = "" if pd.isna(df.iat[rr,col+1]) else str(df.iat[rr,col+1]).strip()
            if RE_NUM.match(s2):
                x = to_float(s2)
                if x is not None:
                    vals.append(+abs(x)); vals.append(-abs(x))
    for v in vals:
        if -v in vals:
            return +abs(v), -abs(v)
    return None, None

# ==============================
# Number extraction from PDJ/TCM rows
# ==============================
def row_numbers(df, r):
    nums=[]
    row=df.iloc[r,:].tolist()
    for i,v in enumerate(row):
        s = "" if pd.isna(v) else str(v)
        for m in RE_PM.findall(s):
            x=to_float(m)
            if x is not None: nums += [+abs(x), -abs(x)]
        if norm(s) in {"±","+/-"}:
            for j in range(i+1, min(len(row), i+4)):
                s2 = "" if pd.isna(row[j]) else str(row[j])
                for m2 in RE_NUM.findall(s2):
                    x=to_float(m2)
                    if x is not None: nums += [+abs(x), -abs(x)]
    for v in row:
        s="" if pd.isna(v) else str(v)
        for m in RE_NUM.findall(s):
            x=to_float(m); 
            if x is not None: nums.append(x)
    return nums

def sheet_numbers(df):
    nums=[]
    for rr in range(df.shape[0]):
        nums += row_numbers(df, rr)
    return nums

def find_key_positions(df, key):
    key=str(key).strip()
    pos=[]
    R,C=df.shape
    for r in range(R):
        for c in range(C):
            s="" if pd.isna(df.iat[r,c]) else str(df.iat[r,c]).strip()
            if s==key: pos.append((r,c))
    return pos

# ==============================
# Epsilon matching
# ==============================
def approx_equal(a,b,tol): 
    return a is not None and b is not None and abs(a-b) <= tol

def contains_value_eps(nums, val, tol):
    return any(approx_equal(x,val,tol) for x in nums)

def contains_pm_pair_eps(nums, mag, tol):
    return any(approx_equal(x,+abs(mag),tol) for x in nums) and \
           any(approx_equal(x,-abs(mag),tol) for x in nums)

def fmt_pm(m):
    s=f"{abs(m):.2f}".rstrip("0").rstrip(".")
    return f"+/- {s}"

# ==============================
# App UI
# ==============================
st.title("🔎 COD, PDJ, TCM Automatic Validation")

header_with_tip(
    "What this does",
    "Reads Nominal & Tolerance from COD (strict stacked headers) and checks PDJ/TCM rows for matches."
)
st.caption("Epsilon allows 1.41 ≈ 1.4")

with st.container():
    st.markdown("<div class='small-input'>", unsafe_allow_html=True)
    eps = st.number_input("Numeric tolerance (epsilon)",
                          0.0, 0.2, 0.02, 0.01,
                          help="Lets ±1.41 match ±1.4")
    st.markdown("</div>", unsafe_allow_html=True)

header_with_tip(
    "Upload COD workbook (.xls/.xlsx)",
    "Reads Codification + Nominal + Tolerance from COD file."
)
cod_file = st.file_uploader("", type=["xls","xlsx"], key="cod")

header_with_tip(
    "Upload PDJ/TCM/others",
    "PDJ & TCM check only row of the key. Others do row-first then sheet."
)
other_files = st.file_uploader("", type=["xls","xlsx"], accept_multiple_files=True, key="others")

# ==============================
# Main logic
# ==============================
if cod_file and other_files:
    cod_bytes = cod_file.read(); cod_file.seek(0)
    try:
        cod_sheets = read_all_sheets(cod_file.name, cod_bytes)
    except Exception as e:
        st.error(f"❌ Failed to read COD: {e}")
        st.stop()

    # --- 1) Extract Codification (value below label)
    s_cod, key_value, _, _ = find_codification_value_below(cod_sheets, "codification")
    if not key_value:
        st.error("Couldn't find Codification value below label.")
        st.stop()

    st.markdown(f"<div class='subtle'>🔑 Compared Key: <code>{key_value}</code></div>",
                unsafe_allow_html=True)

    df_cod = cod_sheets[s_cod]

    # --- 2) Nominal (Objectif → Nominal → Jeu)
    nr, nc = find_stacked_anchor_vertical(df_cod, ["objectif","nominal","jeu"], max_gap=10)
    if nr is None:
        st.error("Couldn't locate stacked header Objectif → Nominal → Jeu in COD.")
        st.stop()
    cod_nominal, _, _ = first_number_below(df_cod, nr, nc)
    if cod_nominal is None:
        st.error("Couldn't extract Nominal under that header.")
        st.stop()

    # --- 3) Tolerance (Calcul → Disp) STRICT: +x & -x in same column
    tr, tc = find_stacked_anchor_vertical(df_cod, ["calcul","disp"], max_gap=10)
    if tr is None:
        st.error("Couldn't locate stacked header Calcul → Disp.")
        st.stop()

    pos_val, neg_val = two_signed_values_below_same_column(df_cod, tr, tc)
    if pos_val is None or neg_val is None:
        pm_mag, _, _ = first_number_below(df_cod, tr, tc)
        if pm_mag is None:
            st.error("Couldn't extract tolerance from COD.")
            st.stop()
        tol_mag = abs(pm_mag)
    else:
        tol_mag = abs(pos_val)

    ref_nom_disp = float(f"{cod_nominal:.2f}")
    ref_tol_disp = float(f"{tol_mag:.2f}")

    st.write(f"**Reference Nominal (COD):** {ref_nom_disp}")
    st.write(f"**Reference Tolerance (COD):** {fmt_pm(ref_tol_disp)}")
    st.caption(f"Epsilon used: {eps}")

    # ==============================
    # 4) Compare with all other workbooks
    # ==============================
    results=[]
    for f in other_files:
        f_bytes = f.read(); f.seek(0)

        try:
            sheets = read_all_sheets(f.name, f_bytes)
        except Exception as e:
            st.warning(f"Skipping {f.name}: {e}")
            continue

        tag = f.name
        is_pdj = tag.upper().startswith("PDJ")
        is_tcm = tag.upper().startswith("TCM")

        for sname, df in sheets.items():

            pos = find_key_positions(df, key_value)
            if not pos: 
                continue

            # For each row where key appears
            for (r, _) in pos:

                # ==============================
                # IMPORTANT FIX:
                # TCM must behave like PDJ: only check the row where key appears
                # ==============================
                if is_pdj:
                    nums = row_numbers(df, r)

                elif is_tcm:
                    nums = row_numbers(df, r)   # <<< FIXED HERE (previously sheet_numbers)

                else:
                    row_nums = row_numbers(df, r)
                    if contains_value_eps(row_nums, cod_nominal, eps) or \
                       contains_pm_pair_eps(row_nums, tol_mag, eps):
                        nums = row_nums
                    else:
                        nums = sheet_numbers(df)

                nominal_ok = contains_value_eps(nums, cod_nominal, eps)
                tol_ok     = contains_pm_pair_eps(nums, tol_mag, eps)

                matched=[]
                if nominal_ok: matched.append(f"{ref_nom_disp}")
                if tol_ok:     matched.append(fmt_pm(ref_tol_disp))

                results.append({
                    "Compared Key": key_value,
                    "File": tag,
                    "Sheet": sname,
                    "Key Row": r+1,
                    "Reference Nominal": ref_nom_disp,
                    "Reference Tolerance": fmt_pm(ref_tol_disp),
                    "Nominal Found": "Yes" if nominal_ok else "No",
                    "Tolerance Found": "Yes" if tol_ok else "No",
                    "Matched Numbers": ", ".join(matched)
                })

    # ==============================
    # Final Output
    # ==============================
    if results:
        df_out = pd.DataFrame(results)
        st.write("### 📊 Results")
        st.dataframe(df_out, use_container_width=True)
        st.download_button("⬇️ Download results (CSV)",
                           df_out.to_csv(index=False),
                           "cod_comparison_results.csv",
                           "text/csv")
    else:
        st.warning("No matches found in uploaded files.")
