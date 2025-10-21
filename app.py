import io
import pandas as pd
import streamlit as st

# ---- your logic (paste from your script) ----
from typing import List
import os

def find_first_present(df: pd.DataFrame, candidates: List[str]):
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    raise KeyError(f"Could not find any of the expected columns: {candidates}. Found columns: {list(df.columns)}")

def coerce_percentage(s: pd.Series) -> pd.Series:
    def to_num(x):
        if pd.isna(x): return pd.NA
        if isinstance(x, (int, float)): return float(x)
        x_str = str(x).strip()
        if x_str.endswith("%"): x_str = x_str[:-1].strip()
        try: return float(x_str)
        except: return pd.NA
    return pd.to_numeric(s.map(to_num), errors="coerce")

def summarize_failing(df: pd.DataFrame, threshold: float = 65.0,
                      student_col_cands=None, pct_col_cands=None, course_col_cands=None):
    if student_col_cands is None: student_col_cands = ["Student", "Name", "Student Name"]
    if pct_col_cands is None: pct_col_cands = ["Pct", "PCT", "Percentage", "Percent"]
    if course_col_cands is None: course_col_cands = ["Course", "Course Name", "Class"]

    student_col = find_first_present(df, student_col_cands)
    pct_col = find_first_present(df, pct_col_cands)
    course_col = find_first_present(df, course_col_cands)

    df = df.copy()
    df[pct_col] = coerce_percentage(df[pct_col])
    failing = df[df[pct_col] < threshold].copy()
    failing = failing.rename(columns={student_col: "Student", course_col: "Course", pct_col: "Pct"})
    failing = failing[["Student", "Course", "Pct"]].sort_values(["Student", "Course"]).reset_index(drop=True)

    failing["Course_SortKey"] = failing["Course"].astype(str)
    failing = failing.sort_values(["Student", "Course_SortKey", "Pct"])

    course_lists = failing.groupby("Student")["Course"].apply(list)
    totals = course_lists.map(len)
    max_courses = totals.max() if len(totals) else 0

    data = []
    for student, courses in course_lists.items():
        row = {"Student": student, "Total": len(courses)}
        for i, c in enumerate(courses, start=1):
            row[f"Course {i}"] = c
        for i in range(len(courses)+1, max_courses+1):
            row[f"Course {i}"] = ""
        data.append(row)
    summary = pd.DataFrame(data)
    if not summary.empty:
        summary = summary.sort_values(["Total", "Student"], ascending=[False, True]).reset_index(drop=True)
    return failing.drop(columns=["Course_SortKey"], errors="ignore"), summary

# ---- UI ----
st.title("Failing Students Summary")
st.write("Upload a CSV or Excel (first sheet by default). Set the failing threshold and download the results.")

uploaded = st.file_uploader("Upload CSV/XLSX", type=["csv", "xlsx", "xls"])
threshold = st.number_input("Failing threshold", value=65.0, step=1.0)

if uploaded:
    # Read file automatically
    if uploaded.name.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded, sheet_name=0)  # Always reads the first sheet
    else:
        df = pd.read_csv(uploaded)

    failing_rows, summary = summarize_failing(df, threshold=threshold)
    st.subheader("Preview – Failing_Summary")
    st.dataframe(summary)

    # build downloadable files in-memory
    # 1) Excel with two sheets
    xls_buf = io.BytesIO()
    with pd.ExcelWriter(xls_buf, engine="openpyxl") as writer:
        failing_rows.to_excel(writer, sheet_name="Failing_Rows", index=False)
        summary.to_excel(writer, sheet_name="Failing_Summary", index=False)
    xls_buf.seek(0)
    st.download_button("⬇️ Download Summary Excel", data=xls_buf,
                       file_name="failing_summary_excel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # 2) CSV (summary only)
    csv_buf = summary.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download Summary CSV", data=csv_buf,
                       file_name="failing_summary_csv.csv", mime="text/csv")
