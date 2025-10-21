import pandas as pd
import argparse
import os
from typing import List

def find_first_present(df: pd.DataFrame, candidates: List[str]):
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    raise KeyError(f"Could not find any of the expected columns: {candidates}. Found columns: {list(df.columns)}")

def coerce_percentage(s: pd.Series) -> pd.Series:
    # Accept numbers like 72, 72.0, "72", "72%", " 72 % "
    def to_num(x):
        if pd.isna(x):
            return pd.NA
        if isinstance(x, (int, float)):
            return float(x)
        x_str = str(x).strip()
        if x_str.endswith("%"):
            x_str = x_str[:-1].strip()
        try:
            return float(x_str)
        except Exception:
            return pd.NA
    out = s.map(to_num)
    return pd.to_numeric(out, errors="coerce")

def summarize_failing(df: pd.DataFrame, threshold: float = 65.0,
                      student_col_cands=None, pct_col_cands=None, course_col_cands=None):
    if student_col_cands is None:
        student_col_cands = ["Student", "Name", "Student Name"]
    if pct_col_cands is None:
        pct_col_cands = ["Pct", "PCT", "Percentage", "Percent"]
    if course_col_cands is None:
        course_col_cands = ["Course", "Course Name", "Class"]

    student_col = find_first_present(df, student_col_cands)
    pct_col = find_first_present(df, pct_col_cands)
    course_col = find_first_present(df, course_col_cands)

    # Clean
    df = df.copy()
    df[pct_col] = coerce_percentage(df[pct_col])

    # Filter failing
    failing = df[df[pct_col] < threshold].copy()
    # Keep only needed columns in consistent names
    failing = failing.rename(columns={student_col: "Student", course_col: "Course", pct_col: "Pct"})
    failing = failing[["Student", "Course", "Pct"]].sort_values(["Student", "Course"]).reset_index(drop=True)

    # Build wide summary like the example
    # For each student, list failing course names in alphabetical order (or by ascending percentage)
    failing["Course_SortKey"] = failing["Course"].astype(str)
    # You can swap the next line to `["Pct", "Course_SortKey"]` if you prefer lowest grades first
    failing = failing.sort_values(["Student", "Course_SortKey", "Pct"], ascending=[True, True, True])

    course_lists = failing.groupby("Student")["Course"].apply(list)
    totals = course_lists.map(len)

    # Determine max number of failing courses for column count
    max_courses = totals.max() if len(totals) else 0

    # Construct the wide DF
    data = []
    for student, courses in course_lists.items():
        row = {"Student": student, "Total": len(courses)}
        for i, c in enumerate(courses, start=1):
            row[f"Course {i}"] = c
        # fill the rest with blanks up to max_courses
        for i in range(len(courses)+1, max_courses+1):
            row[f"Course {i}"] = ""
        data.append(row)

    summary = pd.DataFrame(data)

    # If there are students with no failing classes, they won't appear – by design.
    # Sort summary by Total desc then Student asc
    if not summary.empty:
        summary = summary.sort_values(["Total", "Student"], ascending=[False, True]).reset_index(drop=True)

    return failing.drop(columns=["Course_SortKey"], errors="ignore"), summary

def read_any(path: str, sheet: str | None = None) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in [".csv", ".tsv", ".txt"]:
        # Try comma first; fallback to tab
        try:
            return pd.read_csv(path)
        except Exception:
            return pd.read_csv(path, sep="\t")
    elif ext in [".xlsx", ".xls"]:
        return pd.read_excel(path, sheet_name=sheet if sheet else 0)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Please provide CSV or Excel.")

def write_outputs(failing_rows: pd.DataFrame, summary: pd.DataFrame, out_path: str):
    # Ensure Excel extension
    base, ext = os.path.splitext(out_path)
    if ext.lower() not in [".xlsx", ".xls"]:
        out_path = base + ".xlsx"
    # Write Excel with two sheets
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        failing_rows.to_excel(writer, sheet_name="Failing_Rows", index=False)
        summary.to_excel(writer, sheet_name="Failing_Summary", index=False)
    # Also write a CSV of the summary for quick viewing
    summary_csv = base + "_readable.csv"
    summary.to_csv(summary_csv, index=False)
    return out_path, summary_csv

def main():
    parser = argparse.ArgumentParser(description="Build a failing-students summary from a grade export.")
    parser.add_argument("--in", dest="in_path", required=True, help="Input CSV/Excel grades file")
    parser.add_argument("--sheet", dest="sheet", default=None, help="Sheet name (for Excel). Defaults to first sheet.")
    parser.add_argument("--out", dest="out_path", default="failing_summary.xlsx", help="Output Excel path")
    parser.add_argument("--threshold", dest="threshold", type=float, default=65.0, help="Failing threshold (default=65)")
    args = parser.parse_args()

    df = read_any(args.in_path, args.sheet)
    failing_rows, summary = summarize_failing(df, threshold=args.threshold)
    out_xlsx, out_csv = write_outputs(failing_rows, summary, args.out_path)
    print(f"Wrote Excel to: {out_xlsx}")
    print(f"Wrote CSV to:   {out_csv}")

if __name__ == "__main__":
    # If the user runs this file directly, parse CLI args
    main()