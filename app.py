import io
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Camp Immunization Semester Report", layout="wide")


ESSENTIAL_OUTPUT_COLUMNS = [
    "Period",
    "Organization",
    "Project Name",
    "indicator",
    "S1 Target",
    "S1 Male",
    "S1 Female",
    "S1 Total",
    "S2 Target",
    "S2 Male",
    "S2 Female",
    "S2 Total",
    "Annual Target",
    "Annual Male",
    "Annual Female",
    "Annual Total",
]


GROUPING_KEYS = ["Period", "Organization", "Project Name", "indicator"]


def normalize_text(value: str) -> str:
    text = str(value).strip().lower()
    text = text.replace("_", " ")
    text = text.replace("-", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def compact_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", normalize_text(value))


def to_numeric_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def normalize_group_key(value) -> str:
    text = str(value if pd.notna(value) else "").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def detect_indicator_sheet(file_obj: io.BytesIO) -> str:
    xls = pd.ExcelFile(file_obj)
    candidates: List[Tuple[str, int]] = []

    for sheet in xls.sheet_names:
        try:
            sample = pd.read_excel(xls, sheet_name=sheet, nrows=3)
        except Exception:
            continue

        normalized = [normalize_text(c) for c in sample.columns]
        score = 0

        if any("indicator" in c for c in normalized):
            score += 2
        if any("period" in c for c in normalized):
            score += 1
        if any("project" in c for c in normalized):
            score += 1
        if any("q1" in c or "q2" in c or "q3" in c or "q4" in c for c in normalized):
            score += 1

        if score > 0:
            candidates.append((sheet, score))

    if not candidates:
        raise ValueError("Could not detect an indicator sheet. Please verify the source files.")

    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0]


def get_column_map(columns: List[str]) -> Dict[str, str]:
    mapping = {}
    normalized_map = {normalize_text(c): c for c in columns}

    for normalized, original in normalized_map.items():
        if normalized == "period" and "Period" not in mapping:
            mapping["Period"] = original
        if (
            "organization" in normalized or normalized.startswith("organiz")
        ) and "Organization" not in mapping:
            mapping["Organization"] = original
        if (
            "project name" in normalized
            or normalized.startswith("project na")
            or normalized == "project"
        ) and "Project Name" not in mapping:
            mapping["Project Name"] = original
        if "indicator" in normalized and "indicator" not in mapping:
            mapping["indicator"] = original

    return mapping


def standardize_base_columns(df: pd.DataFrame) -> pd.DataFrame:
    col_map = get_column_map(df.columns.tolist())

    if "Period" not in col_map or "indicator" not in col_map:
        raise ValueError(
            "Missing required columns. Indicator sheets must include Period and indicator columns."
        )

    standardized = df.copy()

    standardized["Period"] = standardized[col_map["Period"]]
    standardized["indicator"] = standardized[col_map["indicator"]]

    if "Project Name" in col_map:
        standardized["Project Name"] = standardized[col_map["Project Name"]]
    else:
        standardized["Project Name"] = "Camp Immunization"

    if "Organization" in col_map:
        standardized["Organization"] = standardized[col_map["Organization"]]
    else:
        standardized["Organization"] = "PRF"

    return standardized


def detect_gender_tokenized(normalized_col_name: str) -> Optional[str]:
    tokens = normalized_col_name.split(" ")
    compact = re.sub(r"[^a-z0-9]", "", normalized_col_name)

    female_tokens = {"female", "fem", "fema", "fen", "fer", "fe"}
    male_tokens = {"male", "mal", "ma"}

    if any(token.startswith("fem") or token in female_tokens for token in tokens):
        return "female"
    if any(tag in compact for tag in ["female", "fema", "fem", "fen", "fer"]):
        return "female"

    if any(token in male_tokens for token in tokens):
        return "male"
    if any(tag in compact for tag in ["male", "mal"]) or compact.endswith("ma"):
        return "male"

    return None


def find_quarter_columns(df: pd.DataFrame) -> Dict[str, Dict[str, List[str]]]:
    quarter_map: Dict[str, Dict[str, List[str]]] = {
        "q1": {"male": [], "female": [], "target": []},
        "q2": {"male": [], "female": [], "target": []},
        "q3": {"male": [], "female": [], "target": []},
        "q4": {"male": [], "female": [], "target": []},
    }

    for col in df.columns:
        n = normalize_text(col)
        c = compact_text(col)
        q = None
        for quarter in ["q1", "q2", "q3", "q4"]:
            if quarter in c:
                q = quarter
                break

        if not q:
            continue

        has_u1 = bool(re.search(r"\bu\s*1\b|\bunder\s*1\b|\b<\s*1\b", n)) or (
            "u1" in c or "under1" in c
        )
        has_15 = bool(re.search(r"\b1\s*(?:to\s*)?5\b", n)) or ("15" in c)
        has_age_band = has_u1 or has_15
        gender = detect_gender_tokenized(n)

        if gender == "male" and has_age_band:
            quarter_map[q]["male"].append(col)
        elif gender == "female" and has_age_band:
            quarter_map[q]["female"].append(col)
        elif re.search(r"\btarg\w*\b", n) or ("targ" in c):
            quarter_map[q]["target"].append(col)

    return quarter_map


def sum_columns(df: pd.DataFrame, columns: List[str]) -> pd.Series:
    if not columns:
        return pd.Series([0] * len(df), index=df.index)

    total = pd.Series([0] * len(df), index=df.index)
    for col in columns:
        total = total + to_numeric_series(df[col])
    return total


def compute_semester_metrics(df: pd.DataFrame) -> pd.DataFrame:
    quarter_cols = find_quarter_columns(df)

    result = pd.DataFrame(index=df.index)

    result["S1 Male"] = sum_columns(df, quarter_cols["q1"]["male"]) + sum_columns(
        df, quarter_cols["q2"]["male"]
    )
    result["S1 Female"] = sum_columns(df, quarter_cols["q1"]["female"]) + sum_columns(
        df, quarter_cols["q2"]["female"]
    )
    result["S1 Target"] = sum_columns(df, quarter_cols["q1"]["target"]) + sum_columns(
        df, quarter_cols["q2"]["target"]
    )
    result["S1 Total"] = result["S1 Male"] + result["S1 Female"]

    result["S2 Male"] = sum_columns(df, quarter_cols["q3"]["male"]) + sum_columns(
        df, quarter_cols["q4"]["male"]
    )
    result["S2 Female"] = sum_columns(df, quarter_cols["q3"]["female"]) + sum_columns(
        df, quarter_cols["q4"]["female"]
    )
    result["S2 Target"] = sum_columns(df, quarter_cols["q3"]["target"]) + sum_columns(
        df, quarter_cols["q4"]["target"]
    )
    result["S2 Total"] = result["S2 Male"] + result["S2 Female"]

    result["Annual Male"] = result["S1 Male"] + result["S2 Male"]
    result["Annual Female"] = result["S1 Female"] + result["S2 Female"]
    result["Annual Target"] = result["S1 Target"] + result["S2 Target"]
    result["Annual Total"] = result["Annual Male"] + result["Annual Female"]

    return result


def prepare_indicator_dataframe(uploaded_file) -> pd.DataFrame:
    data = uploaded_file.read()
    file_io = io.BytesIO(data)

    sheet_name = detect_indicator_sheet(file_io)
    file_io.seek(0)
    raw = pd.read_excel(file_io, sheet_name=sheet_name)

    standardized = standardize_base_columns(raw)

    output = standardized[["Period", "indicator"]].copy()

    metrics = compute_semester_metrics(raw)
    for col in metrics.columns:
        output[col] = metrics[col]

    # Final combined report should represent PRF aggregated from both source files.
    output["Organization"] = "PRF"
    output["Project Name"] = "Camp Immunization"

    output["Period"] = output["Period"].map(normalize_group_key)
    output["indicator"] = output["indicator"].map(normalize_group_key)
    output = output[(output["Period"] != "") & (output["indicator"] != "")]

    numeric_cols = [c for c in output.columns if c not in GROUPING_KEYS]
    for c in numeric_cols:
        output[c] = to_numeric_series(output[c])

    grouped = (
        output.groupby(GROUPING_KEYS, dropna=False, as_index=False)[numeric_cols]
        .sum(min_count=1)
        .fillna(0)
    )

    return grouped


def build_combined_semester_report(file1, file2) -> pd.DataFrame:
    df1 = prepare_indicator_dataframe(file1)
    df2 = prepare_indicator_dataframe(file2)

    combined = pd.concat([df1, df2], ignore_index=True)
    combined["Period"] = combined["Period"].map(normalize_group_key)
    combined["indicator"] = combined["indicator"].map(normalize_group_key)
    combined = combined[(combined["Period"] != "") & (combined["indicator"] != "")]

    numeric_cols = [c for c in combined.columns if c not in GROUPING_KEYS]

    final_df = (
        combined.groupby(GROUPING_KEYS, dropna=False, as_index=False)[numeric_cols]
        .sum(min_count=1)
        .fillna(0)
    )

    for col in ESSENTIAL_OUTPUT_COLUMNS:
        if col not in final_df.columns:
            final_df[col] = 0

    final_df = final_df[ESSENTIAL_OUTPUT_COLUMNS]

    for col in ESSENTIAL_OUTPUT_COLUMNS:
        if col not in ["Period", "Organization", "Project Name", "indicator"]:
            final_df[col] = to_numeric_series(final_df[col]).round(0).astype(int)

    return final_df


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Indicator Semester Achievement", index=False)
    output.seek(0)
    return output.read()


def main() -> None:
    st.title("Camp Immunization Semester Report Generator")
    st.write(
        "Upload the two quarterly EPI report files. The app combines indicator sheets and "
        "creates one semester achievement report (S1, S2, Annual)."
    )

    file1 = st.file_uploader(
        "Upload source file 1 (MaeLa_Camp_EPI_Quarterly_Report...)",
        type=["xlsx"],
        key="source1",
    )
    file2 = st.file_uploader(
        "Upload source file 2 (PRF_Quarterly_report...)",
        type=["xlsx"],
        key="source2",
    )

    if file1 and file2:
        if st.button("Generate Semester Report", type="primary"):
            try:
                final_report = build_combined_semester_report(file1, file2)
                st.success("Semester report generated successfully.")
                st.dataframe(final_report, use_container_width=True)

                excel_data = dataframe_to_excel_bytes(final_report)
                st.download_button(
                    "Download Semester Report Excel",
                    data=excel_data,
                    file_name="Camp_Immunization_Semester_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as exc:
                st.error(f"Failed to generate report: {exc}")
    else:
        st.info("Please upload both source files to continue.")


if __name__ == "__main__":
    main()
