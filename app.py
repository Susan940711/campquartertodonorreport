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

AGE_SEMESTER_OUTPUT_COLUMNS = [
    "Period",
    "Organization",
    "Project Name",
    "indicator",
    "S1 U1 Male",
    "S1 U1 Female",
    "S1 1-5 Male",
    "S1 1-5 Female",
    "S1 Total",
    "S2 U1 Male",
    "S2 U1 Female",
    "S2 1-5 Male",
    "S2 1-5 Female",
    "S2 Total",
    "Annual U1 Male",
    "Annual U1 Female",
    "Annual 1-5 Male",
    "Annual 1-5 Female",
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


def strip_duplicate_suffix(column_name: str) -> str:
    return re.sub(r"\.\d+$", "", str(column_name))


def collapse_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    grouped_columns: Dict[str, List[str]] = {}
    ordered_bases: List[str] = []

    for column in df.columns:
        base_name = strip_duplicate_suffix(column)
        if base_name not in grouped_columns:
            grouped_columns[base_name] = []
            ordered_bases.append(base_name)
        grouped_columns[base_name].append(column)

    collapsed = pd.DataFrame(index=df.index)

    for base_name in ordered_bases:
        columns = grouped_columns[base_name]
        if len(columns) == 1:
            collapsed[base_name] = df[columns[0]]
            continue

        numeric_parts = [pd.to_numeric(df[col], errors="coerce") for col in columns]
        numeric_coverage = [part.notna().sum() for part in numeric_parts]

        if sum(numeric_coverage) > 0:
            collapsed[base_name] = sum(
                (part.fillna(0) for part in numeric_parts),
                pd.Series([0] * len(df), index=df.index),
            )
        else:
            collapsed[base_name] = df[columns].bfill(axis=1).iloc[:, 0]

    return collapsed


def detect_indicator_sheet(file_obj: io.BytesIO) -> str:
    xls = pd.ExcelFile(file_obj)
    candidates: List[Tuple[str, int]] = []

    for sheet in xls.sheet_names:
        try:
            sample = pd.read_excel(xls, sheet_name=sheet, nrows=3)
        except Exception:
            continue

        normalized = [normalize_text(c) for c in sample.columns]
        compact_cols = [compact_text(c) for c in sample.columns]
        sheet_name = normalize_text(sheet)
        score = 0

        if any("indicator" in c for c in normalized):
            score += 2
        if sheet_name == "indicator":
            score += 6
        elif "indicator" in sheet_name:
            score += 3
        if any("period" in c for c in normalized):
            score += 1
        if any("project" in c for c in normalized):
            score += 1
        quarter_hits = sum(
            1 for c in compact_cols if any(q in c for q in ["q1", "q2", "q3", "q4"])
        )
        if quarter_hits > 0:
            score += 8
        if any("uptoq3" in c for c in compact_cols):
            score -= 3
        if any(c.startswith("s1") or c.startswith("s2") for c in compact_cols):
            score -= 1
        if any(c.startswith("annual") for c in compact_cols):
            score -= 1

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


def prepare_indicator_raw_dataframe(uploaded_file) -> pd.DataFrame:
    data = uploaded_file.read()
    file_io = io.BytesIO(data)

    sheet_name = detect_indicator_sheet(file_io)
    file_io.seek(0)
    raw = pd.read_excel(file_io, sheet_name=sheet_name)
    raw = collapse_duplicate_columns(raw)

    standardized = standardize_base_columns(raw)

    output = standardized.copy()
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


def detect_age_band(normalized_col_name: str) -> Optional[str]:
    n = normalize_text(normalized_col_name)
    c = compact_text(normalized_col_name)

    has_u1 = bool(re.search(r"\bu\s*1\b|\bunder\s*1\b|\b<\s*1\b", n)) or (
        "u1" in c or "under1" in c
    )
    has_15 = bool(re.search(r"\b1\s*(?:to\s*)?5\b", n)) or ("15" in c)

    if has_u1:
        return "u1"
    if has_15:
        return "1-5"
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


def find_quarter_age_columns(df: pd.DataFrame) -> Dict[str, Dict[str, List[str]]]:
    quarter_map: Dict[str, Dict[str, List[str]]] = {
        "q1": {"u1_male": [], "u1_female": [], "15_male": [], "15_female": []},
        "q2": {"u1_male": [], "u1_female": [], "15_male": [], "15_female": []},
        "q3": {"u1_male": [], "u1_female": [], "15_male": [], "15_female": []},
        "q4": {"u1_male": [], "u1_female": [], "15_male": [], "15_female": []},
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

        gender = detect_gender_tokenized(n)
        age_band = detect_age_band(n)
        if not gender or not age_band:
            continue

        if age_band == "u1" and gender == "male":
            quarter_map[q]["u1_male"].append(col)
        elif age_band == "u1" and gender == "female":
            quarter_map[q]["u1_female"].append(col)
        elif age_band == "1-5" and gender == "male":
            quarter_map[q]["15_male"].append(col)
        elif age_band == "1-5" and gender == "female":
            quarter_map[q]["15_female"].append(col)

    return quarter_map


def sum_columns(df: pd.DataFrame, columns: List[str]) -> pd.Series:
    if not columns:
        return pd.Series([0] * len(df), index=df.index)

    total = pd.Series([0] * len(df), index=df.index)
    for col in columns:
        total = total + to_numeric_series(df[col])
    return total


def find_semester_columns(df: pd.DataFrame) -> Dict[str, Dict[str, List[str]]]:
    semester_map: Dict[str, Dict[str, List[str]]] = {
        "s1": {"male": [], "female": [], "target": []},
        "s2": {"male": [], "female": [], "target": []},
        "annual": {"male": [], "female": [], "target": []},
    }

    for col in df.columns:
        n = normalize_text(col)
        c = compact_text(col)

        sem = None
        if c.startswith("s1"):
            sem = "s1"
        elif c.startswith("s2"):
            sem = "s2"
        elif c.startswith("annual"):
            sem = "annual"

        if not sem:
            continue

        has_u1 = bool(re.search(r"\bu\s*1\b|\bunder\s*1\b|\b<\s*1\b", n)) or (
            "u1" in c or "under1" in c
        )
        has_15 = bool(re.search(r"\b1\s*(?:to\s*)?5\b", n)) or ("15" in c)
        has_age_band = has_u1 or has_15
        gender = detect_gender_tokenized(n)

        if gender == "male" and has_age_band:
            semester_map[sem]["male"].append(col)
        elif gender == "female" and has_age_band:
            semester_map[sem]["female"].append(col)
        elif "target" in n or "targ" in c:
            semester_map[sem]["target"].append(col)

    return semester_map


def find_semester_age_columns(df: pd.DataFrame) -> Dict[str, Dict[str, List[str]]]:
    semester_map: Dict[str, Dict[str, List[str]]] = {
        "s1": {"u1_male": [], "u1_female": [], "15_male": [], "15_female": []},
        "s2": {"u1_male": [], "u1_female": [], "15_male": [], "15_female": []},
        "annual": {"u1_male": [], "u1_female": [], "15_male": [], "15_female": []},
    }

    for col in df.columns:
        n = normalize_text(col)
        c = compact_text(col)

        sem = None
        if c.startswith("s1"):
            sem = "s1"
        elif c.startswith("s2"):
            sem = "s2"
        elif c.startswith("annual"):
            sem = "annual"

        if not sem:
            continue

        gender = detect_gender_tokenized(n)
        age_band = detect_age_band(n)
        if not gender or not age_band:
            continue

        if age_band == "u1" and gender == "male":
            semester_map[sem]["u1_male"].append(col)
        elif age_band == "u1" and gender == "female":
            semester_map[sem]["u1_female"].append(col)
        elif age_band == "1-5" and gender == "male":
            semester_map[sem]["15_male"].append(col)
        elif age_band == "1-5" and gender == "female":
            semester_map[sem]["15_female"].append(col)

    return semester_map


def compute_semester_metrics(df: pd.DataFrame) -> pd.DataFrame:
    quarter_cols = find_quarter_columns(df)
    semester_cols = find_semester_columns(df)

    result = pd.DataFrame(index=df.index)

    s1_male_quarter = sum_columns(df, quarter_cols["q1"]["male"]) + sum_columns(
        df, quarter_cols["q2"]["male"]
    )
    s1_female_quarter = sum_columns(df, quarter_cols["q1"]["female"]) + sum_columns(
        df, quarter_cols["q2"]["female"]
    )
    s1_target_quarter = sum_columns(df, quarter_cols["q1"]["target"]) + sum_columns(
        df, quarter_cols["q2"]["target"]
    )

    if (
        len(quarter_cols["q1"]["male"]) + len(quarter_cols["q2"]["male"]) == 0
        and len(quarter_cols["q1"]["female"]) + len(quarter_cols["q2"]["female"]) == 0
    ):
        result["S1 Male"] = sum_columns(df, semester_cols["s1"]["male"])
        result["S1 Female"] = sum_columns(df, semester_cols["s1"]["female"])
        result["S1 Target"] = sum_columns(df, semester_cols["s1"]["target"])
    else:
        result["S1 Male"] = s1_male_quarter
        result["S1 Female"] = s1_female_quarter
        result["S1 Target"] = s1_target_quarter
    result["S1 Total"] = result["S1 Male"] + result["S1 Female"]

    s2_male_quarter = sum_columns(df, quarter_cols["q3"]["male"]) + sum_columns(
        df, quarter_cols["q4"]["male"]
    )
    s2_female_quarter = sum_columns(df, quarter_cols["q3"]["female"]) + sum_columns(
        df, quarter_cols["q4"]["female"]
    )
    s2_target_quarter = sum_columns(df, quarter_cols["q3"]["target"]) + sum_columns(
        df, quarter_cols["q4"]["target"]
    )

    if (
        len(quarter_cols["q3"]["male"]) + len(quarter_cols["q4"]["male"]) == 0
        and len(quarter_cols["q3"]["female"]) + len(quarter_cols["q4"]["female"]) == 0
    ):
        result["S2 Male"] = sum_columns(df, semester_cols["s2"]["male"])
        result["S2 Female"] = sum_columns(df, semester_cols["s2"]["female"])
        result["S2 Target"] = sum_columns(df, semester_cols["s2"]["target"])
    else:
        result["S2 Male"] = s2_male_quarter
        result["S2 Female"] = s2_female_quarter
        result["S2 Target"] = s2_target_quarter
    result["S2 Total"] = result["S2 Male"] + result["S2 Female"]

    annual_male_semester = sum_columns(df, semester_cols["annual"]["male"])
    annual_female_semester = sum_columns(df, semester_cols["annual"]["female"])
    annual_target_semester = sum_columns(df, semester_cols["annual"]["target"])

    result["Annual Male"] = result["S1 Male"] + result["S2 Male"]
    result["Annual Female"] = result["S1 Female"] + result["S2 Female"]
    result["Annual Target"] = result["S1 Target"] + result["S2 Target"]

    if len(semester_cols["annual"]["male"]) > 0:
        result["Annual Male"] = annual_male_semester
    if len(semester_cols["annual"]["female"]) > 0:
        result["Annual Female"] = annual_female_semester
    if len(semester_cols["annual"]["target"]) > 0:
        result["Annual Target"] = annual_target_semester

    result["Annual Total"] = result["Annual Male"] + result["Annual Female"]

    return result


def compute_age_semester_metrics(df: pd.DataFrame) -> pd.DataFrame:
    quarter_age_cols = find_quarter_age_columns(df)
    semester_age_cols = find_semester_age_columns(df)

    result = pd.DataFrame(index=df.index)

    q_s1_detected = any(
        len(quarter_age_cols[q][bucket]) > 0
        for q in ["q1", "q2"]
        for bucket in ["u1_male", "u1_female", "15_male", "15_female"]
    )
    q_s2_detected = any(
        len(quarter_age_cols[q][bucket]) > 0
        for q in ["q3", "q4"]
        for bucket in ["u1_male", "u1_female", "15_male", "15_female"]
    )

    if q_s1_detected:
        result["S1 U1 Male"] = sum_columns(df, quarter_age_cols["q1"]["u1_male"]) + sum_columns(
            df, quarter_age_cols["q2"]["u1_male"]
        )
        result["S1 U1 Female"] = sum_columns(df, quarter_age_cols["q1"]["u1_female"]) + sum_columns(
            df, quarter_age_cols["q2"]["u1_female"]
        )
        result["S1 1-5 Male"] = sum_columns(df, quarter_age_cols["q1"]["15_male"]) + sum_columns(
            df, quarter_age_cols["q2"]["15_male"]
        )
        result["S1 1-5 Female"] = sum_columns(
            df, quarter_age_cols["q1"]["15_female"]
        ) + sum_columns(df, quarter_age_cols["q2"]["15_female"])
    else:
        result["S1 U1 Male"] = sum_columns(df, semester_age_cols["s1"]["u1_male"])
        result["S1 U1 Female"] = sum_columns(df, semester_age_cols["s1"]["u1_female"])
        result["S1 1-5 Male"] = sum_columns(df, semester_age_cols["s1"]["15_male"])
        result["S1 1-5 Female"] = sum_columns(df, semester_age_cols["s1"]["15_female"])

    result["S1 Total"] = (
        result["S1 U1 Male"]
        + result["S1 U1 Female"]
        + result["S1 1-5 Male"]
        + result["S1 1-5 Female"]
    )

    if q_s2_detected:
        result["S2 U1 Male"] = sum_columns(df, quarter_age_cols["q3"]["u1_male"]) + sum_columns(
            df, quarter_age_cols["q4"]["u1_male"]
        )
        result["S2 U1 Female"] = sum_columns(df, quarter_age_cols["q3"]["u1_female"]) + sum_columns(
            df, quarter_age_cols["q4"]["u1_female"]
        )
        result["S2 1-5 Male"] = sum_columns(df, quarter_age_cols["q3"]["15_male"]) + sum_columns(
            df, quarter_age_cols["q4"]["15_male"]
        )
        result["S2 1-5 Female"] = sum_columns(
            df, quarter_age_cols["q3"]["15_female"]
        ) + sum_columns(df, quarter_age_cols["q4"]["15_female"])
    else:
        result["S2 U1 Male"] = sum_columns(df, semester_age_cols["s2"]["u1_male"])
        result["S2 U1 Female"] = sum_columns(df, semester_age_cols["s2"]["u1_female"])
        result["S2 1-5 Male"] = sum_columns(df, semester_age_cols["s2"]["15_male"])
        result["S2 1-5 Female"] = sum_columns(df, semester_age_cols["s2"]["15_female"])

    result["S2 Total"] = (
        result["S2 U1 Male"]
        + result["S2 U1 Female"]
        + result["S2 1-5 Male"]
        + result["S2 1-5 Female"]
    )

    result["Annual U1 Male"] = result["S1 U1 Male"] + result["S2 U1 Male"]
    result["Annual U1 Female"] = result["S1 U1 Female"] + result["S2 U1 Female"]
    result["Annual 1-5 Male"] = result["S1 1-5 Male"] + result["S2 1-5 Male"]
    result["Annual 1-5 Female"] = result["S1 1-5 Female"] + result["S2 1-5 Female"]

    if any(len(semester_age_cols["annual"][k]) > 0 for k in ["u1_male", "u1_female", "15_male", "15_female"]):
        if len(semester_age_cols["annual"]["u1_male"]) > 0:
            result["Annual U1 Male"] = sum_columns(df, semester_age_cols["annual"]["u1_male"])
        if len(semester_age_cols["annual"]["u1_female"]) > 0:
            result["Annual U1 Female"] = sum_columns(df, semester_age_cols["annual"]["u1_female"])
        if len(semester_age_cols["annual"]["15_male"]) > 0:
            result["Annual 1-5 Male"] = sum_columns(df, semester_age_cols["annual"]["15_male"])
        if len(semester_age_cols["annual"]["15_female"]) > 0:
            result["Annual 1-5 Female"] = sum_columns(df, semester_age_cols["annual"]["15_female"])

    result["Annual Total"] = (
        result["Annual U1 Male"]
        + result["Annual U1 Female"]
        + result["Annual 1-5 Male"]
        + result["Annual 1-5 Female"]
    )

    return result


def prepare_indicator_dataframe(uploaded_file) -> pd.DataFrame:
    data = uploaded_file.read()
    file_io = io.BytesIO(data)

    sheet_name = detect_indicator_sheet(file_io)
    file_io.seek(0)
    raw = pd.read_excel(file_io, sheet_name=sheet_name)
    raw = collapse_duplicate_columns(raw)

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


def combine_reports_in_order(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    combined = pd.concat([df1, df2], ignore_index=True)
    combined["Period"] = combined["Period"].map(normalize_group_key)
    combined["indicator"] = combined["indicator"].map(normalize_group_key)
    combined = combined[(combined["Period"] != "") & (combined["indicator"] != "")]

    numeric_cols = [c for c in combined.columns if c not in GROUPING_KEYS]
    for col in numeric_cols:
        combined[col] = to_numeric_series(combined[col])

    return combined.reset_index(drop=True)


def aggregate_report_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.reset_index(drop=True)

    group_keys = [col for col in GROUPING_KEYS if col in df.columns]
    if not group_keys:
        return df.reset_index(drop=True)

    numeric_cols = [col for col in df.columns if col not in group_keys]
    for col in numeric_cols:
        df[col] = to_numeric_series(df[col])

    grouped = (
        df.groupby(group_keys, dropna=False, as_index=False)[numeric_cols]
        .sum(min_count=1)
        .fillna(0)
    )

    ordered_columns = group_keys + [col for col in df.columns if col in numeric_cols]
    return grouped[ordered_columns]


def build_combined_indicator_raw_report(file1, file2) -> pd.DataFrame:
    df1 = prepare_indicator_raw_dataframe(file1)
    df2 = prepare_indicator_raw_dataframe(file2)

    combined = combine_reports_in_order(df1, df2)
    return aggregate_report_rows(combined)


def build_combined_semester_report(file1, file2) -> pd.DataFrame:
    df1 = prepare_indicator_dataframe(file1)
    df2 = prepare_indicator_dataframe(file2)

    combined = combine_reports_in_order(df1, df2)
    final_df = aggregate_report_rows(combined)

    for col in ESSENTIAL_OUTPUT_COLUMNS:
        if col not in final_df.columns:
            final_df[col] = 0

    final_df = final_df[ESSENTIAL_OUTPUT_COLUMNS]

    for col in ESSENTIAL_OUTPUT_COLUMNS:
        if col not in ["Period", "Organization", "Project Name", "indicator"]:
            final_df[col] = to_numeric_series(final_df[col]).round(0).astype(int)

    return final_df


def prepare_age_semester_dataframe(uploaded_file) -> pd.DataFrame:
    data = uploaded_file.read()
    file_io = io.BytesIO(data)

    sheet_name = detect_indicator_sheet(file_io)
    file_io.seek(0)
    raw = pd.read_excel(file_io, sheet_name=sheet_name)
    raw = collapse_duplicate_columns(raw)

    standardized = standardize_base_columns(raw)

    output = standardized[["Period", "indicator"]].copy()
    metrics = compute_age_semester_metrics(raw)
    for col in metrics.columns:
        output[col] = metrics[col]

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


def build_combined_age_semester_report(file1, file2) -> pd.DataFrame:
    df1 = prepare_age_semester_dataframe(file1)
    df2 = prepare_age_semester_dataframe(file2)

    combined = combine_reports_in_order(df1, df2)
    final_df = aggregate_report_rows(combined)

    for col in AGE_SEMESTER_OUTPUT_COLUMNS:
        if col not in final_df.columns:
            final_df[col] = 0

    final_df = final_df[AGE_SEMESTER_OUTPUT_COLUMNS]

    for col in AGE_SEMESTER_OUTPUT_COLUMNS:
        if col not in ["Period", "Organization", "Project Name", "indicator"]:
            final_df[col] = to_numeric_series(final_df[col]).round(0).astype(int)

    return final_df


def build_summary_combine_sheet(summary_df: pd.DataFrame) -> pd.DataFrame:
    if summary_df.empty:
        return pd.DataFrame(columns=ESSENTIAL_OUTPUT_COLUMNS)

    summary_copy = summary_df.copy()
    for col in ESSENTIAL_OUTPUT_COLUMNS:
        if col not in summary_copy.columns:
            summary_copy[col] = 0

    numeric_cols = [
        col for col in ESSENTIAL_OUTPUT_COLUMNS if col not in ["Period", "Organization", "Project Name", "indicator"]
    ]
    for col in numeric_cols:
        summary_copy[col] = to_numeric_series(summary_copy[col])

    grouped = (
        summary_copy.groupby(
            ["Period", "Organization", "Project Name", "indicator"],
            dropna=False,
            as_index=False,
        )[numeric_cols]
        .sum(min_count=1)
        .fillna(0)
    )

    for col in ESSENTIAL_OUTPUT_COLUMNS:
        if col not in grouped.columns:
            grouped[col] = 0

    return grouped[ESSENTIAL_OUTPUT_COLUMNS]


def dataframe_to_excel_bytes(
    indicator_df: pd.DataFrame,
    age_df: pd.DataFrame,
    indicator_raw_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        indicator_df.to_excel(writer, sheet_name="Indicator Semester Achievement", index=False)
        age_df.to_excel(writer, sheet_name="Age_semester", index=False)
        indicator_raw_df.to_excel(writer, sheet_name="Indicator Sheet Combined", index=False)
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
                file1.seek(0)
                file2.seek(0)
                final_report = build_combined_semester_report(file1, file2)
                file1.seek(0)
                file2.seek(0)
                age_semester_report = build_combined_age_semester_report(file1, file2)
                file1.seek(0)
                file2.seek(0)
                indicator_raw_report = build_combined_indicator_raw_report(file1, file2)
                st.success("Semester report generated successfully.")
                st.subheader("Indicator Semester Achievement")
                st.dataframe(final_report, use_container_width=True)
                st.subheader("Age_semester")
                st.dataframe(age_semester_report, use_container_width=True)
                st.subheader("Indicator Sheet Combined")
                st.dataframe(indicator_raw_report, use_container_width=True)

                summary_combine_df = build_summary_combine_sheet(final_report)
                excel_data = dataframe_to_excel_bytes(
                    final_report,
                    age_semester_report,
                    indicator_raw_report,
                )
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
