import pandas as pd

from app import build_summary_combine_sheet, combine_reports_in_order


def test_combine_reports_in_order_keeps_rows_in_input_order() -> None:
    df1 = pd.DataFrame(
        [
            {
                "Period": "Q1",
                "Organization": "PRF",
                "Project Name": "Camp Immunization",
                "indicator": "Vaccination",
                "S1 Male": 1,
                "S1 Female": 2,
            }
        ]
    )
    df2 = pd.DataFrame(
        [
            {
                "Period": "Q2",
                "Organization": "PRF",
                "Project Name": "Camp Immunization",
                "indicator": "Vaccination",
                "S1 Male": 3,
                "S1 Female": 4,
            }
        ]
    )

    combined = combine_reports_in_order(df1, df2)

    assert len(combined) == 2
    assert combined["Period"].tolist() == ["Q1", "Q2"]
    assert combined["indicator"].tolist() == ["Vaccination", "Vaccination"]


def test_build_summary_combine_sheet_contains_combined_rows() -> None:
    summary_df = pd.DataFrame(
        [
            {
                "Period": "Q1",
                "Organization": "PRF",
                "Project Name": "Camp Immunization",
                "indicator": "Vaccination",
                "S1 Male": 1,
                "S1 Female": 2,
            },
            {
                "Period": "Q1",
                "Organization": "PRF",
                "Project Name": "Camp Immunization",
                "indicator": "Vaccination",
                "S1 Male": 3,
                "S1 Female": 4,
            },
        ]
    )

    combined_sheet = build_summary_combine_sheet(summary_df)

    assert len(combined_sheet) == 1
    assert combined_sheet.iloc[0]["S1 Male"] == 4
    assert combined_sheet.iloc[0]["S1 Female"] == 6
