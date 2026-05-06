from __future__ import annotations

from datetime import datetime

import pandas as pd
import streamlit as st

from scoring import build_ranked_accounts, dataframe_to_excel_bytes


def metric_card(label: str, value: int) -> None:
    st.metric(label=label, value=f"{value:,}")


def render_summary(summary: dict[str, int]) -> None:
    first_row = st.columns(4)
    with first_row[0]:
        metric_card("Total Accounts", summary["total_accounts"])
    with first_row[1]:
        metric_card("Epic Customers", summary["epic_customers"])
    with first_row[2]:
        metric_card("New Epic Systems", summary["new_epic_systems"])
    with first_row[3]:
        metric_card("Imprivata Customers", summary["imprivata_customers"])

    second_row = st.columns(5)
    with second_row[0]:
        metric_card("Tier 1", summary["tier_1"])
    with second_row[1]:
        metric_card("Tier 2", summary["tier_2"])
    with second_row[2]:
        metric_card("Tier 3", summary["tier_3"])
    with second_row[3]:
        metric_card("Hold", summary["hold"])
    with second_row[4]:
        metric_card("Exclude", summary["exclude"])


def render_filters(ranked: pd.DataFrame) -> pd.DataFrame:
    filter_columns = st.columns([1.2, 1.2, 1])

    tier_options = ["All"] + sorted(ranked["Tier"].dropna().unique().tolist())
    with filter_columns[0]:
        tier = st.selectbox("Tier", tier_options)

    with filter_columns[1]:
        search = st.text_input("Search accounts", placeholder="Type an account name")

    with filter_columns[2]:
        minimum_score = st.slider("Minimum score", 0, 100, 0)

    filtered = ranked.copy()
    if tier != "All":
        filtered = filtered[filtered["Tier"] == tier]
    if search.strip():
        filtered = filtered[
            filtered["Account Name"].str.contains(search.strip(), case=False, na=False)
            | filtered["Normalized Account Name"].str.contains(search.strip(), case=False, na=False)
        ]
    filtered = filtered[filtered["Score"].ge(minimum_score)]
    return filtered


def uploaded_file_signature(epic_file, ehr_file) -> tuple[str, int, str, int]:
    return (
        epic_file.name,
        int(getattr(epic_file, "size", 0) or 0),
        ehr_file.name,
        int(getattr(ehr_file, "size", 0) or 0),
    )


def clear_previous_results() -> None:
    for key in ["ranked_accounts", "summary", "excel_bytes"]:
        st.session_state.pop(key, None)


def main() -> None:
    st.set_page_config(
        page_title="AuthX Epic Targeting Engine",
        layout="wide",
    )

    st.title("AuthX Epic Targeting Engine")

    st.write(
        "Upload the EPIC organization list and Health Systems by EHR workbook to rank target accounts."
    )

    with st.sidebar:
        st.header("Uploads")
        epic_file = st.file_uploader(
            "EPIC Organization list.xlsx",
            type=["xlsx"],
            key="epic_file",
        )
        ehr_file = st.file_uploader(
            "Health Systems by EHR.xlsx",
            type=["xlsx"],
            key="ehr_file",
        )

        st.divider()
        st.caption("The app reads every worksheet from both uploaded files.")

    if not epic_file or not ehr_file:
        st.info("Upload both Excel files to generate the ranked account list.")
        return

    signature = uploaded_file_signature(epic_file, ehr_file)
    if st.session_state.get("uploaded_file_signature") != signature:
        clear_previous_results()
        st.session_state["uploaded_file_signature"] = signature

    st.success(f"Files uploaded: {epic_file.name} and {ehr_file.name}")

    if st.button("Generate ranked accounts", type="primary"):
        with st.spinner("Reading worksheets, deduplicating accounts, and scoring targets..."):
            try:
                ranked_accounts, summary = build_ranked_accounts(epic_file, ehr_file)
                excel_bytes = dataframe_to_excel_bytes(ranked_accounts, summary)
            except Exception as exc:
                clear_previous_results()
                st.error(f"Unable to process the uploaded files: {exc}")
                return

        st.session_state["ranked_accounts"] = ranked_accounts
        st.session_state["summary"] = summary
        st.session_state["excel_bytes"] = excel_bytes
        st.success(f"Generated ranked list with {summary['total_accounts']:,} accounts.")

    if "ranked_accounts" not in st.session_state:
        st.info("Both files are uploaded. Click Generate ranked accounts to create the output.")
        return

    ranked_accounts = st.session_state["ranked_accounts"]
    summary = st.session_state["summary"]
    excel_bytes = st.session_state["excel_bytes"]

    render_summary(summary)

    st.subheader("Ranked Accounts")
    filtered_accounts = render_filters(ranked_accounts)

    st.dataframe(
        filtered_accounts,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Score": st.column_config.ProgressColumn(
                "Score",
                min_value=0,
                max_value=100,
                format="%d",
            ),
            "Epic Customer": st.column_config.CheckboxColumn("Epic"),
            "New Epic System": st.column_config.CheckboxColumn("New Epic"),
            "Imprivata Customer": st.column_config.CheckboxColumn("Imprivata"),
            "Exclude Flag": st.column_config.CheckboxColumn("Exclude"),
        },
    )

    date_stamp = datetime.now().strftime("%Y%m%d")

    st.download_button(
        label="Download ranked Excel file",
        data=excel_bytes,
        file_name=f"authx_epic_targeting_ranked_{date_stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


main()
