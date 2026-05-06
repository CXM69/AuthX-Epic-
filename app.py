from __future__ import annotations

from datetime import datetime

import streamlit as st

from scoring import build_ranked_accounts, dataframe_to_excel_bytes


def metric_card(label: str, value: int) -> None:
    st.metric(label=label, value=f"{value:,}")


def clear_previous_results() -> None:
    for key in ["ranked_accounts", "summary", "excel_bytes", "file_signature"]:
        st.session_state.pop(key, None)


def file_signature(epic_file, ehr_file) -> tuple[str, int, str, int]:
    return (
        epic_file.name,
        int(getattr(epic_file, "size", 0) or 0),
        ehr_file.name,
        int(getattr(ehr_file, "size", 0) or 0),
    )


def render_summary(summary: dict[str, int]) -> None:
    first_row = st.columns(4)
    with first_row[0]:
        metric_card("Total Accounts", summary["total_accounts"])
    with first_row[1]:
        metric_card("Tier 1", summary["tier_1"])
    with first_row[2]:
        metric_card("Tier 2", summary["tier_2"])
    with first_row[3]:
        metric_card("New Epic Systems", summary["new_epic_systems"])

    second_row = st.columns(4)
    with second_row[0]:
        metric_card("Existing Epic", summary["existing_epic_customers"])
    with second_row[1]:
        metric_card("Imprivata", summary["imprivata_customers"])
    with second_row[2]:
        metric_card("Large Systems", summary["large_health_systems"])
    with second_row[3]:
        metric_card("Migration Opps", summary["migration_opportunities"])


def render_top_targets(ranked_accounts) -> None:
    preview_columns = [
        "Account Name",
        "Tier",
        "AuthX Score",
        "Epic Status",
        "Best Role to Pursue First",
        "Best Pitch",
        "Next Sales Action",
        "Confidence",
    ]
    st.dataframe(
        ranked_accounts[preview_columns].head(25),
        use_container_width=True,
        hide_index=True,
        column_config={
            "AuthX Score": st.column_config.ProgressColumn(
                "AuthX Score",
                min_value=0,
                max_value=100,
                format="%d",
            )
        },
    )


def main() -> None:
    st.set_page_config(
        page_title="AuthX Epic Front-Door Targeting Engine",
        layout="wide",
    )

    st.title("AuthX Epic Front-Door Targeting Engine")
    st.write(
        "Upload the Epic organization list and EHR workbook to identify the best front-door authentication targets."
    )

    with st.sidebar:
        st.header("Upload Workbooks")
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
        st.caption("Every worksheet in both workbooks will be read.")

    if not epic_file or not ehr_file:
        st.info("Upload both Excel workbooks to run the targeting engine.")
        return

    current_signature = file_signature(epic_file, ehr_file)
    if st.session_state.get("file_signature") != current_signature:
        clear_previous_results()
        st.session_state["file_signature"] = current_signature

    st.success(f"Uploaded: {epic_file.name} and {ehr_file.name}")

    if st.button("Run scoring", type="primary"):
        with st.spinner("Reading tabs, deduplicating accounts, scoring targets, and building role strategy..."):
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

    if "ranked_accounts" not in st.session_state:
        st.info("Click Run scoring to generate the ranked target list.")
        return

    ranked_accounts = st.session_state["ranked_accounts"]
    summary = st.session_state["summary"]
    excel_bytes = st.session_state["excel_bytes"]

    render_summary(summary)

    st.subheader("Top Target Preview")
    render_top_targets(ranked_accounts)

    date_stamp = datetime.now().strftime("%Y%m%d")
    st.download_button(
        label="Download ranked Excel file",
        data=excel_bytes,
        file_name=f"authx_epic_front_door_targets_{date_stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


main()
