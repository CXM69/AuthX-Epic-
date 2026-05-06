from __future__ import annotations

from datetime import datetime

import pandas as pd
import streamlit as st

from scoring import (
    build_ranked_accounts_from_sources,
    dataframe_to_excel_bytes,
    infer_source_type,
    source_configs_from_uploads,
)


def metric_card(label: str, value: int) -> None:
    st.metric(label=label, value=f"{value:,}")


def clear_previous_results() -> None:
    for key in ["ranked_accounts", "summary", "excel_bytes", "file_signature"]:
        st.session_state.pop(key, None)


def file_signature(uploaded_files) -> tuple[tuple[str, int], ...]:
    return tuple(
        (uploaded_file.name, int(getattr(uploaded_file, "size", 0) or 0))
        for uploaded_file in uploaded_files
    )


def render_uploaded_files(uploaded_files) -> None:
    rows = [
        {
            "File": uploaded_file.name,
            "Size KB": round((int(getattr(uploaded_file, "size", 0) or 0) / 1024), 1),
            "Detected Source Type": infer_source_type(uploaded_file.name),
        }
        for uploaded_file in uploaded_files
    ]
    st.dataframe(
        pd.DataFrame(rows),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Size KB": st.column_config.NumberColumn("Size KB", format="%.1f"),
        },
    )


def render_summary(summary: dict[str, int]) -> None:
    first_row = st.columns(4)
    with first_row[0]:
        metric_card("Uploaded Workbooks", summary.get("uploaded_workbooks", 0))
    with first_row[1]:
        metric_card("Total Accounts", summary["total_accounts"])
    with first_row[2]:
        metric_card("Tier 1", summary["tier_1"])
    with first_row[3]:
        metric_card("Tier 2", summary["tier_2"])

    second_row = st.columns(4)
    with second_row[0]:
        metric_card("Tier 3", summary["tier_3"])
    with second_row[1]:
        metric_card("New Epic Systems", summary["new_epic_systems"])
    with second_row[2]:
        metric_card("Existing Epic", summary["existing_epic_customers"])
    with second_row[3]:
        metric_card("Imprivata", summary["imprivata_customers"])

    third_row = st.columns(4)
    with third_row[0]:
        metric_card("Large Systems", summary["large_health_systems"])
    with third_row[1]:
        metric_card("Migration Opps", summary["migration_opportunities"])
    with third_row[2]:
        metric_card("Processed Workbooks", summary.get("processed_workbooks", 0))
    with third_row[3]:
        metric_card("Skipped Workbooks", summary.get("skipped_workbooks", 0))


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
        "Upload one or more Excel workbooks to identify the best healthcare targets for AuthX front-door authentication."
    )

    uploaded_files = st.file_uploader(
        "Upload Excel workbooks",
        type=["xlsx"],
        accept_multiple_files=True,
        key="uploaded_files",
        help="Use any account, EHR, Epic, Imprivata, security, or workflow spreadsheet. Every worksheet is read.",
    )
    st.caption(
        "The original EPIC Organization list and Health Systems by EHR workbooks are supported, but they are not required."
    )

    if not uploaded_files:
        clear_previous_results()
        st.info("Upload at least one Excel workbook to run the targeting engine.")
        return

    st.subheader("Uploaded Workbooks")
    render_uploaded_files(uploaded_files)

    current_signature = file_signature(uploaded_files)
    if st.session_state.get("file_signature") != current_signature:
        clear_previous_results()
        st.session_state["file_signature"] = current_signature

    st.success(f"{len(uploaded_files)} workbook(s) ready for scoring.")

    if st.button("Generate ranked list", type="primary"):
        with st.spinner("Reading tabs, deduplicating accounts, scoring targets, and building role strategy..."):
            try:
                sources = source_configs_from_uploads(uploaded_files)
                ranked_accounts, summary = build_ranked_accounts_from_sources(sources)
                excel_bytes = dataframe_to_excel_bytes(ranked_accounts, summary)
            except Exception as exc:
                clear_previous_results()
                st.error(f"Unable to process the uploaded files: {exc}")
                return

        st.session_state["ranked_accounts"] = ranked_accounts
        st.session_state["summary"] = summary
        st.session_state["excel_bytes"] = excel_bytes

    if "ranked_accounts" not in st.session_state:
        st.info("Click Generate ranked list to create the ranked target list.")
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
