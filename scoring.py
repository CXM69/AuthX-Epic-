from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import re
from typing import BinaryIO

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


ACCOUNT_COLUMN_CANDIDATES = [
    "account name",
    "account",
    "organization",
    "organization name",
    "org name",
    "health system",
    "health system name",
    "system",
    "system name",
    "customer",
    "customer name",
    "name",
]

EHR_COLUMN_CANDIDATES = [
    "ehr",
    "emr",
    "vendor",
    "ehr vendor",
    "emr vendor",
    "platform",
    "source ehr",
]

STATUS_COLUMN_CANDIDATES = [
    "status",
    "customer status",
    "relationship",
    "stage",
]

STATE_COLUMN_CANDIDATES = ["state", "st", "region"]

EXCLUDE_TERMS = {
    "test",
    "demo",
    "sample",
    "training",
    "unknown",
    "n/a",
    "na",
    "none",
}

LEGAL_SUFFIX_PATTERN = re.compile(
    r"\b(inc|incorporated|llc|l\.l\.c|ltd|limited|corp|corporation|co|company|plc)\b",
    re.IGNORECASE,
)

ORG_WORD_PATTERN = re.compile(
    r"\b(the|healthcare|health care|health system|health systems|hospital system|hospitals|hospital|medical center|medical ctr|clinic|clinics|network)\b",
    re.IGNORECASE,
)

ORGANIZATION_SIGNAL_PATTERN = re.compile(
    r"\b(health|hospital|medical|clinic|system|systems|care|center|centre|regional|memorial|university|children|county|network|partners|group|foundation)\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class SourceConfig:
    label: str
    uploaded_file: BinaryIO


def normalize_column_name(column: object) -> str:
    text = str(column).strip().lower()
    text = re.sub(r"[\s_\-/]+", " ", text)
    return re.sub(r"[^a-z0-9 ]+", "", text).strip()


def normalize_account_name(value: object) -> str:
    if pd.isna(value):
        return ""

    text = str(value).strip().lower()
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = LEGAL_SUFFIX_PATTERN.sub(" ", text)
    text = ORG_WORD_PATTERN.sub(" ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def title_account_name(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def read_all_tabs(source: SourceConfig) -> pd.DataFrame:
    if hasattr(source.uploaded_file, "seek"):
        source.uploaded_file.seek(0)

    workbook = pd.ExcelFile(source.uploaded_file, engine="openpyxl")
    frames: list[pd.DataFrame] = []

    for sheet_name in workbook.sheet_names:
        frame = read_sheet_with_header_detection(workbook, sheet_name)
        frame = frame.dropna(how="all")
        if frame.empty:
            continue

        frame.columns = [str(column).strip() for column in frame.columns]
        frame["Source File"] = source.label
        frame["Source Tab"] = sheet_name
        frames.append(frame)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True, sort=False)


def read_sheet_with_header_detection(workbook: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None, dtype=object)
    raw = raw.dropna(axis=0, how="all").dropna(axis=1, how="all")
    if raw.empty:
        return pd.DataFrame()

    header_row = detect_header_row(raw)
    if header_row is None:
        frame = raw.copy()
        frame.columns = [f"Column {index + 1}" for index in range(len(frame.columns))]
        return frame

    header_values = raw.iloc[header_row].tolist()
    frame = raw.iloc[header_row + 1 :].copy()
    frame.columns = make_unique_columns(header_values)
    return frame.dropna(axis=0, how="all")


def detect_header_row(raw: pd.DataFrame, max_rows: int = 15) -> int | None:
    candidates = (
        ACCOUNT_COLUMN_CANDIDATES
        + EHR_COLUMN_CANDIDATES
        + STATUS_COLUMN_CANDIDATES
        + STATE_COLUMN_CANDIDATES
    )
    normalized_candidates = {normalize_column_name(candidate) for candidate in candidates}
    best_row: int | None = None
    best_score = 0

    for row_index in range(min(max_rows, len(raw))):
        values = [normalize_column_name(value) for value in raw.iloc[row_index].tolist()]
        values = [value for value in values if value and not value.startswith("unnamed")]
        if not values:
            continue

        score = 0
        for value in values:
            if value in normalized_candidates:
                score += 5
            elif any(candidate and candidate in value for candidate in normalized_candidates if len(candidate) > 2):
                score += 2

        if score > best_score:
            best_score = score
            best_row = row_index

    return best_row if best_score else None


def make_unique_columns(values: list[object]) -> list[str]:
    columns: list[str] = []
    counts: dict[str, int] = {}

    for index, value in enumerate(values):
        column = str(value).strip() if not pd.isna(value) else ""
        if not column or column.lower().startswith("unnamed"):
            column = f"Column {index + 1}"

        counts[column] = counts.get(column, 0) + 1
        if counts[column] > 1:
            column = f"{column} {counts[column]}"
        columns.append(column)

    return columns


def find_column(frame: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized_lookup = {
        normalize_column_name(column): column
        for column in frame.columns
        if not normalize_column_name(column).startswith("unnamed")
    }
    normalized_candidates = [normalize_column_name(candidate) for candidate in candidates]

    for normalized_candidate in normalized_candidates:
        if normalized_candidate in normalized_lookup:
            return normalized_lookup[normalized_candidate]

    for column in frame.columns:
        normalized = normalize_column_name(column)
        if normalized.startswith("unnamed"):
            continue
        normalized_tokens = set(normalized.split())
        for candidate in normalized_candidates:
            if len(candidate) <= 2 and candidate in normalized_tokens:
                return column
            if len(candidate) > 2 and candidate in normalized:
                return column

    return None


def infer_account_column(frame: pd.DataFrame) -> str | None:
    ignored_columns = {
        "source file",
        "source tab",
        "source type",
    }
    best_column: str | None = None
    best_score = 0.0

    for column in frame.columns:
        normalized_column = normalize_column_name(column)
        if normalized_column in ignored_columns:
            continue

        values = frame[column].dropna().astype(str).str.strip()
        values = values[
            values.ne("")
            & ~values.str.lower().isin({"nan", "none", "null", "n/a", "na"})
        ].head(300)
        if values.empty:
            continue

        useful_values = values[
            values.str.len().between(3, 120)
            & ~values.str.fullmatch(r"[-+]?\d+(\.\d+)?", na=False)
        ]
        if useful_values.empty:
            continue

        signal_hits = useful_values.str.contains(ORGANIZATION_SIGNAL_PATTERN, na=False).sum()
        uniqueness = useful_values.nunique() / max(len(useful_values), 1)
        average_length = min(useful_values.str.len().mean(), 60)
        score = (len(useful_values) * 2) + (signal_hits * 8) + (uniqueness * 20) + average_length

        if score > best_score:
            best_score = score
            best_column = column

    return best_column


def prepare_source(frame: pd.DataFrame, source_type: str) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(
            columns=[
                "Account Name",
                "Normalized Account Name",
                "EHR",
                "Status",
                "State",
                "Source File",
                "Source Tab",
                "Source Type",
            ]
        )

    account_column = find_column(frame, ACCOUNT_COLUMN_CANDIDATES)
    if account_column is None:
        account_column = infer_account_column(frame)
    if account_column is None:
        raise ValueError(f"Could not find a usable account/name column in {source_type}.")

    ehr_column = find_column(frame, EHR_COLUMN_CANDIDATES)
    status_column = find_column(frame, STATUS_COLUMN_CANDIDATES)
    state_column = find_column(frame, STATE_COLUMN_CANDIDATES)

    prepared = pd.DataFrame()
    prepared["Account Name"] = frame[account_column].map(title_account_name)
    prepared["Normalized Account Name"] = prepared["Account Name"].map(normalize_account_name)
    prepared["EHR"] = frame[ehr_column].fillna("").astype(str).str.strip() if ehr_column else ""
    prepared["Status"] = frame[status_column].fillna("").astype(str).str.strip() if status_column else ""
    prepared["State"] = frame[state_column].fillna("").astype(str).str.strip() if state_column else ""
    prepared["Source File"] = frame.get("Source File", "")
    prepared["Source Tab"] = frame.get("Source Tab", "")
    prepared["Source Type"] = source_type

    prepared = prepared[prepared["Normalized Account Name"].ne("")]
    prepared = prepared[
        ~prepared["Normalized Account Name"].isin(EXCLUDE_TERMS)
        & prepared["Account Name"].str.len().gt(1)
    ]
    return prepared.reset_index(drop=True)


def contains_any(series: pd.Series, terms: list[str]) -> bool:
    combined = " | ".join(series.dropna().astype(str).str.lower().unique())
    return any(term in combined for term in terms)


def collapse_values(series: pd.Series) -> str:
    values = [
        str(value).strip()
        for value in series.dropna().tolist()
        if str(value).strip() and str(value).strip().lower() not in {"nan", "none"}
    ]
    unique_values = list(dict.fromkeys(values))
    return "; ".join(unique_values[:8])


def score_account(row: pd.Series) -> int:
    score = 0

    if row["Epic Customer"]:
        score += 35
    if row["New Epic System"]:
        score += 30
    if row["Imprivata Customer"]:
        score += 20

    source_count = int(row.get("Source Count", 0) or 0)
    source_tab_count = int(row.get("Source Tab Count", 0) or 0)
    score += min(source_count * 5, 10)
    score += min(source_tab_count * 2, 5)

    if row.get("State"):
        score += 3

    if row["Exclude Flag"]:
        return 0

    return max(0, min(100, score))


def assign_priority_score(row: pd.Series) -> str:
    score = int(row["Fit Score"])
    if score >= 60:
        return "1"
    if score >= 40:
        return "2"
    return "Hold"


def build_ranked_accounts_from_sources(sources: list[SourceConfig]) -> tuple[pd.DataFrame, dict[str, int]]:
    if not sources:
        raise ValueError("Upload at least one Excel workbook.")

    prepared_sources: list[pd.DataFrame] = []

    for index, source in enumerate(sources, start=1):
        raw = read_all_tabs(source)
        prepared_sources.append(prepare_source(raw, f"Workbook {index}"))

    combined = pd.concat(prepared_sources, ignore_index=True, sort=False)

    if combined.empty:
        raise ValueError("No usable account rows were found in the uploaded workbooks.")

    combined["EHR Search Text"] = (
        combined["EHR"].fillna("").astype(str)
        + " "
        + combined["Status"].fillna("").astype(str)
        + " "
        + combined["Source File"].fillna("").astype(str)
        + " "
        + combined["Source Tab"].fillna("").astype(str)
        + " "
        + combined["Source Type"].fillna("").astype(str)
    ).str.lower()

    grouped = combined.groupby("Normalized Account Name", dropna=False)
    ranked = grouped.agg(
        Account_Name=("Account Name", "first"),
        Account_Name_Variants=("Account Name", collapse_values),
        State=("State", collapse_values),
        EHR=("EHR", collapse_values),
        Status=("Status", collapse_values),
        Sources=("Source File", collapse_values),
        Source_Tabs=("Source Tab", collapse_values),
        Source_Count=("Source File", "nunique"),
        Source_Tab_Count=("Source Tab", "nunique"),
        Source_Text=("EHR Search Text", collapse_values),
    ).reset_index()

    ranked = ranked.rename(
        columns={
            "Normalized Account Name": "Normalized Account Name",
            "Account_Name": "Account Name",
            "Account_Name_Variants": "Account Name Variants",
            "Source_Tabs": "Source Tabs",
            "Source_Count": "Source Count",
            "Source_Tab_Count": "Source Tab Count",
            "Source_Text": "Source Text",
        }
    )

    ranked["Epic Customer"] = ranked["Source Text"].str.contains("epic", case=False, na=False)
    ranked["New Epic System"] = ranked["Source Text"].str.contains(
        "new epic|implementation|implementing|migration|go live|go-live|installing|new system",
        case=False,
        regex=True,
        na=False,
    )
    ranked["Imprivata Customer"] = ranked["Source Text"].str.contains(
        "imprivata|onesign|one sign|confirm id|fairwarning",
        case=False,
        regex=True,
        na=False,
    )
    ranked["Exclude Flag"] = ranked["Source Text"].str.contains(
        "exclude|competitor|closed|inactive|do not contact|duplicate only",
        case=False,
        regex=True,
        na=False,
    )

    ranked["Fit Score"] = ranked.apply(score_account, axis=1)
    ranked["Priority Score"] = ranked.apply(assign_priority_score, axis=1)
    ranked["Priority Sort"] = ranked["Priority Score"].map({"1": 1, "2": 2, "Hold": 3}).fillna(3)

    ranked = ranked.sort_values(
        by=["Priority Sort", "Fit Score", "Epic Customer", "New Epic System", "Imprivata Customer", "Account Name"],
        ascending=[True, False, False, False, False, True],
    ).reset_index(drop=True)
    ranked.insert(0, "Rank", ranked.index + 1)

    output_columns = [
        "Rank",
        "Account Name",
        "Normalized Account Name",
        "Priority Score",
        "Fit Score",
        "Epic Customer",
        "New Epic System",
        "Imprivata Customer",
        "Exclude Flag",
        "State",
        "EHR",
        "Status",
        "Sources",
        "Source Tabs",
        "Source Count",
        "Source Tab Count",
        "Account Name Variants",
    ]
    ranked = ranked[output_columns]

    summary = {
        "total_accounts": int(len(ranked)),
        "priority_1": int((ranked["Priority Score"] == "1").sum()),
        "priority_2": int((ranked["Priority Score"] == "2").sum()),
        "hold": int((ranked["Priority Score"] == "Hold").sum()),
        "exclude": int(ranked["Exclude Flag"].sum()),
        "epic_customers": int(ranked["Epic Customer"].sum()),
        "new_epic_systems": int(ranked["New Epic System"].sum()),
        "imprivata_customers": int(ranked["Imprivata Customer"].sum()),
    }

    return ranked, summary


def build_ranked_accounts(
    primary_file: BinaryIO,
    secondary_file: BinaryIO,
    primary_label: str = "Workbook 1",
    secondary_label: str = "Workbook 2",
) -> tuple[pd.DataFrame, dict[str, int]]:
    return build_ranked_accounts_from_sources(
        [
            SourceConfig(primary_label, primary_file),
            SourceConfig(secondary_label, secondary_file),
        ]
    )


def dataframe_to_excel_bytes(ranked: pd.DataFrame, summary: dict[str, int]) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        ranked.to_excel(writer, index=False, sheet_name="Ranked Accounts")

        summary_frame = pd.DataFrame(
            [
                ("Total Accounts", summary["total_accounts"]),
                ("Priority 1", summary["priority_1"]),
                ("Priority 2", summary["priority_2"]),
                ("Hold", summary["hold"]),
                ("Excluded", summary["exclude"]),
                ("Epic Customers", summary["epic_customers"]),
                ("New Epic Systems", summary["new_epic_systems"]),
                ("Imprivata Customers", summary["imprivata_customers"]),
            ],
            columns=["Metric", "Value"],
        )
        summary_frame.to_excel(writer, index=False, sheet_name="Summary")

        workbook = writer.book
        ranked_sheet = workbook["Ranked Accounts"]
        summary_sheet = workbook["Summary"]

        header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)

        for sheet in [ranked_sheet, summary_sheet]:
            sheet.freeze_panes = "A2"
            for cell in sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
            for column_cells in sheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells)
                width = min(max(max_length + 2, 12), 42)
                sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = width

        ranked_sheet.auto_filter.ref = ranked_sheet.dimensions
        summary_sheet.auto_filter.ref = summary_sheet.dimensions

    output.seek(0)
    return output.read()
