import io
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Grundeinstellungen
# ---------------------------------------------------------------------------

DAY_NAMES = {
    1: "Montag",
    2: "Dienstag",
    3: "Mittwoch",
    4: "Donnerstag",
    5: "Freitag",
    6: "Samstag",
}

DAY_SHORT = {
    1: "Mo",
    2: "Di",
    3: "Mi",
    4: "Do",
    5: "Fr",
    6: "Sa",
}

# Fallback-Positionen, falls Spaltenüberschriften nicht erkannt werden.
# Tourenplanung: A CSB, B SAP, C Name, D Straße, E PLZ, F Ort, G Mo ... L Sa
DAY_COLUMNS_TOUR = {
    1: 6,
    2: 7,
    3: 8,
    4: 9,
    5: 10,
    6: 11,
}

SAP_COL_INDEX = 0
SAP_DAY_COL_INDEX = 6
TOUR_SAP_COL_INDEX = 1

# Nur diese Blätter der Tourenplanung bilden den Soll-Stand.
TOUR_SHEET_CANDIDATES = [
    "DIREKT",
    "MK",
    "HUPA_NMS",
    "HUPA_MALCHOW",
]

DAY_COLUMN_CANDIDATES = {
    1: ["mo", "montag"],
    2: ["die", "di", "dienstag"],
    3: ["mitt", "mit", "mi", "mittwoch"],
    4: ["don", "do", "donnerstag"],
    5: ["fr", "frei", "freitag"],
    6: ["sam", "sa", "samstag"],
}


# ---------------------------------------------------------------------------
# Helfer
# ---------------------------------------------------------------------------


def value_to_clean_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_sap_series(series: pd.Series) -> pd.Series:
    """Normalisiert SAP-Nummern. Aus 12345.0 wird 12345."""
    if series.empty:
        return series.astype(str)

    result = series.copy()
    numeric = pd.to_numeric(result, errors="coerce")
    is_int = numeric.notna() & (numeric == numeric.round())

    out = result.astype(str)
    out = out.where(~is_int, numeric.where(is_int).astype("Int64").astype(str))
    out = out.str.strip()
    return out.replace({"nan": "", "<NA>": "", "None": ""})


def normalize_header_name(value) -> str:
    text = "" if value is None or pd.isna(value) else str(value)
    text = text.strip().lower()
    text = (
        text.replace("ä", "ae")
        .replace("ö", "oe")
        .replace("ü", "ue")
        .replace("ß", "ss")
    )
    return "".join(ch for ch in text if ch.isalnum())


def normalized_candidates(values: List[str]) -> List[str]:
    return [normalize_header_name(value) for value in values]


def pick_first_matching_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    candidate_set = set(candidates)
    for column in columns:
        if normalize_header_name(column) in candidate_set:
            return column
    return None


def pick_column_by_name_or_position(
    columns: List[str],
    candidates: List[str],
    fallback_index: Optional[int] = None,
) -> Optional[str]:
    found = pick_first_matching_column(columns, normalized_candidates(candidates))
    if found is not None:
        return found
    if fallback_index is not None and len(columns) > fallback_index:
        return columns[fallback_index]
    return None


def make_unique_columns(raw_columns: List[object]) -> List[str]:
    result: List[str] = []
    seen: Dict[str, int] = {}
    for index, value in enumerate(raw_columns, start=1):
        name = value_to_clean_text(value) or f"Spalte_{index}"
        count = seen.get(name, 0) + 1
        seen[name] = count
        if count > 1:
            name = f"{name}_{count}"
        result.append(name)
    return result


def read_excel_with_detected_header(excel: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    """Erkennt die Kopfzeile anhand von SAP und mindestens zwei Tages-Spalten."""
    raw = pd.read_excel(excel, sheet_name=sheet_name, header=None, dtype=object)
    if raw.empty:
        return pd.DataFrame()

    header_row: Optional[int] = None
    max_scan_rows = min(len(raw), 25)
    day_names_flat = {
        normalize_header_name(candidate)
        for values in DAY_COLUMN_CANDIDATES.values()
        for candidate in values
    }

    for row_index in range(max_scan_rows):
        values = [normalize_header_name(value) for value in raw.iloc[row_index].tolist()]
        value_set = set(values)
        has_sap = "sap" in value_set or "sapnummer" in value_set or "sapnr" in value_set
        day_hits = sum(1 for value in values if value in day_names_flat)
        if has_sap and day_hits >= 2:
            header_row = row_index
            break

    if header_row is None:
        return pd.read_excel(excel, sheet_name=sheet_name, header=0, dtype=object)

    df = raw.iloc[header_row + 1:].copy()
    df.columns = make_unique_columns(raw.iloc[header_row].tolist())
    return df.dropna(how="all").reset_index(drop=True)


def select_tour_sheet_names(excel: pd.ExcelFile) -> Tuple[List[str], List[str]]:
    available = {normalize_header_name(name): name for name in excel.sheet_names}
    selected: List[str] = []
    missing: List[str] = []

    for expected in TOUR_SHEET_CANDIDATES:
        real_name = available.get(normalize_header_name(expected))
        if real_name:
            selected.append(real_name)
        else:
            missing.append(expected)

    if selected:
        return selected, missing

    # Falls Blattnamen stark abweichen: die ersten vier Blätter als Fallback.
    return excel.sheet_names[:4], TOUR_SHEET_CANDIDATES


def day_value_is_set(value) -> bool:
    if value is None or pd.isna(value):
        return False
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "<na>", "-", "--"}:
        return False
    number = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.notna(number) and float(number) == 0:
        return False
    return True


def normalize_day_code_series(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    text = series.astype(str).map(normalize_header_name)
    text_map = {
        "1": 1, "mo": 1, "montag": 1,
        "2": 2, "di": 2, "die": 2, "dienstag": 2,
        "3": 3, "mi": 3, "mitt": 3, "mit": 3, "mittwoch": 3,
        "4": 4, "do": 4, "don": 4, "donnerstag": 4,
        "5": 5, "fr": 5, "frei": 5, "freitag": 5,
        "6": 6, "sa": 6, "sam": 6, "samstag": 6,
    }
    mapped = text.map(text_map)
    return numeric.where(numeric.notna(), mapped)


def days_to_text(days: Set[int] | List[int]) -> str:
    return ", ".join(DAY_SHORT[d] for d in sorted(days))


def merge_customer_info(base: Dict[str, Dict[str, str]], sap: str, info: Dict[str, str]) -> None:
    target = base.setdefault(sap, {"name": "", "strasse": "", "ort": ""})
    for key in ["name", "strasse", "ort"]:
        if not target.get(key) and info.get(key):
            target[key] = info[key]


# ---------------------------------------------------------------------------
# Dateien lesen
# ---------------------------------------------------------------------------


def read_sap_file(uploaded_file) -> Tuple[Dict[str, Set[int]], str, int]:
    """SAP Ist-Stand: Fallback A = SAP, G = Liefertag."""
    excel = pd.ExcelFile(uploaded_file)
    sheet_name = excel.sheet_names[0]
    df = read_excel_with_detected_header(excel, sheet_name)

    if df.empty:
        return {}, sheet_name, 0

    columns = list(df.columns)
    sap_column = pick_column_by_name_or_position(
        columns,
        ["SAP", "SAP Nummer", "SAP-Nr", "SAP Nr", "Kundennummer", "Kunden Nummer"],
        SAP_COL_INDEX,
    )
    day_column = pick_column_by_name_or_position(
        columns,
        ["Liefertag", "Liefer Tag", "LT", "Tag", "Liefertag Code", "Liefertagcode"],
        SAP_DAY_COL_INDEX,
    )

    if sap_column is None or day_column is None:
        return {}, sheet_name, 0

    work = df[[sap_column, day_column]].copy()
    work.columns = ["sap", "tag"]
    work["sap"] = normalize_sap_series(work["sap"])
    work["tag_num"] = normalize_day_code_series(work["tag"])

    mask = (
        work["sap"].ne("")
        & work["tag_num"].notna()
        & work["tag_num"].between(1, 6, inclusive="both")
    )
    filtered = work.loc[mask, ["sap", "tag_num"]].copy()
    filtered["tag_int"] = filtered["tag_num"].astype(int)

    days_by_sap: Dict[str, Set[int]] = filtered.groupby("sap")["tag_int"].agg(set).to_dict()
    return days_by_sap, sheet_name, len(filtered)


def read_tourenplanung(
    uploaded_file,
) -> Tuple[pd.DataFrame, List[str], List[str], Dict[str, Dict[str, str]]]:
    """Tourenplanung = verbindlicher Soll-Stand."""
    excel = pd.ExcelFile(uploaded_file)
    sheet_names, missing_sheet_names = select_tour_sheet_names(excel)

    frames: List[pd.DataFrame] = []
    customer_info: Dict[str, Dict[str, str]] = {}

    for sheet_name in sheet_names:
        df = read_excel_with_detected_header(excel, sheet_name)
        if df.empty:
            continue

        columns = list(df.columns)
        sap_column = pick_column_by_name_or_position(
            columns,
            ["SAP", "SAP Nummer", "SAP-Nr", "SAP Nr", "Kundennummer", "Kunden Nummer"],
            TOUR_SAP_COL_INDEX,
        )
        if sap_column is None:
            continue

        name_column = pick_column_by_name_or_position(
            columns,
            ["Name", "Kundenname", "Marktname", "Kunde", "Bezeichnung", "Filialname"],
            2,
        )
        strasse_column = pick_column_by_name_or_position(
            columns,
            ["Strasse", "Straße", "Str", "Anschrift", "Adresse", "Strassenname", "Straßenname", "Strasse Hausnummer", "Straße Hausnummer"],
            3,
        )
        plz_column = pick_column_by_name_or_position(columns, ["Plz", "PLZ", "Postleitzahl"], 4)
        ort_column = pick_column_by_name_or_position(columns, ["Ort", "Stadt", "Plz Ort", "PLZ Ort", "Ortname"], 5)

        rename_map = {sap_column: "sap"}
        for day_num, col_index in DAY_COLUMNS_TOUR.items():
            day_column = pick_column_by_name_or_position(
                columns,
                DAY_COLUMN_CANDIDATES[day_num],
                col_index,
            )
            if day_column and day_column != sap_column:
                rename_map[day_column] = f"tag_{day_num}"

        if name_column and name_column != sap_column:
            rename_map[name_column] = "name"
        if strasse_column and strasse_column != sap_column:
            rename_map[strasse_column] = "strasse"
        if ort_column and ort_column != sap_column:
            rename_map[ort_column] = "ort"
        if plz_column and plz_column != sap_column:
            rename_map[plz_column] = "plz"

        work = df.rename(columns=rename_map).copy()
        work["sap"] = normalize_sap_series(work["sap"])
        work = work[work["sap"].ne("")].copy()
        if work.empty:
            continue

        for column in ["name", "strasse", "ort", "plz"]:
            if column not in work.columns:
                work[column] = ""

        info_df = work[["sap", "name", "strasse", "ort", "plz"]].copy()
        for column in ["name", "strasse", "ort", "plz"]:
            info_df[column] = info_df[column].map(value_to_clean_text)
        info_df["ort_kombi"] = info_df.apply(
            lambda row: " ".join(v for v in [row["plz"], row["ort"]] if v).strip(),
            axis=1,
        )

        for _, row in info_df.iterrows():
            merge_customer_info(
                customer_info,
                row["sap"],
                {
                    "name": row["name"],
                    "strasse": row["strasse"],
                    "ort": row["ort_kombi"] or row["ort"],
                },
            )

        day_value_columns = [
            f"tag_{d}" for d in DAY_COLUMNS_TOUR if f"tag_{d}" in work.columns
        ]
        if not day_value_columns:
            continue

        work["blatt"] = sheet_name
        long = work.melt(
            id_vars=["sap", "blatt"],
            value_vars=day_value_columns,
            var_name="tag_col",
            value_name="wert",
        )
        long["tag_num"] = long["tag_col"].str.replace("tag_", "", regex=False).astype(int)
        long["wert_gesetzt"] = long["wert"].map(day_value_is_set)
        long = long[long["sap"].ne("") & long["wert_gesetzt"]]
        frames.append(long[["sap", "blatt", "tag_num"]])

    if not frames:
        empty = pd.DataFrame(columns=["sap", "blatt", "tag_num"])
        return empty, sheet_names, missing_sheet_names, customer_info

    return pd.concat(frames, ignore_index=True), sheet_names, missing_sheet_names, customer_info


# ---------------------------------------------------------------------------
# EIN Vergleich: SAP gegen Tourenplanung
# ---------------------------------------------------------------------------


RESULT_COLUMNS = [
    "Blatt",
    "SAP Nummer",
    "Name",
    "Straße",
    "Ort",
    "Tourenplanung (Soll)",
    "SAP (Ist)",
    "SAP-Abweichung",
]


def build_sap_differences(
    tour_df: pd.DataFrame,
    days_by_sap: Dict[str, Set[int]],
    customer_info: Dict[str, Dict[str, str]],
) -> pd.DataFrame:
    """
    Die Tourenplanung hat immer Recht.
    Es werden ausschließlich SAP-Abweichungen für Kunden aus der Tourenplanung gezeigt.
    """
    if tour_df.empty:
        return pd.DataFrame(columns=RESULT_COLUMNS)

    expected_by_sap: Dict[str, Set[int]] = tour_df.groupby("sap")["tag_num"].agg(set).to_dict()
    sheets_by_sap: Dict[str, str] = tour_df.groupby("sap")["blatt"].agg(
        lambda x: ", ".join(sorted(set(map(str, x))))
    ).to_dict()

    rows: List[dict] = []

    for sap, expected_days in expected_by_sap.items():
        actual_days = days_by_sap.get(sap, set())
        missing_in_sap = sorted(expected_days - actual_days)
        extra_in_sap = sorted(actual_days - expected_days)

        if not missing_in_sap and not extra_in_sap:
            continue

        parts: List[str] = []
        if missing_in_sap:
            parts.append(f"Fehlt in SAP: {days_to_text(missing_in_sap)}")
        if extra_in_sap:
            parts.append(f"Zusätzlich in SAP: {days_to_text(extra_in_sap)}")

        info = customer_info.get(sap, {})
        rows.append({
            "Blatt": sheets_by_sap.get(sap, ""),
            "SAP Nummer": sap,
            "Name": info.get("name", ""),
            "Straße": info.get("strasse", ""),
            "Ort": info.get("ort", ""),
            "Tourenplanung (Soll)": days_to_text(expected_days) or "–",
            "SAP (Ist)": days_to_text(actual_days) or "–",
            "SAP-Abweichung": " | ".join(parts),
            "_sort": int(sap) if str(sap).isdigit() else 9_999_999_999,
        })

    if not rows:
        return pd.DataFrame(columns=RESULT_COLUMNS)

    result = pd.DataFrame(rows)
    result = result.sort_values(["Blatt", "_sort"]).reset_index(drop=True)
    return result[RESULT_COLUMNS]


def filter_result(df: pd.DataFrame, suche: str, blatt: str) -> pd.DataFrame:
    if df.empty:
        return df

    work = df
    if blatt != "Alle":
        work = work[work["Blatt"].astype(str).str.contains(blatt, na=False, regex=False)]

    if suche.strip():
        term = suche.strip().lower()
        columns = ["SAP Nummer", "Name", "Straße", "Ort", "SAP-Abweichung"]
        mask = pd.Series(False, index=work.index)
        for column in columns:
            mask = mask | work[column].astype(str).str.lower().str.contains(term, na=False)
        work = work[mask]

    return work


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------


def build_excel(result: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="SAP Abweichungen", na_rep="")
        ws = writer.sheets["SAP Abweichungen"]

        header_fill = PatternFill(start_color="FF2F3A4A", end_color="FF2F3A4A", fill_type="solid")
        zebra_fill = PatternFill(start_color="FFF4F6F8", end_color="FFF4F6F8", fill_type="solid")
        warning_fill = PatternFill(start_color="FFFFF0E5", end_color="FFFFF0E5", fill_type="solid")
        thin = Side(style="thin", color="FFD7DEE8")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
            cell.alignment = Alignment(vertical="center")
            cell.border = border

        for row_idx in range(2, len(result) + 2):
            for col_idx in range(1, len(RESULT_COLUMNS) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = Font(name="Calibri", size=11)
                cell.alignment = Alignment(vertical="center", wrap_text=False)
                cell.border = border
                if row_idx % 2 == 1:
                    cell.fill = zebra_fill

            # Abweichung optisch hervorheben
            ws.cell(row=row_idx, column=RESULT_COLUMNS.index("SAP-Abweichung") + 1).fill = warning_fill

        width_hints = {
            "Blatt": 18,
            "SAP Nummer": 12,
            "Name": 32,
            "Straße": 26,
            "Ort": 26,
            "Tourenplanung (Soll)": 22,
            "SAP (Ist)": 22,
            "SAP-Abweichung": 38,
        }
        for idx, column in enumerate(RESULT_COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = width_hints[column]

        ws.freeze_panes = "A2"
        if not result.empty:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(RESULT_COLUMNS))}{len(result) + 1}"
        ws.sheet_view.showGridLines = False
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

    return output.getvalue()


# ---------------------------------------------------------------------------
# Streamlit-Oberfläche
# ---------------------------------------------------------------------------


st.set_page_config(page_title="SAP-Abgleich zur Tourenplanung", layout="wide")

st.markdown(
    """
    <style>
        .block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 1400px; }
        [data-testid="stMetric"] { background: #f7f8fa; border: 1px solid #e3e7ee; border-radius: 14px; padding: 14px 16px; }
        [data-testid="stFileUploader"] section { border: 1px dashed #b9c0cc; border-radius: 14px; background: #fafbfc; }
        div.stButton > button { border-radius: 12px; height: 3rem; font-weight: 700; }
        .hint { color: #586174; font-size: 0.94rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("SAP-Abgleich zur Tourenplanung")
st.markdown(
    """
    <div class="hint">
    <b>Grundlage ist ausschließlich die Tourenplanung.</b><br>
    Die Tourenplanung ist der Soll-Stand. Angezeigt wird nur, wo <b>SAP davon abweicht</b>.
    Kunden, die nur in SAP stehen und nicht in der Tourenplanung vorkommen, werden bewusst ignoriert.
    </div>
    """,
    unsafe_allow_html=True,
)

st.divider()

upload_left, upload_right = st.columns(2)
with upload_left:
    tourenplanung_datei = st.file_uploader(
        "1. Tourenplanung (Soll)",
        help="Diese Datei hat Vorrang und definiert die richtigen Liefertage.",
        type=["xlsx", "xlsm", "xls"],
        key="tourenplanung_datei",
    )

with upload_right:
    sap_datei = st.file_uploader(
        "2. SAP-Datei (Ist)",
        help="SAP wird ausschließlich gegen die Tourenplanung geprüft.",
        type=["xlsx", "xlsm", "xls"],
        key="sap_datei",
    )

run = st.button("SAP-Abweichungen prüfen", type="primary", use_container_width=True)

if run:
    if not sap_datei or not tourenplanung_datei:
        st.error("Bitte Tourenplanung und SAP-Datei hochladen.")
        st.stop()

    try:
        tour_df, tour_sheets, missing_tour_sheets, customer_info = read_tourenplanung(tourenplanung_datei)
        days_by_sap, sap_sheet, sap_rows = read_sap_file(sap_datei)

        if tour_df.empty:
            st.error("In der Tourenplanung wurden keine gültigen Liefertage erkannt.")
            st.stop()
        if sap_rows == 0:
            st.warning("In der SAP-Datei wurden keine gültigen Liefertage erkannt.")
        if missing_tour_sheets:
            st.warning("Nicht gefundene Touren-Blätter: " + ", ".join(missing_tour_sheets))

        differences = build_sap_differences(tour_df, days_by_sap, customer_info)
        excel_bytes = build_excel(differences)

        st.session_state["sap_compare_result"] = {
            "differences": differences,
            "excel_bytes": excel_bytes,
            "tour_sheets": tour_sheets,
            "sap_sheet": sap_sheet,
        }

    except Exception as exc:
        import traceback
        st.error(f"Fehler beim Verarbeiten der Dateien: {exc}")
        with st.expander("Technische Details", expanded=False):
            st.code(traceback.format_exc(), language="python")
        st.session_state.pop("sap_compare_result", None)


result = st.session_state.get("sap_compare_result")
if result:
    differences = result["differences"]

    st.divider()
    header_left, header_right = st.columns([3, 1])
    with header_left:
        st.subheader("SAP-Abweichungen")
        st.caption(
            "Nur Abweichungen von der Tourenplanung werden angezeigt. "
            "Die Tourenplanung ist immer der Soll-Stand."
        )
    with header_right:
        st.download_button(
            "Excel herunterladen",
            data=result["excel_bytes"],
            file_name="SAP_Abweichungen_zur_Tourenplanung.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.metric("SAP-Abweichungen", len(differences))

    if differences.empty:
        st.success("Keine Abweichungen: SAP stimmt mit der Tourenplanung überein.")
    else:
        f1, f2 = st.columns([1, 2])
        with f1:
            blatt_options = ["Alle"] + sorted(
                set(
                    sheet
                    for cell in differences["Blatt"].dropna().astype(str)
                    for sheet in [s.strip() for s in cell.split(",")]
                    if sheet
                )
            )
            blatt = st.selectbox("Blatt", blatt_options)
        with f2:
            suche = st.text_input(
                "Suchen",
                placeholder="SAP Nummer, Name, Straße, Ort oder Abweichung",
            )

        filtered = filter_result(differences, suche, blatt)
        st.caption(f"{len(filtered)} Abweichungen")
        st.dataframe(
            filtered,
            use_container_width=True,
            hide_index=True,
            column_config={
                "SAP Nummer": st.column_config.TextColumn(width="small"),
                "Blatt": st.column_config.TextColumn(width="small"),
                "Tourenplanung (Soll)": st.column_config.TextColumn(width="medium"),
                "SAP (Ist)": st.column_config.TextColumn(width="medium"),
                "SAP-Abweichung": st.column_config.TextColumn(width="large"),
            },
        )
