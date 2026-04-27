import pandas as pd
from datetime import date, timedelta
from pathlib import Path
from io import BytesIO
import zipfile
import re
import unicodedata

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE_PATH = BASE_DIR / "template.xlsx"
DEFAULT_ADDITIONAL_DATA_PATH = BASE_DIR / "ADDITIONAL_DATA.xlsx"

# =========================================================
# HELPERS
# =========================================================

def next_weekday(d: date) -> date:
    nd = d + timedelta(days=1)
    while nd.weekday() >= 5:
        nd += timedelta(days=1)
    return nd


def safe_str(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def safe_float(value, default=0.0):
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass

    s = str(value).strip()
    if s == "":
        return default

    s = s.replace("\xa0", "")
    s = s.replace(" ", "")
    s = s.replace(",", ".")

    try:
        return float(s)
    except Exception:
        return default


def normalize_text(s):
    s = safe_str(s)
    s = s.replace("\xa0", " ")
    s = " ".join(s.split()).strip().casefold()
    return s


def normalize_text_loose(s):
    s = normalize_text(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s


def find_column_label(columns, *aliases):
    normalized = {
        normalize_text_loose(col): col
        for col in columns
    }
    for alias in aliases:
        match = normalized.get(normalize_text_loose(alias))
        if match is not None:
            return match
    return None


def get_value_by_column_or_index(row, fallback_index: int, *aliases):
    label = find_column_label(row.index, *aliases)
    if label is not None:
        return row[label]
    if 0 <= fallback_index < len(row):
        return row.iloc[fallback_index]
    return None


def build_duplicate_headers(raw_headers):
    counts = {}
    result = []

    for h in raw_headers:
        h = safe_str(h)
        if h == "":
            h = "Unnamed"

        if h not in counts:
            counts[h] = 0
            result.append(h)
        else:
            counts[h] += 1
            result.append(f"{h}_{counts[h]}")

    return result


def sanitize_filename(name: str) -> str:
    name = safe_str(name)
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


# =========================================================
# ROBUST XML PARSER FOR EXCEL 2003 XML
# =========================================================

def load_excel_2003_xml_root(file_obj):
    """
    Parse Excel 2003 XML / SpreadsheetML while tolerating:
    - broken encoding declarations
    - invalid control characters
    - non-standard characters
    - minor malformed XML

    Important fix:
    prefer UTF-8 with replacement before cp1250 fallback,
    so valid Croatian letters do not become mojibake like RaÄŤun.
    """
    from lxml import etree

    if hasattr(file_obj, "seek"):
        file_obj.seek(0)

    raw = file_obj.read()

    if isinstance(raw, str):
        raw = raw.encode("utf-8", errors="ignore")

    raw = raw.lstrip(b"\xef\xbb\xbf")

    # 1) Try strict UTF-8 first
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        # 2) If that fails, still prefer UTF-8 but keep going through bad bytes
        text = raw.decode("utf-8", errors="replace")

        # 3) Only if UTF-8 result looks very bad, fall back to cp1250
        #    This avoids turning "Račun" into "RaÄŤun"
        suspicious = text.count("�")
        if suspicious > 20:
            try:
                text_cp1250 = raw.decode("cp1250")
                # keep cp1250 only if it clearly looks better
                if text_cp1250.count("�") < suspicious:
                    text = text_cp1250
            except Exception:
                pass

    # Remove illegal XML 1.0 control characters
    text = re.sub(r"[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]", "", text)

    # Escape stray ampersands
    text = re.sub(
        r"&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9A-Fa-f]+;)",
        "&amp;",
        text
    )

    parser = etree.XMLParser(recover=True, huge_tree=True)
    root = etree.fromstring(text.encode("utf-8"), parser=parser)

    return root


def worksheet_table_to_dataframe(table, ns) -> pd.DataFrame:
    rows = []
    max_cols = 0

    for row in table.findall("ss:Row", namespaces=ns):
        row_values = []
        current_col = 1

        for cell in row.findall("ss:Cell", namespaces=ns):
            index_attr = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
            if index_attr:
                target_col = int(index_attr)
                while current_col < target_col:
                    row_values.append(None)
                    current_col += 1

            data = cell.find("ss:Data", namespaces=ns)
            value = data.text if data is not None else None
            row_values.append(value)
            current_col += 1

        max_cols = max(max_cols, len(row_values))
        rows.append(row_values)

    padded_rows = [r + [None] * (max_cols - len(r)) for r in rows]
    return pd.DataFrame(padded_rows)


def parse_excel_2003_xml_workbook(file_obj) -> dict[str, pd.DataFrame]:
    root = load_excel_2003_xml_root(file_obj)

    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    worksheets = root.findall(".//ss:Worksheet", namespaces=ns)

    if not worksheets:
        raise ValueError("Could not find any worksheets in XML file.")

    workbook = {}
    for idx, worksheet in enumerate(worksheets, start=1):
        sheet_name = worksheet.get("{urn:schemas-microsoft-com:office:spreadsheet}Name") or f"Sheet{idx}"
        table = worksheet.find("ss:Table", namespaces=ns)
        if table is None:
            workbook[sheet_name] = pd.DataFrame()
            continue
        workbook[sheet_name] = worksheet_table_to_dataframe(table, ns)

    return workbook


def parse_excel_2003_xml(file_obj) -> pd.DataFrame:
    workbook = parse_excel_2003_xml_workbook(file_obj)
    first_sheet = next(iter(workbook.values()), None)

    if first_sheet is None:
        raise ValueError("Could not find Worksheet/Table in XML file.")

    return first_sheet


# =========================================================
# LOAD TEMPLATE / ADDITIONAL DATA
# =========================================================

def load_template_workbook(template_source):
    xl = pd.ExcelFile(template_source)
    sheet_name = xl.sheet_names[0]
    template_raw = pd.read_excel(template_source, sheet_name=sheet_name, header=None)
    return template_raw, sheet_name


def load_additional_data(additional_source):
    sheets = pd.read_excel(additional_source, sheet_name=None)

    required = ["GENERAL", "ASSET_CLASS", "ACTUAL_CASH"]
    missing = [s for s in required if s not in sheets]
    if missing:
        raise ValueError(f"Missing sheet(s) in ADDITIONAL_DATA.xlsx: {', '.join(missing)}")

    general = sheets["GENERAL"].copy()
    asset_class = sheets["ASSET_CLASS"].copy()
    actual_cash = sheets["ACTUAL_CASH"].copy()

    general.columns = build_duplicate_headers(general.columns.tolist())
    asset_class.columns = build_duplicate_headers(asset_class.columns.tolist())
    actual_cash.columns = build_duplicate_headers(actual_cash.columns.tolist())

    return {
        "GENERAL": general,
        "ASSET_CLASS": asset_class,
        "ACTUAL_CASH": actual_cash,
    }


# =========================================================
# PORTFELJ PROCESSING
# =========================================================

def extract_fund_name_from_portfelj(df_raw: pd.DataFrame) -> str:
    if df_raw.empty:
        return ""

    col_a = df_raw.iloc[:, 0].astype(str)
    mask = col_a.str.contains("Klijent:", case=False, na=False)

    if not mask.any():
        return ""

    raw = df_raw.loc[mask].iloc[0, 0]
    text = safe_str(raw)

    if "Klijent:" in text:
        return text.split("Klijent:", 1)[-1].strip()
    return text.strip()


def process_portfelj(df_raw: pd.DataFrame, additional_data: dict):
    """
    Column mapping:
    A = 0
    B = 1
    C = 2
    D = 3
    E = 4
    I = 8
    K = 10
    N = 13
    O = 14
    Q = 16
    R = 17
    """
    fund_name = extract_fund_name_from_portfelj(df_raw)

    asset_class_df = additional_data["ASSET_CLASS"]
    actual_cash_df = additional_data["ACTUAL_CASH"]

    asset_class_map = {}
    for _, row in asset_class_df.iterrows():
        cro = safe_str(row.iloc[0])
        eng = safe_str(row.iloc[1]) if len(row) > 1 else ""
        if cro:
            asset_class_map[normalize_text(cro)] = eng

    actual_cash_keys_strict = set()
    actual_cash_keys_loose = set()

    for _, row in actual_cash_df.iterrows():
        key = safe_str(row.iloc[0])
        if key:
            actual_cash_keys_strict.add(normalize_text(key))
            actual_cash_keys_loose.add(normalize_text_loose(key))

    holdings = []

    for _, row in df_raw.iterrows():
        col_a = safe_str(row.iloc[0]) if len(row) > 0 else ""
        col_b = safe_str(row.iloc[1]) if len(row) > 1 else ""
        col_c = safe_str(row.iloc[2]) if len(row) > 2 else ""
        col_d = safe_str(row.iloc[3]) if len(row) > 3 else ""
        col_e = safe_str(row.iloc[4]) if len(row) > 4 else ""
        col_i = row.iloc[8] if len(row) > 8 else None
        col_k = row.iloc[10] if len(row) > 10 else None
        col_n = row.iloc[13] if len(row) > 13 else None
        col_o = row.iloc[14] if len(row) > 14 else None

        if col_e == "RDG" and col_d != "":
            fx_raw = safe_float(col_n, default=0.0)
            if fx_raw == 0:
                fx = 1.0
            else:
                fx = round(1.0 / fx_raw, 4)

            holding_class = asset_class_map.get(normalize_text(col_a), "")

            holdings.append({
                "[HOLDING_ISIN]": col_d,
                "[HOLDING_NAME]": col_c,
                "[HOLDING_QUANTITY]": safe_float(col_i, 0.0),
                "[HOLDING_CURRENCY]": col_b,
                "[HOLDING_PRICE]": safe_float(col_k, 0.0),
                "[HOLDING_ACC_INTEREST]": safe_float(col_o, 0.0),
                "[HOLDING_FX]": fx,
                "[HOLDING_CLASS]": holding_class
            })

    actual_cash = 0.0
    projected_cash = 0.0
    total_nav = 0.0

    for _, row in df_raw.iterrows():
        col_a = safe_str(row.iloc[0]) if len(row) > 0 else ""
        col_c = safe_str(row.iloc[2]) if len(row) > 2 else ""
        col_e = safe_str(row.iloc[4]) if len(row) > 4 else ""
        col_r = row.iloc[17] if len(row) > 17 else None
        r_val = safe_float(col_r, 0.0)

        if col_c != "":
            total_nav += r_val

        a_norm = normalize_text(col_a)
        a_loose = normalize_text_loose(col_a)

        is_actual_cash = (
            a_norm in actual_cash_keys_strict
            or a_loose in actual_cash_keys_loose
        )

        if is_actual_cash:
            actual_cash += r_val

        if col_c != "" and col_e == "" and not is_actual_cash:
            projected_cash += r_val

    return {
        "fund_name": fund_name,
        "holdings": holdings,
        "actual_cash": actual_cash,
        "projected_cash": projected_cash,
        "total_nav": total_nav
    }


# =========================================================
# PRINOS PROCESSING
# =========================================================

def prepare_prinos_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    return df_raw.copy()


def build_table_from_header_row(df_raw: pd.DataFrame, header_row_index: int = 1) -> pd.DataFrame:
    if df_raw.shape[0] <= header_row_index:
        return pd.DataFrame()

    headers = build_duplicate_headers(df_raw.iloc[header_row_index].tolist())
    data = df_raw.iloc[header_row_index + 1:].copy().reset_index(drop=True)
    data.columns = headers
    return data


def class_label_sort_key(label: str):
    match = re.match(r"([A-Za-z]+)(\d+)", safe_str(label))
    if match:
        return match.group(1).upper(), int(match.group(2))
    return safe_str(label).upper(), 0


def extract_prinos_class_data(prinos_workbook: dict[str, pd.DataFrame], portfolio_date: date) -> dict[str, list[dict]]:
    target_date = portfolio_date.strftime("%d.%m.%Y")
    class_data_by_fund = {}

    for sheet_name, sheet_df in list(prinos_workbook.items())[1:]:
        if sheet_df.empty:
            continue

        header_text = safe_str(sheet_df.iloc[0, 0]) if sheet_df.shape[0] > 0 and sheet_df.shape[1] > 0 else ""
        header_match = re.search(r"Fond:\s*(.*?)\s+Klasa udjela:\s*([A-Za-z]+\d+)", header_text)
        if not header_match:
            continue

        fund_name = header_match.group(1).strip()
        class_label = header_match.group(2).strip()

        table_df = build_table_from_header_row(sheet_df, header_row_index=1)
        if table_df.empty:
            continue

        date_column = find_column_label(table_df.columns, "Datum")
        units_column = find_column_label(table_df.columns, "Broj udjela")
        price_column = find_column_label(table_df.columns, "Cijena udjela")
        price_dv_column = find_column_label(table_df.columns, "Cijena udjela dv")

        if date_column is None or units_column is None:
            continue

        matching_rows = table_df[table_df[date_column].astype(str).str.strip() == target_date]
        if matching_rows.empty:
            continue

        row = matching_rows.iloc[0]
        class_data_by_fund.setdefault(fund_name, []).append({
            "sheet_name": sheet_name,
            "class_label": class_label,
            "[NUMBER_OF_UNITS]": safe_float(row[units_column], 0.0),
            "[NAV_PER_SHARE]": safe_float(
                row[price_column] if price_column is not None else row[price_dv_column] if price_dv_column is not None else None,
                0.0
            ),
        })

    for fund_name, entries in class_data_by_fund.items():
        entries.sort(key=lambda item: class_label_sort_key(item["class_label"]))

    return class_data_by_fund


def find_prinos_data(prinos_df: pd.DataFrame, fund_name: str, portfolio_date: date):
    target_date = portfolio_date.strftime("%d.%m.%Y")
    target_fund = normalize_text(fund_name)
    table_df = build_table_from_header_row(prinos_df, header_row_index=1)

    if table_df.empty:
        return None

    date_column = find_column_label(table_df.columns, "Datum")
    fund_column = find_column_label(table_df.columns, "Klijent")
    currency_column = find_column_label(table_df.columns, "Valuta")
    units_column = find_column_label(table_df.columns, "Broj udjela")
    price_column = find_column_label(table_df.columns, "Cijena udjela")
    price_dv_column = find_column_label(table_df.columns, "Cijena udjela dv")

    if date_column is None or fund_column is None:
        return None

    for _, row in table_df.iterrows():
        col_a = safe_str(row[date_column])
        col_b = safe_str(row[fund_column])

        if col_a == target_date and normalize_text(col_b) == target_fund:
            fund_currency = safe_str(row[currency_column]) if currency_column is not None else ""
            number_of_units = safe_float(row[units_column], 0.0) if units_column is not None else 0.0
            nav_per_share = safe_float(
                row[price_column] if price_column is not None else row[price_dv_column] if price_dv_column is not None else None,
                0.0
            )

            return {
                "[FUND_CURRENCY]": fund_currency,
                "[NUMBER_OF_UNITS]": number_of_units,
                "[NAV_PER_SHARE]": nav_per_share
            }

    return None


# =========================================================
# GENERAL LOOKUP
# =========================================================

def get_general_matches(additional_data: dict, fund_name: str) -> pd.DataFrame:
    general = additional_data["GENERAL"].copy()
    fund_name_column = find_column_label(general.columns, "FUND NAME")

    if fund_name_column is not None:
        fund_name_values = general[fund_name_column]
    elif general.shape[1] > 1:
        fund_name_values = general.iloc[:, 1]
    elif general.shape[1] > 0:
        fund_name_values = general.iloc[:, 0]
    else:
        return general.iloc[0:0].copy()

    matches = general[
        fund_name_values.astype(str).str.strip().str.casefold() == fund_name.strip().casefold()
    ].copy()

    return matches


# =========================================================
# TEMPLATE FILLING
# =========================================================

def replace_placeholders_in_row(row_values, replacements: dict):
    result = []
    for val in row_values:
        sval = safe_str(val)
        result.append(replacements.get(sval, val))
    return result


def build_output_from_template(template_raw: pd.DataFrame, fund_level_values: dict, holdings: list[dict]) -> pd.DataFrame:
    rows = template_raw.fillna("").values.tolist()

    holding_placeholder_names = {
        "[HOLDING_ISIN]",
        "[HOLDING_NAME]",
        "[HOLDING_QUANTITY]",
        "[HOLDING_CURRENCY]",
        "[HOLDING_PRICE]",
        "[HOLDING_ACC_INTEREST]",
        "[HOLDING_FX]",
        "[HOLDING_CLASS]",
    }

    holding_row_idx = None
    for i, row in enumerate(rows):
        row_set = set(safe_str(v) for v in row)
        if row_set & holding_placeholder_names:
            holding_row_idx = i
            break

    if holding_row_idx is None:
        raise ValueError("Could not find holdings template row in template.xlsx")

    output_rows = []

    for i, row in enumerate(rows):
        if i == holding_row_idx:
            if holdings:
                for h in holdings:
                    repl = fund_level_values.copy()
                    repl.update(h)
                    output_rows.append(replace_placeholders_in_row(row, repl))
            else:
                empty_holding = {
                    "[HOLDING_ISIN]": "",
                    "[HOLDING_NAME]": "",
                    "[HOLDING_QUANTITY]": "",
                    "[HOLDING_CURRENCY]": "",
                    "[HOLDING_PRICE]": "",
                    "[HOLDING_ACC_INTEREST]": "",
                    "[HOLDING_FX]": "",
                    "[HOLDING_CLASS]": "",
                }
                repl = fund_level_values.copy()
                repl.update(empty_holding)
                output_rows.append(replace_placeholders_in_row(row, repl))
        else:
            output_rows.append(replace_placeholders_in_row(row, fund_level_values))

    return pd.DataFrame(output_rows)


# =========================================================
# CSV EXPORT
# =========================================================

def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    csv_text = df.to_csv(index=False, header=False)
    return csv_text.encode("utf-8-sig")


def build_zip(file_map: dict[str, bytes]) -> bytes:
    mem = BytesIO()
    with zipfile.ZipFile(mem, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, content in file_map.items():
            zf.writestr(filename, content)
    mem.seek(0)
    return mem.getvalue()


# =========================================================
# MAIN GENERATION LOGIC
# =========================================================

def generate_outputs(
    template_source,
    additional_source,
    prinos_source,
    portfelj_files,
    portfolio_date: date,
    trade_date: date
):
    template_raw, _ = load_template_workbook(template_source)
    additional_data = load_additional_data(additional_source)

    if hasattr(prinos_source, "seek"):
        prinos_source.seek(0)

    prinos_workbook = parse_excel_2003_xml_workbook(prinos_source)
    prinos_df_raw = next(iter(prinos_workbook.values()))
    prinos_df = prepare_prinos_dataframe(prinos_df_raw)
    prinos_class_data = extract_prinos_class_data(prinos_workbook, portfolio_date)

    output_files = {}
    log_rows = []

    portfolio_date_str = portfolio_date.strftime("%Y%m%d")

    for pf in portfelj_files:
        if hasattr(pf, "seek"):
            pf.seek(0)

        portfelj_df_raw = parse_excel_2003_xml(pf)
        portfelj_info = process_portfelj(portfelj_df_raw, additional_data)

        fund_name = portfelj_info["fund_name"]
        if not fund_name:
            log_rows.append({
                "PORTFELJ_FILE": pf.name,
                "FUND_NAME": "",
                "STATUS": "Skipped",
                "MESSAGE": "Could not find 'Klijent:' row in PORTFELJ file."
            })
            continue

        prinos_info = find_prinos_data(prinos_df, fund_name, portfolio_date)
        if prinos_info is None:
            log_rows.append({
                "PORTFELJ_FILE": pf.name,
                "FUND_NAME": fund_name,
                "STATUS": "Skipped",
                "MESSAGE": f"No matching PRINOS data for date {portfolio_date.strftime('%d.%m.%Y')} and fund '{fund_name}'."
            })
            continue

        general_matches = get_general_matches(additional_data, fund_name)
        if general_matches.empty:
            log_rows.append({
                "PORTFELJ_FILE": pf.name,
                "FUND_NAME": fund_name,
                "STATUS": "Skipped",
                "MESSAGE": "No matching rows found in ADDITIONAL_DATA / GENERAL."
            })
            continue

        class_entries = prinos_class_data.get(fund_name, [])

        for class_index, (_, g_row) in enumerate(general_matches.iterrows()):
            fund_isin = safe_str(get_value_by_column_or_index(g_row, 0, "FUND ISIN"))
            fund_name_general = safe_str(get_value_by_column_or_index(g_row, 1, "FUND NAME")) or fund_name
            fund_ticker = safe_str(get_value_by_column_or_index(g_row, 3, "BLOOMBERG TICKER"))
            output_name = safe_str(get_value_by_column_or_index(g_row, 4, "OUTPUT FILE NAME"))

            if fund_isin == "":
                log_rows.append({
                    "PORTFELJ_FILE": pf.name,
                    "FUND_NAME": fund_name,
                    "STATUS": "Skipped",
                    "MESSAGE": "Matching GENERAL row exists, but FUND ISIN is empty."
                })
                continue

            if output_name == "":
                log_rows.append({
                    "PORTFELJ_FILE": pf.name,
                    "FUND_NAME": fund_name,
                    "STATUS": "Skipped",
                    "MESSAGE": "Matching GENERAL row exists, but GENERAL column E (output filename) is empty."
                })
                continue

            fund_level_values = {
                "[FUND_NAME]": fund_name_general,
                "[PORTFOLIO_DATE]": portfolio_date_str,
                "[TRADE_DATE]": trade_date.strftime("%Y%m%d"),
                "[FUND_ISIN]": fund_isin,
                "[FUND_TICKER]": fund_ticker,
                "[ACTUAL_CASH]": portfelj_info["actual_cash"],
                "[PROJECTED_CASH]": portfelj_info["projected_cash"],
                "[TOTAL_NAV]": portfelj_info["total_nav"],
            }
            fund_level_values.update(prinos_info)

            if class_index < len(class_entries):
                fund_level_values.update({
                    "[NUMBER_OF_UNITS]": class_entries[class_index]["[NUMBER_OF_UNITS]"],
                    "[NAV_PER_SHARE]": class_entries[class_index]["[NAV_PER_SHARE]"],
                })

            output_df = build_output_from_template(
                template_raw=template_raw,
                fund_level_values=fund_level_values,
                holdings=portfelj_info["holdings"]
            )

            filename = f"{sanitize_filename(output_name)}_{portfolio_date_str}.csv"
            output_files[filename] = dataframe_to_csv_bytes(output_df)

            log_rows.append({
                "PORTFELJ_FILE": pf.name,
                "FUND_NAME": fund_name,
                "STATUS": "Created",
                "MESSAGE": f"Created {filename}"
            })

    log_df = pd.DataFrame(log_rows)
    return output_files, log_df


def run_streamlit_app():
    import streamlit as st

    st.set_page_config(page_title="Tradeweb CSV Generator", layout="wide")

    st.title("Tradeweb CSV Generator")

    st.markdown(
        """
Upload:
- **1 PRINOS XML**
- **1 or more PORTFELJ XML files**
- optional **ADDITIONAL_DATA.xlsx**
- optional **template.xlsx**

If ADDITIONAL_DATA or template are not uploaded, the app will use files from the same folder as `Tradeweb.py`.
"""
    )

    col1, col2 = st.columns(2)

    with col1:
        portfolio_date = st.date_input("Portfolio date", value=date.today())

    with col2:
        default_trade_date = next_weekday(portfolio_date)
        trade_date = st.date_input("Trade date", value=default_trade_date)

    st.divider()

    uploaded_prinos = st.file_uploader(
        "Upload PRINOS XML",
        type=["xml"],
        accept_multiple_files=False
    )

    uploaded_portfelj = st.file_uploader(
        "Upload PORTFELJ XML files",
        type=["xml"],
        accept_multiple_files=True
    )

    uploaded_additional = st.file_uploader(
        "Upload ADDITIONAL_DATA.xlsx (optional)",
        type=["xlsx"],
        accept_multiple_files=False
    )

    uploaded_template = st.file_uploader(
        "Upload template.xlsx (optional)",
        type=["xlsx"],
        accept_multiple_files=False
    )

    st.divider()

    if st.button("Generate CSV files", type="primary"):
        try:
            if uploaded_prinos is None:
                st.error("Please upload the PRINOS XML file.")
                st.stop()

            if not uploaded_portfelj:
                st.error("Please upload at least one PORTFELJ XML file.")
                st.stop()

            if uploaded_template is not None:
                template_source = uploaded_template
            else:
                if not DEFAULT_TEMPLATE_PATH.exists():
                    st.error(f"Template file not found: {DEFAULT_TEMPLATE_PATH}")
                    st.stop()
                template_source = DEFAULT_TEMPLATE_PATH

            if uploaded_additional is not None:
                additional_source = uploaded_additional
            else:
                if not DEFAULT_ADDITIONAL_DATA_PATH.exists():
                    st.error(f"Default ADDITIONAL_DATA file not found: {DEFAULT_ADDITIONAL_DATA_PATH}")
                    st.stop()
                additional_source = DEFAULT_ADDITIONAL_DATA_PATH

            with st.spinner("Generating CSV files..."):
                outputs, log_df = generate_outputs(
                    template_source=template_source,
                    additional_source=additional_source,
                    prinos_source=uploaded_prinos,
                    portfelj_files=uploaded_portfelj,
                    portfolio_date=portfolio_date,
                    trade_date=trade_date
                )

            st.subheader("Processing log")
            if not log_df.empty:
                st.dataframe(log_df, use_container_width=True)
            else:
                st.info("No log entries were produced.")

            if outputs:
                st.success(f"Created {len(outputs)} CSV file(s).")

                zip_bytes = build_zip(outputs)
                st.download_button(
                    label="Download all CSV files as ZIP",
                    data=zip_bytes,
                    file_name=f"tradeweb_csvs_{trade_date.strftime('%Y%m%d')}.zip",
                    mime="application/zip"
                )

                st.subheader("Individual files")
                for filename, content in outputs.items():
                    st.download_button(
                        label=f"Download {filename}",
                        data=content,
                        file_name=filename,
                        mime="text/csv",
                        key=filename
                    )
            else:
                st.warning("No CSV files were created. Check the processing log above.")

        except Exception as e:
            st.exception(e)


if __name__ == "__main__":
    run_streamlit_app()
