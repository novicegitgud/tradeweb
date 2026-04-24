import streamlit as st
import pandas as pd
from datetime import date, timedelta
from pathlib import Path
from io import BytesIO
import zipfile
import re
import unicodedata

# =========================================================
# CONFIG
# =========================================================

st.set_page_config(page_title="Tradeweb CSV Generator", layout="wide")

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


def derive_trade_date(portfolio_date: date) -> date:
    return next_weekday(portfolio_date)


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


PORTFELJ_COLUMN_ALIASES = {
    "position_type": ("Vrsta pozicije", "Position type", "Vrsta"),
    "currency": ("Valuta", "Currency"),
    "position_name": ("Pozicija", "Position", "Holding name", "Naziv pozicije"),
    "isin": ("ISIN",),
    "instrument_type": ("Tip", "Type", "Instrument type"),
    "quantity": ("Količina", "Kolicina", "Quantity"),
    "price": ("Cijena", "Price"),
    "fx_rate": ("Tečaj DV", "Tecaj DV", "FX", "FX rate", "Exchange rate"),
    "accrued_interest": ("Kamata", "Accrued interest", "Interest"),
    "amount_base": ("Iznos DV", "Amount DV", "Base amount", "Amount in base currency"),
}

PRINOS_COLUMN_ALIASES = {
    "date": ("Datum", "Date", "Portfolio date", "NAV date"),
    "fund_name": ("Fund name", "Naziv fonda", "Client", "Klijent"),
    "fund_currency": ("Fund currency", "Currency", "Valuta", "Valuta fonda"),
    "number_of_units": ("NUMBER_OF_UNITS", "Number of units", "Broj udjela", "Broj jedinica", "Units"),
    "nav_per_share": ("NAV_PER_SHARE", "NAV per share", "NAV/share", "Share NAV", "Cijena udjela"),
}

GENERAL_COLUMN_ALIASES = {
    "fund_isin": ("FUND ISIN", "Fund ISIN", "ISIN"),
    "fund_name": ("FUND NAME", "Fund name", "Naziv fonda"),
    "fund_ticker": ("BLOOMBERG TICKER", "Bloomberg ticker", "Ticker", "Bloomberg"),
    "output_name": ("OUTPUT FILE NAME", "Output file name", "File name", "Naziv izlazne datoteke"),
}

ASSET_CLASS_COLUMN_ALIASES = {
    "croatian": ("HRVATSKI", "Croatian", "Asset class HR"),
    "english": ("ENGLISH", "English", "Asset class EN"),
}

ACTUAL_CASH_COLUMN_ALIASES = {
    "cash_item": ("SMATRA SE ACTUAL CASH-OM", "Actual cash", "Cash item", "Vrsta pozicije"),
}


def normalize_header_name(value):
    text = normalize_text_loose(value)
    return re.sub(r"[^a-z0-9]+", "", text)


def column_name_matches(normalized_label: str, normalized_alias: str) -> bool:
    if not normalized_label or not normalized_alias:
        return False

    if normalized_label == normalized_alias:
        return True

    if len(normalized_alias) >= 6 and normalized_alias in normalized_label:
        return True

    if len(normalized_label) >= 6 and normalized_label in normalized_alias:
        return True

    return False


def resolve_columns_from_labels(labels, alias_map):
    candidates = []
    for label in labels:
        label_text = safe_str(label)
        if label_text == "" or label_text.startswith("Unnamed"):
            continue
        candidates.append((label, normalize_header_name(label_text)))

    resolved = {}

    for field, aliases in alias_map.items():
        normalized_aliases = [normalize_header_name(alias) for alias in aliases]
        normalized_aliases = [alias for alias in normalized_aliases if alias]

        exact_match = next(
            (
                label
                for label, normalized_label in candidates
                if normalized_label in normalized_aliases
            ),
            None,
        )
        if exact_match is not None:
            resolved[field] = exact_match
            continue

        fuzzy_match = next(
            (
                label
                for label, normalized_label in candidates
                if any(
                    column_name_matches(normalized_label, normalized_alias)
                    for normalized_alias in normalized_aliases
                )
            ),
            None,
        )
        if fuzzy_match is not None:
            resolved[field] = fuzzy_match

    return resolved


def resolve_dataframe_columns(df: pd.DataFrame, alias_map, required_fields, context: str, fallback_indices=None):
    resolved = resolve_columns_from_labels(df.columns.tolist(), alias_map)

    if fallback_indices:
        for field, idx in fallback_indices.items():
            if field in resolved:
                continue
            if 0 <= idx < len(df.columns):
                resolved[field] = df.columns[idx]

    missing = [field for field in required_fields if field not in resolved]
    if missing:
        available_headers = ", ".join(safe_str(col) for col in df.columns if safe_str(col))
        raise ValueError(
            f"Could not resolve required column(s) {', '.join(missing)} in {context}. "
            f"Available headers: {available_headers}"
        )

    return resolved


def build_named_table_from_raw(
    df_raw: pd.DataFrame,
    alias_map,
    required_fields,
    context: str,
    min_matches: int,
    fallback_indices=None,
):
    if df_raw.empty:
        raise ValueError(f"{context} is empty.")

    header_row_idx = None
    best_match_count = -1

    for idx in range(len(df_raw)):
        header_candidates = build_duplicate_headers(df_raw.iloc[idx].tolist())
        resolved = resolve_columns_from_labels(header_candidates, alias_map)
        match_count = len(resolved)

        if match_count > best_match_count:
            best_match_count = match_count

        if match_count >= min_matches:
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise ValueError(
            f"Could not find a header row in {context}. "
            f"Best matched {best_match_count} expected header(s)."
        )

    headers = build_duplicate_headers(df_raw.iloc[header_row_idx].tolist())
    data_df = df_raw.iloc[header_row_idx + 1 :].copy().reset_index(drop=True)
    data_df.columns = headers

    resolved_columns = resolve_dataframe_columns(
        data_df,
        alias_map=alias_map,
        required_fields=required_fields,
        context=context,
        fallback_indices=fallback_indices,
    )

    return data_df, resolved_columns


def get_row_value(row: pd.Series, columns: dict, key: str):
    column_name = columns.get(key)
    if not column_name:
        return None
    return row.get(column_name)


# =========================================================
# ROBUST XML PARSER FOR EXCEL 2003 XML
# =========================================================

def parse_excel_2003_xml(file_obj) -> pd.DataFrame:
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

    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    table = root.find(".//ss:Worksheet/ss:Table", namespaces=ns)

    if table is None:
        raise ValueError("Could not find Worksheet/Table in XML file.")

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

    general_columns = resolve_dataframe_columns(
        general,
        alias_map=GENERAL_COLUMN_ALIASES,
        required_fields=["fund_isin", "fund_name", "output_name"],
        context="ADDITIONAL_DATA / GENERAL",
        fallback_indices={
            "fund_isin": 0,
            "fund_name": 1,
            "fund_ticker": 3,
            "output_name": 4,
        },
    )
    asset_class_columns = resolve_dataframe_columns(
        asset_class,
        alias_map=ASSET_CLASS_COLUMN_ALIASES,
        required_fields=["croatian", "english"],
        context="ADDITIONAL_DATA / ASSET_CLASS",
        fallback_indices={"croatian": 0, "english": 1},
    )
    actual_cash_columns = resolve_dataframe_columns(
        actual_cash,
        alias_map=ACTUAL_CASH_COLUMN_ALIASES,
        required_fields=["cash_item"],
        context="ADDITIONAL_DATA / ACTUAL_CASH",
        fallback_indices={"cash_item": 0},
    )

    return {
        "GENERAL": general,
        "GENERAL_COLUMNS": general_columns,
        "ASSET_CLASS": asset_class,
        "ASSET_CLASS_COLUMNS": asset_class_columns,
        "ACTUAL_CASH": actual_cash,
        "ACTUAL_CASH_COLUMNS": actual_cash_columns,
    }


# =========================================================
# PORTFELJ PROCESSING
# =========================================================

def extract_fund_name_from_portfelj(df_raw: pd.DataFrame) -> str:
    if df_raw.empty:
        return ""

    for _, row in df_raw.iterrows():
        row_values = [safe_str(value) for value in row.tolist() if safe_str(value)]
        if not row_values:
            continue

        joined = " ".join(row_values)
        normalized_joined = normalize_text_loose(joined)
        if "klijent" not in normalized_joined:
            continue

        match = re.search(r"klijent\s*:\s*(.+)", joined, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()

        for idx, value in enumerate(row_values):
            normalized_value = normalize_header_name(value)
            if normalized_value == "klijent":
                if idx + 1 < len(row_values):
                    return row_values[idx + 1]
                return ""

            if normalize_text_loose(value).startswith("klijent"):
                return re.sub(r"(?i)^klijent\s*:?\s*", "", value).strip()

        return joined.strip()

    return ""


def prepare_portfelj_dataframe(df_raw: pd.DataFrame):
    return build_named_table_from_raw(
        df_raw,
        alias_map=PORTFELJ_COLUMN_ALIASES,
        required_fields=[
            "position_type",
            "currency",
            "position_name",
            "isin",
            "instrument_type",
            "quantity",
            "price",
            "fx_rate",
            "accrued_interest",
            "amount_base",
        ],
        context="PORTFELJ XML",
        min_matches=6,
    )


def process_portfelj(df_raw: pd.DataFrame, additional_data: dict):
    fund_name = extract_fund_name_from_portfelj(df_raw)
    portfelj_df, portfelj_columns = prepare_portfelj_dataframe(df_raw)

    asset_class_df = additional_data["ASSET_CLASS"]
    asset_class_columns = additional_data["ASSET_CLASS_COLUMNS"]
    actual_cash_df = additional_data["ACTUAL_CASH"]
    actual_cash_columns = additional_data["ACTUAL_CASH_COLUMNS"]

    asset_class_map = {}
    for _, row in asset_class_df.iterrows():
        cro = safe_str(get_row_value(row, asset_class_columns, "croatian"))
        eng = safe_str(get_row_value(row, asset_class_columns, "english"))
        if cro:
            asset_class_map[normalize_text(cro)] = eng

    actual_cash_keys_strict = set()
    actual_cash_keys_loose = set()

    for _, row in actual_cash_df.iterrows():
        key = safe_str(get_row_value(row, actual_cash_columns, "cash_item"))
        if key:
            actual_cash_keys_strict.add(normalize_text(key))
            actual_cash_keys_loose.add(normalize_text_loose(key))

    holdings = []

    for _, row in portfelj_df.iterrows():
        position_type = safe_str(get_row_value(row, portfelj_columns, "position_type"))
        currency = safe_str(get_row_value(row, portfelj_columns, "currency"))
        position_name = safe_str(get_row_value(row, portfelj_columns, "position_name"))
        isin = safe_str(get_row_value(row, portfelj_columns, "isin"))
        instrument_type = safe_str(get_row_value(row, portfelj_columns, "instrument_type"))
        quantity = get_row_value(row, portfelj_columns, "quantity")
        price = get_row_value(row, portfelj_columns, "price")
        fx_source = get_row_value(row, portfelj_columns, "fx_rate")
        accrued_interest = get_row_value(row, portfelj_columns, "accrued_interest")

        if normalize_text(instrument_type) == "rdg" and isin != "":
            fx_raw = safe_float(fx_source, default=0.0)
            if fx_raw == 0:
                fx = 1.0
            else:
                fx = round(1.0 / fx_raw, 4)

            holding_class = asset_class_map.get(normalize_text(position_type), "")

            holdings.append({
                "[HOLDING_ISIN]": isin,
                "[HOLDING_NAME]": position_name,
                "[HOLDING_QUANTITY]": safe_float(quantity, 0.0),
                "[HOLDING_CURRENCY]": currency,
                "[HOLDING_PRICE]": safe_float(price, 0.0),
                "[HOLDING_ACC_INTEREST]": safe_float(accrued_interest, 0.0),
                "[HOLDING_FX]": fx,
                "[HOLDING_CLASS]": holding_class
            })

    actual_cash = 0.0
    projected_cash = 0.0
    total_nav = 0.0

    for _, row in portfelj_df.iterrows():
        position_type = safe_str(get_row_value(row, portfelj_columns, "position_type"))
        position_name = safe_str(get_row_value(row, portfelj_columns, "position_name"))
        instrument_type = safe_str(get_row_value(row, portfelj_columns, "instrument_type"))
        amount_base = get_row_value(row, portfelj_columns, "amount_base")
        amount_value = safe_float(amount_base, 0.0)

        if position_name != "":
            total_nav += amount_value

        a_norm = normalize_text(position_type)
        a_loose = normalize_text_loose(position_type)

        is_actual_cash = (
            a_norm in actual_cash_keys_strict
            or a_loose in actual_cash_keys_loose
        )

        if is_actual_cash:
            actual_cash += amount_value

        if position_name != "" and instrument_type == "" and not is_actual_cash:
            projected_cash += amount_value

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

def prepare_prinos_dataframe(df_raw: pd.DataFrame):
    try:
        prinos_df, prinos_columns = build_named_table_from_raw(
            df_raw,
            alias_map=PRINOS_COLUMN_ALIASES,
            required_fields=[
                "date",
                "fund_name",
                "fund_currency",
                "number_of_units",
                "nav_per_share",
            ],
            context="PRINOS XML",
            min_matches=3,
        )
        return {
            "data": prinos_df,
            "columns": prinos_columns,
        }
    except ValueError:
        prinos_df = df_raw.copy().reset_index(drop=True)
        prinos_df.columns = [f"Column_{idx}" for idx in range(len(prinos_df.columns))]
        prinos_columns = resolve_dataframe_columns(
            prinos_df,
            alias_map={},
            required_fields=[
                "date",
                "fund_name",
                "fund_currency",
                "number_of_units",
                "nav_per_share",
            ],
            context="PRINOS XML (legacy positional fallback)",
            fallback_indices={
                "date": 0,
                "fund_name": 1,
                "fund_currency": 4,
                "number_of_units": 24,
                "nav_per_share": 28,
            },
        )
        return {
            "data": prinos_df,
            "columns": prinos_columns,
        }


def find_prinos_data(prinos_data, fund_name: str, portfolio_date: date):
    prinos_df = prinos_data["data"]
    prinos_columns = prinos_data["columns"]
    target_date = portfolio_date.strftime("%d.%m.%Y")
    target_fund = normalize_text(fund_name)

    for _, row in prinos_df.iterrows():
        row_date = safe_str(get_row_value(row, prinos_columns, "date"))
        row_fund_name = safe_str(get_row_value(row, prinos_columns, "fund_name"))

        if row_date == target_date and normalize_text(row_fund_name) == target_fund:
            fund_currency = safe_str(get_row_value(row, prinos_columns, "fund_currency"))
            number_of_units = safe_float(get_row_value(row, prinos_columns, "number_of_units"), 0.0)
            nav_per_share = safe_float(get_row_value(row, prinos_columns, "nav_per_share"), 0.0)

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
    general_columns = additional_data["GENERAL_COLUMNS"]
    fund_name_column = general_columns["fund_name"]

    matches = general[
        general[fund_name_column].astype(str).str.strip().str.casefold() == fund_name.strip().casefold()
    ].copy()

    return matches


def get_expected_output_entries(
    additional_data: dict,
    portfolio_date: date,
    fund_names: set[str] | None = None,
) -> list[dict]:
    general = additional_data["GENERAL"]
    general_columns = additional_data["GENERAL_COLUMNS"]
    portfolio_date_str = portfolio_date.strftime("%Y%m%d")

    normalized_fund_names = None
    if fund_names is not None:
        normalized_fund_names = {
            normalize_text(fund_name)
            for fund_name in fund_names
            if safe_str(fund_name)
        }

    expected_entries = []

    for _, row in general.iterrows():
        fund_name = safe_str(get_row_value(row, general_columns, "fund_name"))
        fund_isin = safe_str(get_row_value(row, general_columns, "fund_isin"))
        fund_ticker = safe_str(get_row_value(row, general_columns, "fund_ticker"))
        output_name = safe_str(get_row_value(row, general_columns, "output_name"))

        if not fund_name or not fund_isin or not output_name:
            continue

        if normalized_fund_names is not None and normalize_text(fund_name) not in normalized_fund_names:
            continue

        expected_entries.append({
            "fund_name": fund_name,
            "fund_isin": fund_isin,
            "fund_ticker": fund_ticker,
            "output_name": output_name,
            "filename": f"{sanitize_filename(output_name)}_{portfolio_date_str}.csv",
        })

    return expected_entries


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
    trade_date: date | None = None
):
    template_raw, _ = load_template_workbook(template_source)
    additional_data = load_additional_data(additional_source)
    trade_date = derive_trade_date(portfolio_date)

    if hasattr(prinos_source, "seek"):
        prinos_source.seek(0)

    prinos_df_raw = parse_excel_2003_xml(prinos_source)
    prinos_data = prepare_prinos_dataframe(prinos_df_raw)

    output_files = {}
    log_rows = []
    processed_fund_names = set()

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

        processed_fund_names.add(fund_name)

        prinos_info = find_prinos_data(prinos_data, fund_name, portfolio_date)
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

        general_columns = additional_data["GENERAL_COLUMNS"]
        for _, g_row in general_matches.iterrows():
            fund_isin = safe_str(get_row_value(g_row, general_columns, "fund_isin"))
            fund_name_general = safe_str(get_row_value(g_row, general_columns, "fund_name")) or fund_name
            fund_ticker = safe_str(get_row_value(g_row, general_columns, "fund_ticker"))
            output_name = safe_str(get_row_value(g_row, general_columns, "output_name"))

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
                    "MESSAGE": "Matching GENERAL row exists, but the output filename is empty."
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

    expected_entries = get_expected_output_entries(
        additional_data=additional_data,
        portfolio_date=portfolio_date,
        fund_names=processed_fund_names,
    )
    created_filenames = set(output_files.keys())
    missing_entries = [
        entry for entry in expected_entries
        if entry["filename"] not in created_filenames
    ]

    for entry in missing_entries:
        missing_parts = [entry["fund_isin"]]
        if entry["fund_ticker"]:
            missing_parts.append(entry["fund_ticker"])
        identifier = " / ".join(missing_parts)

        log_rows.append({
            "PORTFELJ_FILE": "",
            "FUND_NAME": entry["fund_name"],
            "STATUS": "Missing",
            "MESSAGE": f"Expected share-class PCF not created: {identifier}"
        })

    log_df = pd.DataFrame(log_rows)
    summary = {
        "trade_date": trade_date,
        "expected_output_count": len(expected_entries),
        "created_output_count": len(output_files),
        "missing_entries": missing_entries,
    }
    return output_files, log_df, summary


# =========================================================
# UI
# =========================================================

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
    trade_date = derive_trade_date(portfolio_date)
    st.text_input("Trade date", value=trade_date.strftime("%d.%m.%Y"), disabled=True)

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
            outputs, log_df, summary = generate_outputs(
                template_source=template_source,
                additional_source=additional_source,
                prinos_source=uploaded_prinos,
                portfelj_files=uploaded_portfelj,
                portfolio_date=portfolio_date
            )

        st.subheader("Processing log")
        if not log_df.empty:
            st.dataframe(log_df, use_container_width=True)
        else:
            st.info("No log entries were produced.")

        st.caption(
            f"Trade date used for all generated PCFs: {summary['trade_date'].strftime('%d.%m.%Y')}"
        )

        if outputs:
            if summary["missing_entries"]:
                st.warning(
                    "Created "
                    f"{summary['created_output_count']} of {summary['expected_output_count']} "
                    "expected share-class PCF(s) for the uploaded funds."
                )
            else:
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
