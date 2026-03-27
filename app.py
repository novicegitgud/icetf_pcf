import io
import math
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet


# ----------------------------
# Basic page setup
# ----------------------------
st.set_page_config(page_title="ETF PCF Prototype", layout="wide")
st.title("ETF PCF Prototype")
st.caption("Proof of concept for creation/redemption basket configuration")

EXCEL_FILE = Path("ICETF_PCF.xlsx")


# ----------------------------
# Data loading
# ----------------------------
@st.cache_data
def load_data(file_path: Path) -> dict:
    data = {
        "pcf_creation": pd.read_excel(file_path, sheet_name="PCF_Creation"),
        "pcf_redemption": pd.read_excel(file_path, sheet_name="PCF_Redemption"),
        "cash_creation": pd.read_excel(file_path, sheet_name="Cash_Creation"),
        "cash_redemption": pd.read_excel(file_path, sheet_name="Cash_Redemption"),
        "fx": pd.read_excel(file_path, sheet_name="Currency_Exchange"),
        "allowed": pd.read_excel(file_path, sheet_name="Allowed_Currencies"),
    }

    for _, df in data.items():
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"]).dt.date

    return data


if not EXCEL_FILE.exists():
    st.error("Excel file not found. Make sure ICETF_PCF.xlsx is in the same folder as app.py.")
    st.stop()

data = load_data(EXCEL_FILE)


# ----------------------------
# Helper functions
# ----------------------------
def get_pcf_df(transaction_label: str) -> pd.DataFrame:
    return data["pcf_creation"].copy() if transaction_label == "Creation" else data["pcf_redemption"].copy()


def get_cash_df(transaction_label: str) -> pd.DataFrame:
    return data["cash_creation"].copy() if transaction_label == "Creation" else data["cash_redemption"].copy()


def get_available_dates_for_etf(etf: str) -> list:
    creation_df = data["pcf_creation"]
    return sorted(creation_df.loc[creation_df["ETF"] == etf, "Date"].dropna().unique().tolist())


def get_allowed_cash_currencies(etf: str, basket_date) -> list:
    allowed = data["allowed"]
    filtered = allowed[(allowed["ETF"] == etf) & (allowed["Date"] == basket_date)]
    return sorted(filtered["Currency"].dropna().unique().tolist())


def get_fx_rate(from_ccy: str, to_ccy: str, basket_date) -> float:
    fx = data["fx"]
    row = fx[
        (fx["From"] == from_ccy) &
        (fx["To"] == to_ccy) &
        (fx["Date"] == basket_date)
    ]
    if row.empty:
        raise ValueError(f"No FX rate found for {from_ccy} -> {to_ccy} on {basket_date}")
    return float(row.iloc[0]["Rate"])


def get_base_cash_10000(etf: str, basket_date, transaction_label: str, output_currency: str) -> float:
    cash_df = get_cash_df(transaction_label)
    row = cash_df[(cash_df["ETF"] == etf) & (cash_df["Date"] == basket_date)]

    if row.empty:
        raise ValueError(f"No cash row found for {etf} / {basket_date} / {transaction_label}")

    amount = float(row.iloc[0]["Cash_10000"])
    base_currency = row.iloc[0]["Cash_Currency"]
    fx_rate = get_fx_rate(base_currency, output_currency, basket_date)

    return amount * fx_rate


def integer_bounds_from_ideal(ideal_qty: float) -> tuple[int, int]:
    """
    Actual stock quantities must be whole numbers.

    Rules:
    - If ideal is positive, allowed actual is from 0 to floor(ideal)
    - If ideal is negative, allowed actual is from ceil(ideal) to 0
    """
    if ideal_qty >= 0:
        return 0, math.floor(ideal_qty)
    return math.ceil(ideal_qty), 0


def default_actual_from_ideal(ideal_qty: float) -> int:
    """
    Default AP input as the nearest allowed whole-stock quantity:
    - positive ideal -> floor
    - negative ideal -> ceil
    """
    if ideal_qty >= 0:
        return math.floor(ideal_qty)
    return math.ceil(ideal_qty)


def build_scaled_basket(etf: str, basket_date, transaction_label: str, requested_units: float, cash_currency: str) -> pd.DataFrame:
    pcf = get_pcf_df(transaction_label)
    basket = pcf[(pcf["ETF"] == etf) & (pcf["Date"] == basket_date)].copy()

    if basket.empty:
        raise ValueError(f"No PCF rows found for {etf} / {basket_date} / {transaction_label}")

    scale_factor = requested_units / 10000.0
    basket["Ideal_Quantity"] = basket["Quantity_10000"] * scale_factor

    min_list = []
    max_list = []
    default_list = []

    for ideal in basket["Ideal_Quantity"]:
        min_val, max_val = integer_bounds_from_ideal(ideal)
        min_list.append(min_val)
        max_list.append(max_val)
        default_list.append(default_actual_from_ideal(ideal))

    basket["Min_Actual_Integer"] = min_list
    basket["Max_Actual_Integer"] = max_list
    basket["Default_AP_Quantity"] = default_list
    basket["AP_Input_Quantity"] = basket["Default_AP_Quantity"]

    basket["FX_to_Cash"] = basket["Currency"].apply(lambda ccy: get_fx_rate(ccy, cash_currency, basket_date))
    basket["Price_in_Cash_Currency"] = basket["Price"] * basket["FX_to_Cash"]

    return basket


def validate_actual_vs_bounds(actual_qty: int, min_qty: int, max_qty: int) -> bool:
    return min_qty <= actual_qty <= max_qty


def calculate_cash_result(edited_basket: pd.DataFrame, base_cash_scaled: float) -> tuple[float, pd.DataFrame]:
    result = edited_basket.copy()
    result["Difference_Qty"] = result["Ideal_Quantity"] - result["AP_Input_Quantity"]
    result["Cash_Adjustment"] = result["Difference_Qty"] * result["Price_in_Cash_Currency"]
    result["Changed_By_AP"] = result["AP_Input_Quantity"] != result["Default_AP_Quantity"]
    result["Delta_vs_Default"] = result["AP_Input_Quantity"] - result["Default_AP_Quantity"]
    final_cash = base_cash_scaled + result["Cash_Adjustment"].sum()
    return final_cash, result


def create_pdf(
    ap_name: str,
    request_reference: str,
    request_date,
    settlement_date,
    comments: str,
    etf: str,
    transaction_label: str,
    basket_date,
    units: float,
    cash_currency: str,
    base_cash_scaled: float,
    final_cash: float,
    result_df: pd.DataFrame,
) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("ETF Creation / Redemption Request", styles["Title"]))
    story.append(Spacer(1, 6))

    details = [
        ["AP Name", ap_name or ""],
        ["Reference", request_reference or ""],
        ["Request Date", str(request_date)],
        ["Settlement Date", str(settlement_date)],
        ["ETF", etf],
        ["Transaction Type", transaction_label],
        ["Basket Date", str(basket_date)],
        ["ETF Units", f"{units:,.0f}"],
        ["Cash Currency", cash_currency],
        ["Base Cash Component", f"{base_cash_scaled:,.2f} {cash_currency}"],
        ["Final Cash Component", f"{final_cash:,.2f} {cash_currency}"],
    ]

    details_table = Table(details, colWidths=[45 * mm, 120 * mm])
    details_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("PADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(details_table)
    story.append(Spacer(1, 10))

    if comments:
        story.append(Paragraph(f"<b>Comments:</b> {comments}", styles["BodyText"]))
        story.append(Spacer(1, 8))

    table_data = [[
        "Ticker", "ISIN", "Currency", "Ideal Qty", "Default Qty", "AP Input Qty", "Difference", "Cash Adjustment"
    ]]

    for _, row in result_df.iterrows():
        table_data.append([
            str(row["Ticker"]),
            str(row["ISIN"]),
            str(row["Currency"]),
            f"{row['Ideal_Quantity']:,.4f}",
            f"{int(row['Default_AP_Quantity']):,}",
            f"{int(row['AP_Input_Quantity']):,}",
            f"{row['Difference_Qty']:,.4f}",
            f"{row['Cash_Adjustment']:,.2f}",
        ])

    basket_table = Table(table_data, repeatRows=1)
    basket_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("PADDING", (0, 0), (-1, -1), 3),
    ]))

    story.append(Paragraph("Basket Details", styles["Heading2"]))
    story.append(basket_table)
    story.append(Spacer(1, 18))

    changed_rows = result_df[result_df["Changed_By_AP"]].copy()
    if not changed_rows.empty:
        story.append(Paragraph("Rows Changed by AP", styles["Heading2"]))
        changed_table_data = [["Ticker", "Default Qty", "AP Input Qty", "Delta vs Default"]]
        for _, row in changed_rows.iterrows():
            changed_table_data.append([
                str(row["Ticker"]),
                f"{int(row['Default_AP_Quantity']):,}",
                f"{int(row['AP_Input_Quantity']):,}",
                f"{int(row['Delta_vs_Default']):,}",
            ])

        changed_table = Table(changed_table_data, repeatRows=1)
        changed_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("PADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(changed_table)
        story.append(Spacer(1, 12))

    story.append(Paragraph("Authorized Participant Signature: ____________________________", styles["BodyText"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph("ETF Manager Signature: ______________________________________", styles["BodyText"]))

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


# ----------------------------
# Sidebar inputs
# ----------------------------
st.sidebar.header("Request inputs")

all_etfs = sorted(data["pcf_creation"]["ETF"].dropna().unique().tolist())
selected_etf = st.sidebar.selectbox("ETF", all_etfs)

transaction_label = st.sidebar.radio("Transaction type", ["Creation", "Redemption"])

available_dates = get_available_dates_for_etf(selected_etf)
selected_date = st.sidebar.selectbox("Basket date", available_dates)

requested_units = st.sidebar.number_input(
    "ETF units",
    min_value=1.0,
    value=10000.0,
    step=1000.0,
    format="%.0f",
)

allowed_currencies = get_allowed_cash_currencies(selected_etf, selected_date)
selected_cash_currency = st.sidebar.selectbox("Cash currency", allowed_currencies)

st.sidebar.divider()
st.sidebar.subheader("Request details")

ap_name = st.sidebar.text_input("AP name", value="")
request_reference = st.sidebar.text_input("Reference / Order ID", value="")
request_date = st.sidebar.date_input("Request date", value=date.today())
settlement_date = st.sidebar.date_input("Settlement date", value=date.today())
comments = st.sidebar.text_area("Comments", value="")


# ----------------------------
# Main calculation
# ----------------------------
try:
    basket = build_scaled_basket(
        etf=selected_etf,
        basket_date=selected_date,
        transaction_label=transaction_label,
        requested_units=requested_units,
        cash_currency=selected_cash_currency,
    )

    base_cash_10000_converted = get_base_cash_10000(
        etf=selected_etf,
        basket_date=selected_date,
        transaction_label=transaction_label,
        output_currency=selected_cash_currency,
    )
    scale_factor = requested_units / 10000.0
    base_cash_scaled = base_cash_10000_converted * scale_factor

except Exception as exc:
    st.error(str(exc))
    st.stop()


# ----------------------------
# Editable basket state
# ----------------------------
editor_df = basket[
    [
        "Ticker",
        "AP_Input_Quantity",
        "Default_AP_Quantity",
        "Ideal_Quantity",
        "Min_Actual_Integer",
        "Max_Actual_Integer",
        "Currency",
        "Price",
        "ISIN",
    ]
].copy()

state_key = "editable_basket_df"
context_key = "editor_context"

current_context = (
    selected_etf,
    transaction_label,
    selected_date,
    float(requested_units),
    selected_cash_currency,
)

if context_key not in st.session_state or st.session_state[context_key] != current_context:
    st.session_state[context_key] = current_context
    st.session_state[state_key] = editor_df.copy()

if state_key not in st.session_state:
    st.session_state[state_key] = editor_df.copy()


# ----------------------------
# Summary cards
# ----------------------------
col1, col2, col3, col4 = st.columns(4)
col1.metric("ETF", selected_etf)
col2.metric("Type", transaction_label)
col3.metric("Units", f"{requested_units:,.0f}")
col4.metric("Base cash (scaled)", f"{base_cash_scaled:,.2f} {selected_cash_currency}")

st.divider()

st.subheader("Basket")
st.warning(
    "Enter whole-stock quantities only in the input column: **AP Input Qty**. "
    "Any fractional remainder versus the ideal basket is automatically settled in cash."
)

button_col1, button_col2 = st.columns([1, 4])
with button_col1:
    if st.button("Reset to default rounded quantities"):
        reset_df = st.session_state[state_key].copy()
        reset_df["AP_Input_Quantity"] = reset_df["Default_AP_Quantity"]
        st.session_state[state_key] = reset_df
        st.rerun()

edited_df = st.data_editor(
    st.session_state[state_key],
    use_container_width=True,
    hide_index=True,
    disabled=[
        "Ticker",
        "Default_AP_Quantity",
        "Ideal_Quantity",
        "Min_Actual_Integer",
        "Max_Actual_Integer",
        "Currency",
        "Price",
        "ISIN",
    ],
    column_order=[
        "Ticker",
        "AP_Input_Quantity",
        "Default_AP_Quantity",
        "Ideal_Quantity",
        "Min_Actual_Integer",
        "Max_Actual_Integer",
        "Currency",
        "Price",
        "ISIN",
    ],
    column_config={
        "Ticker": st.column_config.TextColumn("Ticker"),
        "AP_Input_Quantity": st.column_config.NumberColumn(
            "AP Input Qty",
            help="Editable whole-number stock quantity entered by the AP.",
            step=1,
            format="%d",
        ),
        "Default_AP_Quantity": st.column_config.NumberColumn(
            "Default Qty",
            help="Default whole-number quantity derived from the ideal basket.",
            format="%d",
        ),
        "Ideal_Quantity": st.column_config.NumberColumn(
            "Ideal Qty",
            help="Exact scaled basket quantity before whole-share rounding.",
            format="%.4f",
        ),
        "Min_Actual_Integer": st.column_config.NumberColumn(
            "Min Allowed",
            format="%d",
        ),
        "Max_Actual_Integer": st.column_config.NumberColumn(
            "Max Allowed",
            format="%d",
        ),
        "Currency": st.column_config.TextColumn("CCY"),
        "Price": st.column_config.NumberColumn("Price", format="%.4f"),
        "ISIN": st.column_config.TextColumn("ISIN"),
    },
    key="basket_editor",
)

st.session_state[state_key] = edited_df.copy()

basket_for_calc = basket.copy()
basket_for_calc["AP_Input_Quantity"] = edited_df["AP_Input_Quantity"]
basket_for_calc["Default_AP_Quantity"] = edited_df["Default_AP_Quantity"]

validation_errors = []
for _, row in basket_for_calc.iterrows():
    try:
        actual_qty = int(row["AP_Input_Quantity"])
    except Exception:
        validation_errors.append(f"{row['Ticker']}: AP input must be a whole number.")
        continue

    if actual_qty != row["AP_Input_Quantity"]:
        validation_errors.append(f"{row['Ticker']}: AP input must be a whole number.")
        continue

    if not validate_actual_vs_bounds(actual_qty, int(row["Min_Actual_Integer"]), int(row["Max_Actual_Integer"])):
        validation_errors.append(
            f"{row['Ticker']}: AP input must stay between "
            f"{int(row['Min_Actual_Integer'])} and {int(row['Max_Actual_Integer'])}."
        )

if validation_errors:
    st.error("Please fix these quantity errors:")
    for err in validation_errors:
        st.write(f"- {err}")
    st.stop()

final_cash, result_df = calculate_cash_result(basket_for_calc, base_cash_scaled)

changed_rows_df = result_df[result_df["Changed_By_AP"]].copy()
changed_count = len(changed_rows_df)

if changed_count == 0:
    st.success("No rows changed by AP. All quantities match the default rounded basket.")
else:
    st.info(f"{changed_count} row(s) changed by AP.")

if changed_count > 0:
    st.subheader("Change Review")
    review_df = changed_rows_df[
        ["Ticker", "Default_AP_Quantity", "AP_Input_Quantity", "Delta_vs_Default"]
    ].copy()

    review_df = review_df.rename(columns={
        "Default_AP_Quantity": "Default Qty",
        "AP_Input_Quantity": "AP Input Qty",
        "Delta_vs_Default": "Delta vs Default",
    })

    st.dataframe(
        review_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Default Qty": st.column_config.NumberColumn(format="%d"),
            "AP Input Qty": st.column_config.NumberColumn(format="%d"),
            "Delta vs Default": st.column_config.NumberColumn(format="%d"),
        },
    )

st.divider()
st.subheader("Request Summary")

summary_col1, summary_col2 = st.columns(2)

with summary_col1:
    st.write(f"**AP Name:** {ap_name if ap_name else '-'}")
    st.write(f"**Reference:** {request_reference if request_reference else '-'}")
    st.write(f"**Request Date:** {request_date}")
    st.write(f"**Settlement Date:** {settlement_date}")

with summary_col2:
    st.write(f"**ETF:** {selected_etf}")
    st.write(f"**Transaction Type:** {transaction_label}")
    st.write(f"**Basket Date:** {selected_date}")
    st.write(f"**Cash Currency:** {selected_cash_currency}")

st.metric("Final cash component", f"{final_cash:,.2f} {selected_cash_currency}")

if final_cash > 0:
    st.info("Positive cash means the AP gives cash to the ETF.")
elif final_cash < 0:
    st.info("Negative cash means the ETF gives cash to the AP.")
else:
    st.info("Cash component is zero.")

st.subheader("Calculation Details")

result_view = result_df[
    [
        "Ticker",
        "ISIN",
        "Currency",
        "Ideal_Quantity",
        "Default_AP_Quantity",
        "AP_Input_Quantity",
        "Difference_Qty",
        "Price_in_Cash_Currency",
        "Cash_Adjustment",
    ]
].copy()

result_view = result_view.rename(columns={
    "Ideal_Quantity": "Ideal Qty",
    "Default_AP_Quantity": "Default Qty",
    "AP_Input_Quantity": "AP Input Qty",
    "Difference_Qty": "Difference Qty",
    "Price_in_Cash_Currency": f"Price in {selected_cash_currency}",
    "Cash_Adjustment": f"Cash Adjustment ({selected_cash_currency})",
})

st.dataframe(
    result_view,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Ideal Qty": st.column_config.NumberColumn(format="%.4f"),
        "Default Qty": st.column_config.NumberColumn(format="%d"),
        "AP Input Qty": st.column_config.NumberColumn(format="%d"),
        "Difference Qty": st.column_config.NumberColumn(format="%.4f"),
        f"Price in {selected_cash_currency}": st.column_config.NumberColumn(format="%.4f"),
        f"Cash Adjustment ({selected_cash_currency})": st.column_config.NumberColumn(format="%.2f"),
    },
)

pdf_bytes = create_pdf(
    ap_name=ap_name,
    request_reference=request_reference,
    request_date=request_date,
    settlement_date=settlement_date,
    comments=comments,
    etf=selected_etf,
    transaction_label=transaction_label,
    basket_date=selected_date,
    units=requested_units,
    cash_currency=selected_cash_currency,
    base_cash_scaled=base_cash_scaled,
    final_cash=final_cash,
    result_df=result_df,
)

st.download_button(
    label="Download PDF request",
    data=pdf_bytes,
    file_name=f"{selected_etf}_{transaction_label}_{int(requested_units)}.pdf",
    mime="application/pdf",
)

st.divider()
st.success("Prototype ready: reset button, AP change review, whole-number stock quantities, cleaner formatting, and downloadable PDF.")