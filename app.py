from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from zipfile import ZipFile
import re
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


MONTHS_PT = {
    "janeiro": 1,
    "jan": 1,
    "fevereiro": 2,
    "fev": 2,
    "marco": 3,
    "março": 3,
    "mar": 3,
    "abril": 4,
    "abr": 4,
    "maio": 5,
    "mai": 5,
    "junho": 6,
    "jun": 6,
    "julho": 7,
    "jul": 7,
    "agosto": 8,
    "ago": 8,
    "setembro": 9,
    "set": 9,
    "outubro": 10,
    "out": 10,
    "novembro": 11,
    "nov": 11,
    "dezembro": 12,
    "dez": 12,
}


@dataclass
class Scenario:
    month: pd.Timestamp
    revenue_liquid: float
    barber_expenses: float
    personal_expenses: float
    saved_balance: float
    product_sales_value: float
    products_sold: float
    services_done: float


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_money(value: object) -> float:
    text = normalize_text(value)
    if not text:
        return 0.0
    text = text.replace("R$", "").replace(" ", "")
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_number(value: object) -> float:
    text = normalize_text(value)
    if not text:
        return 0.0
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return parse_money(text)


def parse_month(value: object) -> pd.Timestamp | pd.NaT:
    text = normalize_text(value).lower().replace(".", "")
    if not text:
        return pd.NaT
    match = re.search(r"([a-zç]+)[/\- ]+(\d{2,4})", text)
    if not match:
        return pd.NaT
    month_name, year_text = match.groups()
    month = MONTHS_PT.get(month_name)
    if not month:
        return pd.NaT
    year = int(year_text)
    if year < 100:
        year += 2000
    return pd.Timestamp(year=year, month=month, day=1)


def column_number(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return 1
    total = 0
    for char in match.group(1):
        total = total * 26 + ord(char) - 64
    return total


def read_xlsx_raw(uploaded_file: BytesIO) -> list[list[object]]:
    """Read first worksheet values directly from XLSX XML.

    Some AppBarber exports contain style records that openpyxl rejects. Reading
    the XML keeps the importer tolerant while still preserving cell values.
    """
    uploaded_file.seek(0)
    with ZipFile(uploaded_file) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("main:si", NS):
                parts = [node.text or "" for node in item.findall(".//main:t", NS)]
                shared_strings.append("".join(parts))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {node.attrib["Id"]: node.attrib["Target"] for node in rels}
        first_sheet = workbook.findall("main:sheets/main:sheet", NS)[0]
        rel_id = first_sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = rel_map[rel_id].lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"

        sheet = ET.fromstring(archive.read(target))
        rows: list[list[object]] = []
        for row in sheet.findall("main:sheetData/main:row", NS):
            values: list[object] = []
            current_col = 0
            for cell in row.findall("main:c", NS):
                cell_ref = cell.attrib.get("r", "A1")
                col = column_number(cell_ref)
                while current_col + 1 < col:
                    values.append(None)
                    current_col += 1
                cell_type = cell.attrib.get("t")
                value_node = cell.find("main:v", NS)
                inline_node = cell.find("main:is", NS)
                value: object = None
                if cell_type == "s" and value_node is not None:
                    value = shared_strings[int(value_node.text or "0")]
                elif cell_type == "inlineStr" and inline_node is not None:
                    value = "".join(node.text or "" for node in inline_node.findall(".//main:t", NS))
                elif value_node is not None:
                    value = value_node.text
                values.append(value)
                current_col = col
            if any(v not in (None, "") for v in values):
                rows.append(values)
        return rows


def rows_to_dataframe(rows: list[list[object]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    headers = [normalize_text(value) or f"Coluna {idx + 1}" for idx, value in enumerate(rows[0])]
    data = rows[1:]
    width = len(headers)
    normalized = [row + [None] * max(0, width - len(row)) for row in data]
    return pd.DataFrame([row[:width] for row in normalized], columns=headers)


def load_uploaded_excel(uploaded_file) -> pd.DataFrame:
    if not uploaded_file:
        return pd.DataFrame()
    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file)
    except Exception:
        rows = read_xlsx_raw(uploaded_file)
        return rows_to_dataframe(rows)


def clean_receipts(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, float]]:
    if df.empty:
        return pd.DataFrame(columns=["Mes", "Total Bruto", "Total Pago", "Total Liquido"]), {}

    working = df.copy()
    working.columns = [normalize_text(col) for col in working.columns]
    first_col = working.columns[0]

    summary: dict[str, float] = {}
    data_rows = []
    for _, row in working.iterrows():
        first = normalize_text(row.get(first_col))
        parsed_month = parse_month(first)
        if pd.notna(parsed_month):
            data_rows.append(
                {
                    "Mes": parsed_month,
                    "Total Bruto": parse_money(row.get(working.columns[1])),
                    "Total Pago": parse_money(row.get(working.columns[2])),
                    "Total Liquido": parse_money(row.get(working.columns[3])),
                }
            )
        else:
            for cell in row.tolist():
                text = normalize_text(cell)
                if ":" in text:
                    label, raw_value = text.split(":", 1)
                    summary[label.strip().title()] = parse_money(raw_value)

    receipts = pd.DataFrame(data_rows)
    if not receipts.empty:
        receipts = receipts.sort_values("Mes")
    return receipts, summary


def clean_ranking(df: pd.DataFrame, name_col: str, qty_col: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[name_col, qty_col])
    working = df.copy()
    working.columns = [normalize_text(col) for col in working.columns]
    source_name = working.columns[0]
    source_qty = working.columns[1] if len(working.columns) > 1 else working.columns[0]
    result = pd.DataFrame(
        {
            name_col: working[source_name].map(normalize_text),
            qty_col: working[source_qty].map(parse_number),
        }
    )
    result = result[result[name_col] != ""]
    return result.sort_values(qty_col, ascending=False)


def format_brl(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_pct(value: float) -> str:
    return f"{value:.1%}".replace(".", ",")


def build_export(
    receipts: pd.DataFrame,
    services: pd.DataFrame,
    products: pd.DataFrame,
    manual_expenses: pd.DataFrame,
    scenario_rows: pd.DataFrame,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        scenario_rows.to_excel(writer, index=False, sheet_name="Cenario")
        receipts.to_excel(writer, index=False, sheet_name="Recebimentos")
        services.to_excel(writer, index=False, sheet_name="Ranking_Servicos")
        products.to_excel(writer, index=False, sheet_name="Ranking_Produtos")
        manual_expenses.to_excel(writer, index=False, sheet_name="Despesas_Manuais")
    return output.getvalue()


st.set_page_config(page_title="Painel AppBarber", layout="wide")

st.title("Painel AppBarber")
st.caption("Suba os relatórios exportados, preencha despesas/valor guardado e veja o cenário atual.")

with st.sidebar:
    st.header("Relatórios AppBarber")
    receb_file = st.file_uploader("Total recebimento período", type=["xlsx"], key="receb")
    serv_file = st.file_uploader("Ranking serviços", type=["xlsx"], key="serv")
    prod_file = st.file_uploader("Ranking produtos", type=["xlsx"], key="prod")

    st.header("Metas")
    revenue_goal = st.number_input("Meta faturamento mensal", min_value=0.0, value=9000.0, step=100.0)
    expense_goal = st.number_input("Despesa saudável até (%)", min_value=0.0, max_value=100.0, value=42.0, step=1.0) / 100
    product_goal = st.number_input("Produtos no faturamento (%)", min_value=0.0, max_value=100.0, value=8.0, step=0.5) / 100
    cash_goal = st.number_input("Caixa livre mínimo", min_value=0.0, value=3000.0, step=100.0)


receipts_raw = load_uploaded_excel(receb_file)
services_raw = load_uploaded_excel(serv_file)
products_raw = load_uploaded_excel(prod_file)

receipts, receipt_summary = clean_receipts(receipts_raw)
services = clean_ranking(services_raw, "Servico", "Realizado")
products = clean_ranking(products_raw, "Produto", "Vendidos")

if receipts.empty:
    st.info("Comece subindo o relatório `total recebimento periodo.xlsx`.")
    st.stop()

month_options = receipts["Mes"].dt.strftime("%b/%Y").tolist()
selected_label = st.selectbox("Mês analisado", month_options, index=len(month_options) - 1)
selected_month = receipts.loc[receipts["Mes"].dt.strftime("%b/%Y") == selected_label, "Mes"].iloc[0]
selected_receipt = receipts[receipts["Mes"] == selected_month].iloc[0]

st.subheader("Entradas manuais do mês")
col1, col2, col3, col4 = st.columns(4)
with col1:
    fixed_barber = st.number_input("Fixas barbearia", min_value=0.0, value=0.0, step=50.0)
with col2:
    variable_barber = st.number_input("Variáveis barbearia", min_value=0.0, value=0.0, step=50.0)
with col3:
    one_off_barber = st.number_input("Investimentos/pontuais", min_value=0.0, value=0.0, step=50.0)
with col4:
    personal_expenses = st.number_input("Despesas pessoais", min_value=0.0, value=0.0, step=50.0)

saved_balance = st.number_input("Valor guardado na conta neste mês", min_value=0.0, value=0.0, step=100.0)

barber_expenses = fixed_barber + variable_barber + one_off_barber
revenue_liquid = float(selected_receipt["Total Liquido"])
cash_free = revenue_liquid - barber_expenses
cash_after_saving = cash_free - saved_balance
expense_pct = barber_expenses / revenue_liquid if revenue_liquid else 0.0
products_value = receipt_summary.get("Produtos", 0.0)
period_revenue = receipts["Total Bruto"].sum()
products_share = products_value / period_revenue if period_revenue else 0.0
products_sold = products["Vendidos"].sum() if not products.empty else 0.0
services_done = services["Realizado"].sum() if not services.empty else 0.0

st.subheader("Cenário atual")
k1, k2, k3, k4 = st.columns(4)
k1.metric("Faturamento líquido", format_brl(revenue_liquid), "Meta " + format_brl(revenue_goal))
k2.metric("Despesas barbearia", format_brl(barber_expenses), format_pct(expense_pct))
k3.metric("Caixa livre", format_brl(cash_free), "antes de guardar")
k4.metric("Caixa após guardar", format_brl(cash_after_saving), "guardado " + format_brl(saved_balance))

k5, k6, k7, k8 = st.columns(4)
k5.metric("Produtos / faturamento", format_pct(products_share), "Meta " + format_pct(product_goal))
k6.metric("Produtos vendidos", f"{products_sold:,.0f}".replace(",", "."))
k7.metric("Serviços realizados", f"{services_done:,.0f}".replace(",", "."))
k8.metric("Despesas pessoais", format_brl(personal_expenses))

alerts = []
if revenue_liquid < revenue_goal:
    alerts.append(f"Faturamento abaixo da meta: faltam {format_brl(revenue_goal - revenue_liquid)}.")
if expense_pct > expense_goal:
    alerts.append(f"Despesas da barbearia acima do limite saudável de {format_pct(expense_goal)}.")
if products_share < product_goal:
    alerts.append(f"Produtos estão em {format_pct(products_share)} do faturamento do período; meta é {format_pct(product_goal)}.")
if cash_after_saving < cash_goal:
    alerts.append(f"Caixa após guardar ficou abaixo do mínimo de {format_brl(cash_goal)}.")

if alerts:
    st.warning("\n\n".join(alerts))
else:
    st.success("Cenário saudável dentro das metas configuradas.")

tab_overview, tab_rankings, tab_data = st.tabs(["Gráficos", "Rankings", "Dados"])

with tab_overview:
    chart_receipts = receipts.set_index(receipts["Mes"].dt.strftime("%b/%Y"))[["Total Liquido"]]
    st.bar_chart(chart_receipts)

    composition = pd.DataFrame(
        {
            "Valor": [barber_expenses, personal_expenses, saved_balance],
        },
        index=["Despesas barbearia", "Despesas pessoais", "Valor guardado"],
    )
    st.bar_chart(composition)

with tab_rankings:
    left, right = st.columns(2)
    with left:
        st.markdown("**Top serviços**")
        st.dataframe(services.head(10), use_container_width=True, hide_index=True)
        if not services.empty:
            st.bar_chart(services.head(8).set_index("Servico")[["Realizado"]])
    with right:
        st.markdown("**Top produtos**")
        st.dataframe(products.head(10), use_container_width=True, hide_index=True)
        if not products.empty:
            st.bar_chart(products.head(8).set_index("Produto")[["Vendidos"]])

with tab_data:
    st.markdown("**Recebimentos limpos**")
    st.dataframe(receipts, use_container_width=True, hide_index=True)
    st.markdown("**Resumo do período encontrado no export**")
    st.json(receipt_summary)

scenario_df = pd.DataFrame(
    [
        {"Indicador": "Mes analisado", "Valor": selected_month.strftime("%m/%Y")},
        {"Indicador": "Faturamento liquido", "Valor": revenue_liquid},
        {"Indicador": "Despesas barbearia", "Valor": barber_expenses},
        {"Indicador": "% despesas", "Valor": expense_pct},
        {"Indicador": "Caixa livre", "Valor": cash_free},
        {"Indicador": "Valor guardado", "Valor": saved_balance},
        {"Indicador": "Caixa apos guardar", "Valor": cash_after_saving},
        {"Indicador": "Despesas pessoais", "Valor": personal_expenses},
        {"Indicador": "Produtos / faturamento", "Valor": products_share},
        {"Indicador": "Produtos vendidos", "Valor": products_sold},
        {"Indicador": "Servicos realizados", "Valor": services_done},
    ]
)
manual_df = pd.DataFrame(
    [
        {
            "Mes": selected_month,
            "Fixas Barbearia": fixed_barber,
            "Variaveis Barbearia": variable_barber,
            "Investimentos/Pontuais": one_off_barber,
            "Total Barbearia": barber_expenses,
            "Pessoais": personal_expenses,
            "Valor guardado": saved_balance,
            "Caixa livre": cash_free,
            "Caixa apos guardar": cash_after_saving,
        }
    ]
)

export_bytes = build_export(receipts, services, products, manual_df, scenario_df)
st.download_button(
    "Baixar cenário em Excel",
    data=export_bytes,
    file_name=f"cenario_appbarber_{selected_month.strftime('%Y_%m')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
