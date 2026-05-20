from __future__ import annotations

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

HISTORY_COLUMNS = [
    "Mes",
    "Faturamento Liquido",
    "Fixas Barbearia",
    "Variaveis Barbearia",
    "Investimentos/Pontuais",
    "Total Barbearia",
    "Despesas Pessoais",
    "Valor Guardado",
    "Caixa Livre",
    "Caixa Apos Guardar",
    "% Despesas",
    "Produtos Valor",
    "% Produtos",
    "Produtos Vendidos",
    "Servicos Realizados",
    "Status",
]

ACCOUNT_COLUMNS = [
    "Centro",
    "Tipo",
    "Descricao",
    "Valor Mensal",
    "Vencimento",
    "Forma Pagamento",
    "Mes Inicial",
    "Mes Final",
    "Status",
    "Observacoes",
]


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
    if isinstance(value, pd.Timestamp):
        return pd.Timestamp(year=value.year, month=value.month, day=1)
    text = normalize_text(value).lower().replace(".", "")
    if not text:
        return pd.NaT
    match = re.search(r"([a-zç]+)[/\- ]+(\d{2,4})", text)
    if not match:
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
        if pd.isna(parsed):
            return pd.NaT
        return pd.Timestamp(year=parsed.year, month=parsed.month, day=1)
    month_name, year_text = match.groups()
    month = MONTHS_PT.get(month_name)
    if not month:
        return pd.NaT
    year = int(year_text)
    if year < 100:
        year += 2000
    return pd.Timestamp(year=year, month=month, day=1)


def month_label(value: object) -> str:
    parsed = parse_month(value)
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%m/%Y")


def format_brl(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_pct(value: float) -> str:
    return f"{value:.1%}".replace(".", ",")


def column_number(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        return 1
    total = 0
    for char in match.group(1):
        total = total * 26 + ord(char) - 64
    return total


def read_xlsx_raw(uploaded_file) -> list[list[object]]:
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
        return rows_to_dataframe(read_xlsx_raw(uploaded_file))


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


def default_accounts() -> pd.DataFrame:
    rows = [
        ["Barbearia", "Fixo", "Internet e Spotify", 276.45, "", "Pix", "05/2026", "", "Ativo", ""],
        ["Barbearia", "Parcela", "MEI", 390.00, "Dia 25", "Boleto", "05/2026", "", "Ativo", ""],
        ["Barbearia", "Fixo", "App barber", 109.90, "02/03/2026", "Pix", "05/2026", "", "Ativo", ""],
        ["Barbearia", "Fixo", "Laura", 320.00, "Dia 10", "Pix", "05/2026", "", "Ativo", ""],
        ["Barbearia", "Fixo", "Luz", 200.00, "", "Pix", "05/2026", "", "Ativo", ""],
        ["Barbearia", "Fixo", "Marketing", 1750.00, "", "Pix", "05/2026", "", "Ativo", ""],
        ["Barbearia", "Parcela", "Mesinha", 63.93, "Dia 16", "Pix", "05/2026", "", "Ativo", "4 parcelas"],
        ["Barbearia", "Fixo", "Maquina Cafe", 256.00, "", "Pix", "05/2026", "", "Ativo", ""],
        ["Barbearia", "Parcela", "Maca", 180.00, "", "Pix", "05/2026", "", "Ativo", "5 parcelas"],
        ["Barbearia", "Parcela", "Ar condicionado", 213.00, "", "Pix", "05/2026", "", "Ativo", "10 parcelas"],
        ["Barbearia", "Parcela", "Papo de Barbeira", 99.00, "Dia 30", "Pix", "05/2026", "", "Ativo", "3 parcelas"],
        ["Pessoal", "Fixo", "Vo", 480.00, "", "Pix", "05/2026", "", "Ativo", "120 por semana"],
        ["Pessoal", "Fixo", "Unimed Pedro", 300.00, "", "Pix", "05/2026", "", "Ativo", ""],
        ["Pessoal", "Fixo", "Cartao de todos", 40.00, "", "Pix", "05/2026", "", "Ativo", ""],
        ["Pessoal", "Parcela", "IPTU", 80.00, "", "Pix", "05/2026", "", "Ativo", "11 parcelas"],
        ["Pessoal", "Fixo", "Psicologo", 180.00, "", "Pix", "05/2026", "", "Ativo", "45 por semana"],
    ]
    return pd.DataFrame(rows, columns=ACCOUNT_COLUMNS)


def load_history_base(uploaded_file) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not uploaded_file:
        return pd.DataFrame(columns=HISTORY_COLUMNS), default_accounts()
    uploaded_file.seek(0)
    try:
        history = pd.read_excel(uploaded_file, sheet_name="Historico")
    except Exception:
        history = pd.DataFrame(columns=HISTORY_COLUMNS)
    uploaded_file.seek(0)
    try:
        accounts = pd.read_excel(uploaded_file, sheet_name="Contas")
    except Exception:
        accounts = default_accounts()
    history = normalize_history(history)
    accounts = normalize_accounts(accounts)
    return history, accounts


def normalize_history(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=HISTORY_COLUMNS)
    result = df.copy()
    for col in HISTORY_COLUMNS:
        if col not in result.columns:
            result[col] = None
    result = result[HISTORY_COLUMNS]
    result["Mes"] = result["Mes"].map(parse_month)
    for col in HISTORY_COLUMNS:
        if col not in ["Mes", "Status"]:
            result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0.0)
    return result.dropna(subset=["Mes"]).sort_values("Mes")


def normalize_accounts(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return default_accounts()
    result = df.copy()
    for col in ACCOUNT_COLUMNS:
        if col not in result.columns:
            result[col] = ""
    result = result[ACCOUNT_COLUMNS]
    result["Valor Mensal"] = pd.to_numeric(result["Valor Mensal"], errors="coerce").fillna(0.0)
    result["Status"] = result["Status"].replace("", "Ativo")
    return result


def active_accounts_for_month(accounts: pd.DataFrame, month: pd.Timestamp) -> pd.DataFrame:
    if accounts.empty:
        return accounts
    working = accounts.copy()
    start = working["Mes Inicial"].map(parse_month)
    end = working["Mes Final"].map(parse_month)
    status_ok = working["Status"].fillna("Ativo").eq("Ativo")
    starts_before = start.isna() | (start <= month)
    ends_after = end.isna() | (end >= month)
    return working[status_ok & starts_before & ends_after]


def status_label(revenue: float, expense_pct: float, cash_after_saving: float, revenue_goal: float, expense_goal: float, cash_goal: float) -> str:
    if revenue < revenue_goal * 0.9 or expense_pct > expense_goal * 1.15 or cash_after_saving < cash_goal * 0.5:
        return "Critico"
    if revenue < revenue_goal or expense_pct > expense_goal or cash_after_saving < cash_goal:
        return "Atencao"
    return "OK"


def build_export(history: pd.DataFrame, accounts: pd.DataFrame, receipts: pd.DataFrame, services: pd.DataFrame, products: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        history.to_excel(writer, index=False, sheet_name="Historico")
        accounts.to_excel(writer, index=False, sheet_name="Contas")
        receipts.to_excel(writer, index=False, sheet_name="Recebimentos_AppBarber")
        services.to_excel(writer, index=False, sheet_name="Ranking_Servicos")
        products.to_excel(writer, index=False, sheet_name="Ranking_Produtos")
    return output.getvalue()


st.set_page_config(page_title="Painel AppBarber", layout="wide")
st.title("Painel AppBarber")
st.caption("Histórico mensal, contas recorrentes, despesas pessoais, parcelas e evolução do negócio.")

with st.sidebar:
    st.header("Base histórica")
    base_file = st.file_uploader("Base anterior do painel (.xlsx)", type=["xlsx"], key="base")
    history_base, accounts_base = load_history_base(base_file)

    st.header("Relatórios do mês")
    receb_file = st.file_uploader("Total recebimento período", type=["xlsx"], key="receb")
    serv_file = st.file_uploader("Ranking serviços", type=["xlsx"], key="serv")
    prod_file = st.file_uploader("Ranking produtos", type=["xlsx"], key="prod")

    st.header("Metas")
    revenue_goal = st.number_input("Meta faturamento mensal", min_value=0.0, value=9000.0, step=100.0)
    expense_goal = st.number_input("Despesa saudável até (%)", min_value=0.0, max_value=100.0, value=42.0, step=1.0) / 100
    product_goal = st.number_input("Produtos no faturamento (%)", min_value=0.0, max_value=100.0, value=8.0, step=0.5) / 100
    cash_goal = st.number_input("Caixa após guardar mínimo", min_value=0.0, value=3000.0, step=100.0)

receipts, receipt_summary = clean_receipts(load_uploaded_excel(receb_file))
services = clean_ranking(load_uploaded_excel(serv_file), "Servico", "Realizado")
products = clean_ranking(load_uploaded_excel(prod_file), "Produto", "Vendidos")

if "accounts_editor" not in st.session_state:
    st.session_state.accounts_editor = accounts_base
if base_file and "base_loaded_name" not in st.session_state:
    st.session_state.accounts_editor = accounts_base
    st.session_state.base_loaded_name = base_file.name

accounts = st.session_state.accounts_editor
history = history_base.copy()

if receipts.empty:
    st.info("Suba o relatório `total recebimento periodo.xlsx` para calcular o mês atual. Você já pode editar as contas na aba Contas.")

month_options = receipts["Mes"].dt.strftime("%m/%Y").tolist() if not receipts.empty else [pd.Timestamp.today().strftime("%m/%Y")]
selected_label = st.selectbox("Mês analisado", month_options, index=len(month_options) - 1)
selected_month = parse_month(selected_label)
selected_receipt = receipts[receipts["Mes"].eq(selected_month)].iloc[0] if not receipts.empty and receipts["Mes"].eq(selected_month).any() else None

tab_current, tab_accounts, tab_history, tab_rankings, tab_export = st.tabs(
    ["Cenário do mês", "Contas e parcelas", "Histórico e evolução", "Rankings", "Exportar base"]
)

with tab_accounts:
    st.subheader("Cadastro de contas fixas, parcelas e despesas pessoais")
    st.caption("Use Status = Ativo para contar no mês. Use Mes Inicial e Mes Final no formato mm/aaaa.")
    edited_accounts = st.data_editor(
        accounts,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Centro": st.column_config.SelectboxColumn("Centro", options=["Barbearia", "Pessoal"], required=True),
            "Tipo": st.column_config.SelectboxColumn("Tipo", options=["Fixo", "Variavel", "Parcela"], required=True),
            "Valor Mensal": st.column_config.NumberColumn("Valor Mensal", min_value=0.0, step=10.0, format="R$ %.2f"),
            "Forma Pagamento": st.column_config.SelectboxColumn("Forma Pagamento", options=["Pix", "Boleto", "Cartao", "Dinheiro", "Outro"]),
            "Status": st.column_config.SelectboxColumn("Status", options=["Ativo", "Pausado", "Quitado"], required=True),
        },
        key="accounts_grid",
    )
    st.session_state.accounts_editor = normalize_accounts(edited_accounts)
    accounts = st.session_state.accounts_editor
    active_accounts = active_accounts_for_month(accounts, selected_month)
    c1, c2, c3 = st.columns(3)
    c1.metric("Barbearia ativo", format_brl(active_accounts.loc[active_accounts["Centro"].eq("Barbearia"), "Valor Mensal"].sum()))
    c2.metric("Pessoal ativo", format_brl(active_accounts.loc[active_accounts["Centro"].eq("Pessoal"), "Valor Mensal"].sum()))
    c3.metric("Total ativo", format_brl(active_accounts["Valor Mensal"].sum()))

with tab_current:
    st.subheader("Cenário do mês")
    active_accounts = active_accounts_for_month(accounts, selected_month)
    suggested_fixed_barber = active_accounts[
        active_accounts["Centro"].eq("Barbearia") & active_accounts["Tipo"].isin(["Fixo", "Parcela"])
    ]["Valor Mensal"].sum()
    suggested_personal = active_accounts[active_accounts["Centro"].eq("Pessoal")]["Valor Mensal"].sum()

    col1, col2, col3, col4 = st.columns(4)
    fixed_barber = col1.number_input("Fixas/parcelas barbearia", min_value=0.0, value=float(suggested_fixed_barber), step=50.0)
    variable_barber = col2.number_input("Variáveis barbearia", min_value=0.0, value=0.0, step=50.0)
    one_off_barber = col3.number_input("Investimentos/pontuais", min_value=0.0, value=0.0, step=50.0)
    personal_expenses = col4.number_input("Despesas pessoais", min_value=0.0, value=float(suggested_personal), step=50.0)
    saved_balance = st.number_input("Valor guardado na conta neste mês", min_value=0.0, value=0.0, step=100.0)

    revenue_liquid = float(selected_receipt["Total Liquido"]) if selected_receipt is not None else 0.0
    barber_expenses = fixed_barber + variable_barber + one_off_barber
    cash_free = revenue_liquid - barber_expenses
    cash_after_saving = cash_free - saved_balance
    expense_pct = barber_expenses / revenue_liquid if revenue_liquid else 0.0
    products_value = receipt_summary.get("Produtos", 0.0)
    period_revenue = receipts["Total Bruto"].sum() if not receipts.empty else 0.0
    products_share = products_value / period_revenue if period_revenue else 0.0
    products_sold = products["Vendidos"].sum() if not products.empty else 0.0
    services_done = services["Realizado"].sum() if not services.empty else 0.0
    status = status_label(revenue_liquid, expense_pct, cash_after_saving, revenue_goal, expense_goal, cash_goal)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Faturamento líquido", format_brl(revenue_liquid), "Meta " + format_brl(revenue_goal))
    k2.metric("Despesas barbearia", format_brl(barber_expenses), format_pct(expense_pct))
    k3.metric("Caixa livre", format_brl(cash_free), "antes de guardar")
    k4.metric("Caixa após guardar", format_brl(cash_after_saving), "guardado " + format_brl(saved_balance))

    k5, k6, k7, k8 = st.columns(4)
    k5.metric("Produtos / faturamento", format_pct(products_share), "Meta " + format_pct(product_goal))
    k6.metric("Produtos vendidos", f"{products_sold:,.0f}".replace(",", "."))
    k7.metric("Serviços realizados", f"{services_done:,.0f}".replace(",", "."))
    k8.metric("Status", status)

    current_row = pd.DataFrame(
        [
            {
                "Mes": selected_month,
                "Faturamento Liquido": revenue_liquid,
                "Fixas Barbearia": fixed_barber,
                "Variaveis Barbearia": variable_barber,
                "Investimentos/Pontuais": one_off_barber,
                "Total Barbearia": barber_expenses,
                "Despesas Pessoais": personal_expenses,
                "Valor Guardado": saved_balance,
                "Caixa Livre": cash_free,
                "Caixa Apos Guardar": cash_after_saving,
                "% Despesas": expense_pct,
                "Produtos Valor": products_value,
                "% Produtos": products_share,
                "Produtos Vendidos": products_sold,
                "Servicos Realizados": services_done,
                "Status": status,
            }
        ]
    )

    history_without_month = history[~history["Mes"].eq(selected_month)] if not history.empty else history
    updated_history = pd.concat([history_without_month, current_row], ignore_index=True)
    updated_history = normalize_history(updated_history)

    if status == "OK":
        st.success("Cenário saudável dentro das metas configuradas.")
    elif status == "Atencao":
        st.warning("Cenário pede atenção. Revise faturamento, despesas ou valor guardado.")
    else:
        st.error("Cenário crítico. Priorize caixa, despesas e meta de faturamento.")

with tab_history:
    st.subheader("Histórico e evolução")
    if "updated_history" not in locals():
        updated_history = history
    if updated_history.empty:
        st.info("Ainda não há histórico. Calcule um mês e baixe a base atualizada.")
    else:
        view = updated_history.copy()
        view["Mes Label"] = view["Mes"].dt.strftime("%m/%Y")
        st.dataframe(view.drop(columns=["Mes"]).rename(columns={"Mes Label": "Mes"}), use_container_width=True, hide_index=True)
        chart = view.set_index("Mes Label")[["Faturamento Liquido", "Total Barbearia", "Valor Guardado", "Caixa Apos Guardar"]]
        st.line_chart(chart)
        st.bar_chart(view.set_index("Mes Label")[["% Despesas", "% Produtos"]])

with tab_rankings:
    left, right = st.columns(2)
    with left:
        st.subheader("Ranking de serviços")
        st.dataframe(services, use_container_width=True, hide_index=True)
        if not services.empty:
            st.bar_chart(services.head(10).set_index("Servico")[["Realizado"]])
    with right:
        st.subheader("Ranking de produtos")
        st.dataframe(products, use_container_width=True, hide_index=True)
        if not products.empty:
            st.bar_chart(products.head(10).set_index("Produto")[["Vendidos"]])

with tab_export:
    st.subheader("Exportar base atualizada")
    st.caption("Baixe este arquivo e use como `Base anterior do painel` na próxima atualização mensal.")
    if "updated_history" not in locals():
        updated_history = history
    export_bytes = build_export(updated_history, accounts, receipts, services, products)
    st.download_button(
        "Baixar base histórica atualizada",
        data=export_bytes,
        file_name=f"base_painel_appbarber_{month_label(selected_month).replace('/', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
