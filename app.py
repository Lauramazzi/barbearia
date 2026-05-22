"""
Painel de Gestão – Barbearia Heloisa Mazzi
==========================================
Funcionalidades:
- Dashboard do mês atual com KPIs
- Registro manual de saídas (Heloisa pode lançar diretamente)
- Histórico mês a mês com gráficos de evolução
- Metas progressivas mensais
- Rankings de serviços e produtos
- Relatório PDF de fechamento mensal
- Base histórica persistente (download/upload)
"""

from __future__ import annotations

import io
import re
import json
import datetime
import xml.etree.ElementTree as ET
from io import BytesIO
from zipfile import ZipFile

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st
from fpdf import FPDF

# ──────────────────────────────────────────────────────────────
# CONSTANTES
# ──────────────────────────────────────────────────────────────

MESES_PT = {
    "janeiro": 1, "jan": 1, "fevereiro": 2, "fev": 2,
    "marco": 3, "março": 3, "mar": 3, "abril": 4, "abr": 4,
    "maio": 5, "mai": 5, "junho": 6, "jun": 6,
    "julho": 7, "jul": 7, "agosto": 8, "ago": 8,
    "setembro": 9, "set": 9, "outubro": 10, "out": 10,
    "novembro": 11, "nov": 11, "dezembro": 12, "dez": 12,
}

NOMES_MESES = [
    "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

COLUNAS_HISTORICO = [
    "Mes", "Faturamento", "Total Despesas", "Despesas Barbearia",
    "Despesas Pessoais", "Valor Guardado", "Saldo Livre",
    "Meta Faturamento", "Meta Despesas %", "Atingiu Meta",
    "Produtos Valor", "Produtos Vendidos", "Servicos Realizados",
    "Status",
]

COLUNAS_SAIDAS = [
    "Mes", "Data", "Centro", "Categoria", "Tipo",
    "Descricao", "Valor", "Forma Pagamento",
    "Parcela Atual", "Total Parcelas", "Status", "Observacoes",
]

COLUNAS_METAS = [
    "Mes", "Meta Faturamento", "Meta Despesas %",
    "Meta Valor Guardado", "Meta Saldo Livre",
    "Meta Servicos", "Meta Produtos Vendidos",
]

CATEGORIAS_BARBEARIA = [
    "Marketing", "Funcionários", "Energia/Água", "Internet/Streaming",
    "Equipamentos", "Produtos/Insumos", "Aluguel", "Sistema/App",
    "Manutenção", "MEI/Impostos", "Parcela", "Outros Barbearia",
]

CATEGORIAS_PESSOAL = [
    "Alimentação", "Transporte", "Saúde", "Educação",
    "Lazer", "Casa/Moradia", "Roupas", "Farmácia",
    "Impostos Pessoais", "Outros Pessoal",
]

FORMAS_PAGTO = ["Pix", "Débito", "Crédito", "Boleto", "Dinheiro", "Outro"]

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# ──────────────────────────────────────────────────────────────
# UTILS
# ──────────────────────────────────────────────────────────────

def formatar_brl(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_pct(v: float) -> str:
    return f"{v:.1f}%"

def parse_money(value) -> float:
    text = str(value or "").replace("R$", "").replace(" ", "").strip()
    if not text or text in ("-", "nan"):
        return 0.0
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0

def parse_mes(value) -> pd.Timestamp | None:
    if isinstance(value, pd.Timestamp):
        return pd.Timestamp(year=value.year, month=value.month, day=1)
    if isinstance(value, datetime.date):
        return pd.Timestamp(year=value.year, month=value.month, day=1)
    text = str(value or "").strip().lower().replace(".", "")
    if not text or text == "nan":
        return None
    # Formato ISO: 2026-05-01
    m_iso = re.match(r"(\d{4})-(\d{2})-(\d{2})", text)
    if m_iso:
        return pd.Timestamp(year=int(m_iso.group(1)), month=int(m_iso.group(2)), day=1)
    m = re.search(r"([a-zç]+)[/\- ]+(\d{2,4})", text)
    if m:
        nome, ano = m.groups()
        mes_num = MESES_PT.get(nome)
        if mes_num:
            ano_int = int(ano)
            if ano_int < 100:
                ano_int += 2000
            return pd.Timestamp(year=ano_int, month=mes_num, day=1)
    m2 = re.search(r"(\d{1,2})[/\-](\d{2,4})", text)
    if m2:
        mes_num, ano = int(m2.group(1)), int(m2.group(2))
        if ano < 100:
            ano += 2000
        if 1 <= mes_num <= 12:
            return pd.Timestamp(year=ano, month=mes_num, day=1)
    try:
        p = pd.to_datetime(value, errors="coerce", dayfirst=True)
        if pd.notna(p):
            return pd.Timestamp(year=p.year, month=p.month, day=1)
    except Exception:
        pass
    return None

def label_mes(value) -> str:
    p = parse_mes(value)
    if p is None:
        return ""
    return p.strftime("%m/%Y")

def mes_extenso(value) -> str:
    p = parse_mes(value)
    if p is None:
        return ""
    return f"{NOMES_MESES[p.month]}/{p.year}"

def mes_atual_ts() -> pd.Timestamp:
    today = datetime.date.today()
    return pd.Timestamp(year=today.year, month=today.month, day=1)

def cor_status(status: str) -> str:
    return {"OK": "🟢", "Atenção": "🟡", "Crítico": "🔴"}.get(status, "⚪")

# ──────────────────────────────────────────────────────────────
# LEITURA DE XLSX (compatível com AppBarber)
# ──────────────────────────────────────────────────────────────

def _col_num(ref: str) -> int:
    m = re.match(r"([A-Z]+)", ref)
    if not m:
        return 1
    total = 0
    for c in m.group(1):
        total = total * 26 + ord(c) - 64
    return total

def ler_xlsx_raw(file_bytes: bytes) -> list[list]:
    with ZipFile(io.BytesIO(file_bytes)) as arc:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in arc.namelist():
            root = ET.fromstring(arc.read("xl/sharedStrings.xml"))
            for item in root.findall("main:si", NS):
                parts = [n.text or "" for n in item.findall(".//main:t", NS)]
                shared.append("".join(parts))
        wb = ET.fromstring(arc.read("xl/workbook.xml"))
        rels = ET.fromstring(arc.read("xl/_rels/workbook.xml.rels"))
        rel_map = {n.attrib["Id"]: n.attrib["Target"] for n in rels}
        first = wb.findall("main:sheets/main:sheet", NS)[0]
        rid = first.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = rel_map[rid].lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        sheet = ET.fromstring(arc.read(target))
        rows = []
        for row in sheet.findall("main:sheetData/main:row", NS):
            vals, curr = [], 0
            for cell in row.findall("main:c", NS):
                ref = cell.attrib.get("r", "A1")
                col = _col_num(ref)
                while curr + 1 < col:
                    vals.append(None)
                    curr += 1
                ct = cell.attrib.get("t")
                vn = cell.find("main:v", NS)
                val = None
                if ct == "s" and vn is not None:
                    val = shared[int(vn.text or "0")]
                elif vn is not None:
                    val = vn.text
                vals.append(val)
                curr = col
            if any(v not in (None, "") for v in vals):
                rows.append(vals)
        return rows

def ler_excel_upload(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    file.seek(0)
    data = file.read()
    try:
        file.seek(0)
        return pd.read_excel(file)
    except Exception:
        pass
    try:
        rows = ler_xlsx_raw(data)
        if not rows:
            return pd.DataFrame()
        headers = [str(v or f"Col{i}").strip() for i, v in enumerate(rows[0])]
        return pd.DataFrame(rows[1:], columns=headers)
    except Exception:
        return pd.DataFrame()

def limpar_ranking(df: pd.DataFrame, col_nome: str, col_qtd: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[col_nome, col_qtd])
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    c0, c1 = df.columns[0], df.columns[1] if len(df.columns) > 1 else df.columns[0]
    result = pd.DataFrame({
        col_nome: df[c0].astype(str).str.strip(),
        col_qtd: df[c1].apply(lambda x: float(str(x).replace(",", ".") or 0) if x else 0),
    })
    result = result[result[col_nome].str.len() > 0]
    return result.sort_values(col_qtd, ascending=False).reset_index(drop=True)

def processar_recebimentos(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Mes", "Faturamento"])
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    c0 = df.columns[0]
    rows = []
    for _, row in df.iterrows():
        mes = parse_mes(row.iloc[0])
        if mes is not None:
            val = parse_money(row.iloc[1]) if len(row) > 1 else 0.0
            rows.append({"Mes": mes, "Faturamento": val})
    return pd.DataFrame(rows).sort_values("Mes") if rows else pd.DataFrame(columns=["Mes", "Faturamento"])

# ──────────────────────────────────────────────────────────────
# DADOS DEFAULT
# ──────────────────────────────────────────────────────────────

def saidas_default() -> pd.DataFrame:
    """Popula saídas de Maio/2026 com base no relatório PDF."""
    mes = pd.Timestamp(2026, 5, 1)
    rows = [
        # Fixas
        (mes, "01/05/2026", "Barbearia", "Sistema/App", "Fixo", "App Barber", 109.90, "Pix", "", "", "Pago", ""),
        (mes, "10/05/2026", "Barbearia", "Funcionários", "Fixo", "Laura", 320.00, "Pix", "", "", "Pago", ""),
        (mes, "01/05/2026", "Barbearia", "Internet/Streaming", "Fixo", "Internet e Spotify", 276.45, "Pix", "", "", "Pago", ""),
        (mes, "01/05/2026", "Barbearia", "Equipamentos", "Fixo", "Máquina de Café", 256.00, "Pix", "", "", "Pago", ""),
        (mes, "01/05/2026", "Barbearia", "Energia/Água", "Fixo", "Luz", 200.00, "Pix", "", "", "Pago", ""),
        (mes, "01/05/2026", "Barbearia", "Marketing", "Fixo", "Marketing", 1750.00, "Pix", "", "", "Pago", ""),
        # Parcelas
        (mes, "25/05/2026", "Barbearia", "MEI/Impostos", "Parcela", "MEI", 390.00, "Boleto", "", "", "Pago", "Recorrente"),
        (mes, "01/05/2026", "Barbearia", "Equipamentos", "Parcela", "Ar Condicionado", 213.00, "Pix", "1", "10", "Pago", ""),
        (mes, "01/05/2026", "Barbearia", "Equipamentos", "Parcela", "Maca", 180.00, "Pix", "1", "5", "Pago", ""),
        (mes, "30/05/2026", "Barbearia", "Educação", "Parcela", "Papo de Barbeira", 99.00, "Pix", "1", "3", "Pago", ""),
        (mes, "16/05/2026", "Barbearia", "Equipamentos", "Parcela", "Mesinha", 63.93, "Pix", "1", "4", "Pago", ""),
        # Pessoal
        (mes, "01/05/2026", "Pessoal", "Casa/Moradia", "Fixo", "Vó (semanas)", 480.00, "Pix", "", "", "Pago", "R$120/semana"),
        (mes, "01/05/2026", "Pessoal", "Saúde", "Fixo", "Unimed Pedro", 300.00, "Pix", "", "", "Pago", ""),
        (mes, "01/05/2026", "Pessoal", "Outros Pessoal", "Fixo", "Cartão de Todos", 40.00, "Pix", "", "", "Pago", ""),
        (mes, "01/05/2026", "Pessoal", "Impostos Pessoais", "Parcela", "IPTU", 80.00, "Pix", "1", "11", "Pago", ""),
        (mes, "01/05/2026", "Pessoal", "Saúde", "Fixo", "Psicólogo", 180.00, "Pix", "", "", "Pago", "R$45/semana"),
    ]
    return pd.DataFrame(rows, columns=COLUNAS_SAIDAS)

def metas_default() -> pd.DataFrame:
    """Metas progressivas para 2026."""
    rows = []
    base_fat = 9000.0
    for m in range(2, 13):
        mes = pd.Timestamp(2026, m, 1)
        fator = 1 + (m - 2) * 0.03
        rows.append({
            "Mes": mes,
            "Meta Faturamento": round(base_fat * fator, -1),
            "Meta Despesas %": max(35.0, 42.0 - (m - 2) * 0.5),
            "Meta Valor Guardado": 500.0 + (m - 2) * 50,
            "Meta Saldo Livre": 4000.0 + (m - 2) * 100,
            "Meta Servicos": 450.0 + (m - 2) * 10,
            "Meta Produtos Vendidos": 30.0 + (m - 2) * 2,
        })
    return pd.DataFrame(rows, columns=COLUNAS_METAS)

def historico_default_appbarber() -> pd.DataFrame:
    """Histórico do AppBarber (Fev-Abr/2026) com base no relatório."""
    rows = [
        {
            "Mes": pd.Timestamp(2026, 2, 1),
            "Faturamento": 8300.0,
            "Total Despesas": 3858.28,
            "Despesas Barbearia": 3858.28,
            "Despesas Pessoais": 1080.0,
            "Valor Guardado": 500.0,
            "Saldo Livre": 8300.0 - 3858.28 - 1080.0 - 500.0,
            "Meta Faturamento": 9000.0,
            "Meta Despesas %": 42.0,
            "Atingiu Meta": False,
            "Produtos Valor": 0.0,
            "Produtos Vendidos": 0.0,
            "Servicos Realizados": 0.0,
            "Status": "Atenção",
        },
        {
            "Mes": pd.Timestamp(2026, 3, 1),
            "Faturamento": 9060.50,
            "Total Despesas": 3858.28,
            "Despesas Barbearia": 3858.28,
            "Despesas Pessoais": 1080.0,
            "Valor Guardado": 500.0,
            "Saldo Livre": 9060.50 - 3858.28 - 1080.0 - 500.0,
            "Meta Faturamento": 9000.0,
            "Meta Despesas %": 42.0,
            "Atingiu Meta": True,
            "Produtos Valor": 0.0,
            "Produtos Vendidos": 0.0,
            "Servicos Realizados": 0.0,
            "Status": "OK",
        },
        {
            "Mes": pd.Timestamp(2026, 4, 1),
            "Faturamento": 9060.50,
            "Total Despesas": 3858.28,
            "Despesas Barbearia": 3858.28,
            "Despesas Pessoais": 1080.0,
            "Valor Guardado": 500.0,
            "Saldo Livre": 9060.50 - 3858.28 - 1080.0 - 500.0,
            "Meta Faturamento": 9000.0,
            "Meta Despesas %": 42.0,
            "Atingiu Meta": True,
            "Produtos Valor": 0.0,
            "Produtos Vendidos": 0.0,
            "Servicos Realizados": 0.0,
            "Status": "OK",
        },
    ]
    return pd.DataFrame(rows, columns=COLUNAS_HISTORICO)

# ──────────────────────────────────────────────────────────────
# SESSION STATE
# ──────────────────────────────────────────────────────────────

def init_state():
    if "saidas" not in st.session_state:
        st.session_state.saidas = saidas_default()
    if "historico" not in st.session_state:
        st.session_state.historico = historico_default_appbarber()
    if "metas" not in st.session_state:
        st.session_state.metas = metas_default()
    if "servicos" not in st.session_state:
        st.session_state.servicos = pd.DataFrame(columns=["Serviço", "Realizados"])
    if "produtos" not in st.session_state:
        st.session_state.produtos = pd.DataFrame(columns=["Produto", "Vendidos"])
    if "faturamento_mes" not in st.session_state:
        st.session_state.faturamento_mes = {}
    if "valor_guardado_mes" not in st.session_state:
        st.session_state.valor_guardado_mes = {}
    if "mes_selecionado" not in st.session_state:
        st.session_state.mes_selecionado = label_mes(mes_atual_ts())

# ──────────────────────────────────────────────────────────────
# CÁLCULOS
# ──────────────────────────────────────────────────────────────

def saidas_do_mes(mes_ts: pd.Timestamp) -> pd.DataFrame:
    df = st.session_state.saidas.copy()
    if df.empty:
        return df
    df["_mes"] = df["Mes"].apply(parse_mes)
    return df[df["_mes"] == mes_ts].drop(columns=["_mes"])

def metas_do_mes(mes_ts: pd.Timestamp) -> dict:
    df = st.session_state.metas.copy()
    df["_mes"] = df["Mes"].apply(parse_mes)
    m = df[df["_mes"] == mes_ts]
    if m.empty:
        m = df.sort_values("_mes")
        row = m.iloc[-1] if not m.empty else None
    else:
        row = m.iloc[0]
    if row is None:
        return {"Meta Faturamento": 9000, "Meta Despesas %": 42, "Meta Valor Guardado": 500,
                "Meta Saldo Livre": 4000, "Meta Servicos": 450, "Meta Produtos Vendidos": 30}
    return {c: float(row.get(c, 0)) for c in COLUNAS_METAS if c != "Mes"}

def calcular_resumo_mes(mes_ts: pd.Timestamp) -> dict:
    label = label_mes(mes_ts)
    saidas = saidas_do_mes(mes_ts)
    faturamento = st.session_state.faturamento_mes.get(label, 0.0)
    valor_guardado = st.session_state.valor_guardado_mes.get(label, 0.0)

    desp_barb = saidas[saidas["Centro"] == "Barbearia"]["Valor"].sum() if not saidas.empty else 0.0
    desp_pess = saidas[saidas["Centro"] == "Pessoal"]["Valor"].sum() if not saidas.empty else 0.0
    total_desp = desp_barb + desp_pess

    saldo_livre = faturamento - total_desp - valor_guardado
    pct_desp = (desp_barb / faturamento * 100) if faturamento > 0 else 0.0

    metas = metas_do_mes(mes_ts)
    atingiu = (
        faturamento >= metas["Meta Faturamento"] and
        pct_desp <= metas["Meta Despesas %"]
    )
    if faturamento == 0:
        status = "Sem dados"
    elif pct_desp > metas["Meta Despesas %"] * 1.15 or faturamento < metas["Meta Faturamento"] * 0.85:
        status = "Crítico"
    elif pct_desp > metas["Meta Despesas %"] or faturamento < metas["Meta Faturamento"]:
        status = "Atenção"
    else:
        status = "OK"

    return {
        "Faturamento": faturamento,
        "Despesas Barbearia": desp_barb,
        "Despesas Pessoais": desp_pess,
        "Total Despesas": total_desp,
        "Valor Guardado": valor_guardado,
        "Saldo Livre": saldo_livre,
        "% Despesas Barb": pct_desp,
        "Metas": metas,
        "Atingiu Meta": atingiu,
        "Status": status,
    }

def atualizar_historico(mes_ts: pd.Timestamp):
    resumo = calcular_resumo_mes(mes_ts)
    label = label_mes(mes_ts)
    hist = st.session_state.historico.copy()
    hist["_mes"] = hist["Mes"].apply(parse_mes)
    hist = hist[hist["_mes"] != mes_ts].drop(columns=["_mes"])

    nova_linha = {
        "Mes": mes_ts,
        "Faturamento": resumo["Faturamento"],
        "Total Despesas": resumo["Total Despesas"],
        "Despesas Barbearia": resumo["Despesas Barbearia"],
        "Despesas Pessoais": resumo["Despesas Pessoais"],
        "Valor Guardado": resumo["Valor Guardado"],
        "Saldo Livre": resumo["Saldo Livre"],
        "Meta Faturamento": resumo["Metas"]["Meta Faturamento"],
        "Meta Despesas %": resumo["Metas"]["Meta Despesas %"],
        "Atingiu Meta": resumo["Atingiu Meta"],
        "Produtos Valor": 0.0,
        "Produtos Vendidos": float(st.session_state.produtos["Vendidos"].sum()) if not st.session_state.produtos.empty and "Vendidos" in st.session_state.produtos.columns else 0.0,
        "Servicos Realizados": float(st.session_state.servicos["Realizados"].sum()) if not st.session_state.servicos.empty and "Realizados" in st.session_state.servicos.columns else 0.0,
        "Status": resumo["Status"],
    }
    hist = pd.concat([hist, pd.DataFrame([nova_linha])], ignore_index=True)
    hist["_mes"] = hist["Mes"].apply(parse_mes)
    hist = hist.sort_values("_mes").drop(columns=["_mes"])
    st.session_state.historico = hist

# ──────────────────────────────────────────────────────────────
# EXPORTAÇÃO BASE (xlsx)
# ──────────────────────────────────────────────────────────────

def exportar_base() -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        st.session_state.historico.to_excel(w, index=False, sheet_name="Historico")
        st.session_state.saidas.to_excel(w, index=False, sheet_name="Saidas")
        st.session_state.metas.to_excel(w, index=False, sheet_name="Metas")
        pd.DataFrame(
            list(st.session_state.faturamento_mes.items()), columns=["Mes", "Faturamento"]
        ).to_excel(w, index=False, sheet_name="Faturamento")
        pd.DataFrame(
            list(st.session_state.valor_guardado_mes.items()), columns=["Mes", "ValorGuardado"]
        ).to_excel(w, index=False, sheet_name="ValorGuardado")
        st.session_state.servicos.to_excel(w, index=False, sheet_name="Ranking_Servicos")
        st.session_state.produtos.to_excel(w, index=False, sheet_name="Ranking_Produtos")
    return buf.getvalue()

def importar_base(file):
    try:
        file.seek(0)
        xl = pd.ExcelFile(file)
        sheets = xl.sheet_names

        if "Historico" in sheets:
            hist = xl.parse("Historico")
            if not hist.empty:
                hist["Mes"] = hist["Mes"].apply(parse_mes)
                st.session_state.historico = hist

        if "Saidas" in sheets:
            saidas = xl.parse("Saidas")
            if not saidas.empty:
                saidas["Mes"] = saidas["Mes"].apply(parse_mes)
                st.session_state.saidas = saidas

        if "Metas" in sheets:
            metas = xl.parse("Metas")
            if not metas.empty:
                metas["Mes"] = metas["Mes"].apply(parse_mes)
                st.session_state.metas = metas

        if "Faturamento" in sheets:
            fat = xl.parse("Faturamento")
            for _, row in fat.iterrows():
                st.session_state.faturamento_mes[str(row["Mes"])] = float(row.get("Faturamento", 0))

        if "ValorGuardado" in sheets:
            vg = xl.parse("ValorGuardado")
            for _, row in vg.iterrows():
                st.session_state.valor_guardado_mes[str(row["Mes"])] = float(row.get("ValorGuardado", 0))

        if "Ranking_Servicos" in sheets:
            st.session_state.servicos = xl.parse("Ranking_Servicos")
        if "Ranking_Produtos" in sheets:
            st.session_state.produtos = xl.parse("Ranking_Produtos")

        return True
    except Exception as e:
        st.error(f"Erro ao importar base: {e}")
        return False

# ──────────────────────────────────────────────────────────────
# GERAÇÃO DE PDF
# ──────────────────────────────────────────────────────────────

class PDFRelatorio(FPDF):
    def header(self):
        self.set_fill_color(30, 30, 30)
        self.rect(0, 0, 210, 28, "F")
        self.set_text_color(255, 255, 255)
        self.set_font("Helvetica", "B", 16)
        self.set_xy(10, 8)
        self.cell(0, 8, "Barbearia Heloisa Mazzi", ln=False)
        self.set_font("Helvetica", "", 10)
        self.set_xy(10, 18)
        self.cell(0, 6, "Relatório de Fechamento Mensal · Documento Confidencial", ln=True)
        self.set_text_color(0, 0, 0)
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Página {self.page_no()} · Gerado automaticamente pelo Painel de Gestão", align="C")

    def kpi_box(self, x, y, w, h, titulo, valor, cor=(40, 167, 69)):
        self.set_fill_color(*cor)
        self.rect(x, y, w, h, "F")
        self.set_text_color(255, 255, 255)
        self.set_font("Helvetica", "B", 7)
        self.set_xy(x + 2, y + 2)
        self.cell(w - 4, 4, titulo.upper(), ln=True)
        self.set_font("Helvetica", "B", 11)
        self.set_xy(x + 2, y + 7)
        self.cell(w - 4, 6, valor, ln=True)
        self.set_text_color(0, 0, 0)

    def section_title(self, title):
        self.set_font("Helvetica", "B", 12)
        self.set_fill_color(245, 245, 245)
        self.set_text_color(30, 30, 30)
        self.cell(0, 8, title, ln=True, fill=True)
        self.set_draw_color(200, 200, 200)
        self.ln(1)

    def tabela_simples(self, headers, rows, col_widths=None):
        if col_widths is None:
            w = (self.w - 20) / len(headers)
            col_widths = [w] * len(headers)
        self.set_font("Helvetica", "B", 8)
        self.set_fill_color(50, 50, 50)
        self.set_text_color(255, 255, 255)
        for h, cw in zip(headers, col_widths):
            self.cell(cw, 6, str(h), border=0, fill=True)
        self.ln()
        self.set_font("Helvetica", "", 8)
        fill = False
        for i, row in enumerate(rows):
            self.set_fill_color(248, 248, 248) if fill else self.set_fill_color(255, 255, 255)
            self.set_text_color(0, 0, 0)
            for val, cw in zip(row, col_widths):
                self.cell(cw, 5, str(val or ""), border=0, fill=True)
            self.ln()
            fill = not fill
        self.ln(3)


def gerar_pdf(mes_ts: pd.Timestamp) -> bytes:
    resumo = calcular_resumo_mes(mes_ts)
    saidas = saidas_do_mes(mes_ts)
    metas = resumo["Metas"]
    status = resumo["Status"]
    cor_status_map = {"OK": (40, 167, 69), "Atenção": (255, 193, 7), "Crítico": (220, 53, 69), "Sem dados": (108, 117, 125)}

    pdf = PDFRelatorio()
    pdf.add_page()

    # Título do mês
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(0, 10, f"Fechamento — {mes_extenso(mes_ts)}", ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, f"Status: {status}  |  Meta de faturamento: {formatar_brl(metas['Meta Faturamento'])}", ln=True)
    pdf.ln(4)

    # KPIs em grid
    cor = cor_status_map.get(status, (108, 117, 125))
    kpis = [
        ("Faturamento", formatar_brl(resumo["Faturamento"]), (40, 167, 69)),
        ("Despesas Barbearia", formatar_brl(resumo["Despesas Barbearia"]), (220, 53, 69)),
        ("Despesas Pessoais", formatar_brl(resumo["Despesas Pessoais"]), (255, 100, 50)),
        ("Valor Guardado", formatar_brl(resumo["Valor Guardado"]), (13, 110, 253)),
        ("Saldo Livre", formatar_brl(resumo["Saldo Livre"]), cor),
        ("% Despesas s/ Fat.", formatar_pct(resumo["% Despesas Barb"]), cor),
    ]
    x0 = 10
    box_w, box_h = 61, 18
    for i, (titulo, valor, cor_box) in enumerate(kpis):
        col = i % 3
        row = i // 3
        pdf.kpi_box(x0 + col * (box_w + 3), pdf.get_y() if col == 0 else pdf.get_y() - box_h - 3 + row * (box_h + 3),
                    box_w, box_h, titulo, valor, cor_box)
        if col == 2:
            pdf.ln(box_h + 3)
    pdf.ln(5)

    # Metas
    pdf.section_title("2. Comparativo com Metas")
    meta_rows = [
        ["Faturamento", formatar_brl(resumo["Faturamento"]), formatar_brl(metas["Meta Faturamento"]),
         "✓" if resumo["Faturamento"] >= metas["Meta Faturamento"] else "✗"],
        ["Despesas s/Fat (%)", formatar_pct(resumo["% Despesas Barb"]), formatar_pct(metas["Meta Despesas %"]),
         "✓" if resumo["% Despesas Barb"] <= metas["Meta Despesas %"] else "✗"],
        ["Valor Guardado", formatar_brl(resumo["Valor Guardado"]), formatar_brl(metas["Meta Valor Guardado"]),
         "✓" if resumo["Valor Guardado"] >= metas["Meta Valor Guardado"] else "✗"],
        ["Saldo Livre", formatar_brl(resumo["Saldo Livre"]), formatar_brl(metas["Meta Saldo Livre"]),
         "✓" if resumo["Saldo Livre"] >= metas["Meta Saldo Livre"] else "✗"],
    ]
    pdf.tabela_simples(["Indicador", "Realizado", "Meta", "Status"], meta_rows, [70, 40, 40, 20])

    # Saídas por categoria
    if not saidas.empty:
        pdf.section_title("3. Saídas por Categoria")
        cat_group = saidas.groupby(["Centro", "Categoria"])["Valor"].sum().reset_index()
        cat_rows = []
        for _, r in cat_group.sort_values("Valor", ascending=False).iterrows():
            cat_rows.append([r["Centro"], r["Categoria"], formatar_brl(r["Valor"])])
        pdf.tabela_simples(["Centro", "Categoria", "Valor"], cat_rows, [50, 90, 40])

        pdf.section_title("4. Detalhamento de Saídas")
        det_rows = []
        for _, r in saidas.sort_values(["Centro", "Categoria"]).iterrows():
            det_rows.append([
                r.get("Centro", ""),
                r.get("Descricao", "")[:30],
                r.get("Tipo", ""),
                formatar_brl(float(r.get("Valor", 0))),
                r.get("Status", ""),
            ])
        pdf.tabela_simples(["Centro", "Descrição", "Tipo", "Valor", "Status"], det_rows, [35, 65, 25, 35, 20])

    # Observações
    pdf.section_title("5. Observações e Próximos Passos")
    obs_lines = []
    if resumo["Faturamento"] >= metas["Meta Faturamento"]:
        obs_lines.append("✓ Meta de faturamento atingida neste mês.")
    else:
        diff = metas["Meta Faturamento"] - resumo["Faturamento"]
        obs_lines.append(f"✗ Faturamento abaixo da meta em {formatar_brl(diff)}.")
    if resumo["% Despesas Barb"] > metas["Meta Despesas %"]:
        obs_lines.append(f"⚠ Despesas comprometendo {formatar_pct(resumo['% Despesas Barb'])} do faturamento (meta: {formatar_pct(metas['Meta Despesas %'])}).")
    if resumo["Valor Guardado"] > 0:
        obs_lines.append(f"✓ {formatar_brl(resumo['Valor Guardado'])} guardados este mês.")
    if resumo["Saldo Livre"] > 0:
        obs_lines.append(f"✓ Saldo livre de {formatar_brl(resumo['Saldo Livre'])} disponível.")

    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(30, 30, 30)
    for line in obs_lines:
        pdf.cell(0, 6, line, ln=True)

    pdf.ln(5)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(130, 130, 130)
    pdf.cell(0, 5, f"Relatório gerado em {datetime.date.today().strftime('%d/%m/%Y')} pelo Painel de Gestão — Barbearia Heloisa Mazzi", ln=True)

    return bytes(pdf.output())

# ──────────────────────────────────────────────────────────────
# INTERFACE
# ──────────────────────────────────────────────────────────────

def render_gauge(valor, meta, titulo, fmt="%"):
    if fmt == "R$":
        texto_val = formatar_brl(valor)
        texto_meta = formatar_brl(meta)
        pct = (valor / meta * 100) if meta > 0 else 0
    else:
        texto_val = formatar_pct(valor)
        texto_meta = formatar_pct(meta)
        # Para % despesas, menor é melhor
        pct = max(0, 100 - (valor / meta * 100)) if meta > 0 else 100

    cor = "#28a745" if pct >= 80 else "#ffc107" if pct >= 50 else "#dc3545"
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=valor,
        number={"prefix": "R$ " if fmt == "R$" else "", "suffix": "%" if fmt == "%" else "", "valueformat": ",.2f" if fmt == "R$" else ".1f"},
        delta={"reference": meta, "relative": False,
               "increasing": {"color": "#28a745"} if fmt == "R$" else {"color": "#dc3545"},
               "decreasing": {"color": "#dc3545"} if fmt == "R$" else {"color": "#28a745"}},
        title={"text": titulo, "font": {"size": 12}},
        gauge={
            "axis": {"range": [0, meta * 1.5 if fmt == "R$" else meta * 1.5]},
            "bar": {"color": cor},
            "steps": [{"range": [0, meta * 0.5], "color": "#ffebeb"},
                      {"range": [meta * 0.5, meta], "color": "#fff3cd"},
                      {"range": [meta, meta * 1.5], "color": "#d4edda"}],
            "threshold": {"line": {"color": "black", "width": 2}, "thickness": 0.75, "value": meta},
        }
    ))
    fig.update_layout(height=200, margin=dict(l=10, r=10, t=40, b=10))
    return fig

def aba_dashboard(mes_ts: pd.Timestamp):
    st.header(f"📊 Dashboard — {mes_extenso(mes_ts)}")

    # Faturamento e valor guardado
    col1, col2 = st.columns(2)
    label = label_mes(mes_ts)
    with col1:
        fat = st.number_input(
            "💰 Faturamento do mês (R$)",
            value=float(st.session_state.faturamento_mes.get(label, 0.0)),
            min_value=0.0, step=100.0, format="%.2f",
            help="Digite o faturamento líquido total do AppBarber para este mês",
            key=f"fat_{label}",
        )
        st.session_state.faturamento_mes[label] = fat
    with col2:
        vg = st.number_input(
            "🏦 Valor Guardado este mês (R$)",
            value=float(st.session_state.valor_guardado_mes.get(label, 0.0)),
            min_value=0.0, step=50.0, format="%.2f",
            help="Quanto foi guardado/poupado neste mês",
            key=f"vg_{label}",
        )
        st.session_state.valor_guardado_mes[label] = vg

    resumo = calcular_resumo_mes(mes_ts)
    metas = resumo["Metas"]
    status = resumo["Status"]
    emoji = cor_status(status)

    # Banner status
    cor_banner = {"OK": "#28a745", "Atenção": "#ffc107", "Crítico": "#dc3545", "Sem dados": "#6c757d"}.get(status, "#6c757d")
    st.markdown(f"""
    <div style="background:{cor_banner};color:white;padding:12px 20px;border-radius:8px;font-size:18px;font-weight:bold;margin-bottom:16px">
    {emoji} Status do mês: {status}
    </div>
    """, unsafe_allow_html=True)

    # KPIs principais
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    kpis = [
        (c1, "💰 Faturamento", resumo["Faturamento"], "success"),
        (c2, "🔴 Desp. Barbearia", resumo["Despesas Barbearia"], "error"),
        (c3, "🟠 Desp. Pessoais", resumo["Despesas Pessoais"], "warning"),
        (c4, "🏦 Guardado", resumo["Valor Guardado"], "info"),
        (c5, "✅ Saldo Livre", resumo["Saldo Livre"], "success" if resumo["Saldo Livre"] >= 0 else "error"),
        (c6, "📊 % Despesas", None, None),
    ]
    for col, titulo, valor, tipo in kpis[:-1]:
        col.metric(titulo, formatar_brl(valor))
    c6.metric("📊 % Desp/Fat", formatar_pct(resumo["% Despesas Barb"]),
              delta=f"Meta: {formatar_pct(metas['Meta Despesas %'])}",
              delta_color="inverse")

    st.divider()

    # Gráficos de meta
    st.subheader("🎯 Metas do Mês")
    g1, g2, g3 = st.columns(3)
    with g1:
        st.plotly_chart(render_gauge(resumo["Faturamento"], metas["Meta Faturamento"], "Faturamento", "R$"),
                        use_container_width=True, key="gauge_fat")
    with g2:
        st.plotly_chart(render_gauge(resumo["% Despesas Barb"], metas["Meta Despesas %"], "% Despesas (menor=melhor)", "%"),
                        use_container_width=True, key="gauge_desp")
    with g3:
        st.plotly_chart(render_gauge(resumo["Valor Guardado"], metas["Meta Valor Guardado"], "Valor Guardado", "R$"),
                        use_container_width=True, key="gauge_guard")

    # Composição das despesas
    if not saidas_do_mes(mes_ts).empty:
        st.subheader("🧾 Composição das Despesas")
        saidas = saidas_do_mes(mes_ts)
        col_pie, col_bar = st.columns(2)
        with col_pie:
            grupo = saidas.groupby("Centro")["Valor"].sum().reset_index()
            fig_pie = px.pie(grupo, values="Valor", names="Centro",
                             color_discrete_sequence=["#dc3545", "#ff7f50"],
                             title="Por Centro")
            fig_pie.update_layout(height=280, margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(fig_pie, use_container_width=True, key="pie_centro")
        with col_bar:
            cat_group = saidas.groupby("Categoria")["Valor"].sum().reset_index().sort_values("Valor", ascending=True)
            fig_bar = px.bar(cat_group, x="Valor", y="Categoria", orientation="h",
                             color="Valor", color_continuous_scale="Reds",
                             title="Por Categoria")
            fig_bar.update_layout(height=280, margin=dict(l=0, r=0, t=40, b=0), coloraxis_showscale=False)
            st.plotly_chart(fig_bar, use_container_width=True, key="bar_cat")

    # Botão fechar mês
    st.divider()
    col_btn1, col_btn2 = st.columns([1, 3])
    with col_btn1:
        if st.button("🔒 Fechar Mês e Salvar no Histórico", type="primary", use_container_width=True):
            atualizar_historico(mes_ts)
            st.success(f"✅ Mês {mes_extenso(mes_ts)} salvo no histórico!")
            st.rerun()
    with col_btn2:
        pdf_bytes = gerar_pdf(mes_ts)
        st.download_button(
            "📄 Baixar Relatório PDF do Mês",
            data=pdf_bytes,
            file_name=f"relatorio_barbearia_{label_mes(mes_ts).replace('/', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )


def aba_saidas(mes_ts: pd.Timestamp):
    st.header(f"💸 Registro de Saídas — {mes_extenso(mes_ts)}")

    # Importar do AppBarber
    with st.expander("📥 Importar Relatórios do AppBarber", expanded=False):
        st.info("Faça upload dos arquivos exportados diretamente do AppBarber.")
        col1, col2, col3 = st.columns(3)
        with col1:
            f_rec = st.file_uploader("Total Recebimento", type=["xlsx"], key="up_rec")
        with col2:
            f_serv = st.file_uploader("Ranking Serviços", type=["xlsx"], key="up_serv")
        with col3:
            f_prod = st.file_uploader("Ranking Produtos", type=["xlsx"], key="up_prod")

        if st.button("📊 Processar Uploads"):
            if f_rec:
                df_rec = ler_excel_upload(f_rec)
                result = processar_recebimentos(df_rec)
                for _, row in result.iterrows():
                    lbl = label_mes(row["Mes"])
                    st.session_state.faturamento_mes[lbl] = row["Faturamento"]
                st.success(f"✅ Recebimentos importados: {len(result)} meses")
            if f_serv:
                df_s = ler_excel_upload(f_serv)
                st.session_state.servicos = limpar_ranking(df_s, "Serviço", "Realizados")
                st.success("✅ Ranking de serviços importado")
            if f_prod:
                df_p = ler_excel_upload(f_prod)
                st.session_state.produtos = limpar_ranking(df_p, "Produto", "Vendidos")
                st.success("✅ Ranking de produtos importado")
            st.rerun()

    # Formulário de nova saída
    with st.expander("➕ Registrar Nova Saída", expanded=True):
        st.subheader("Nova Saída de Caixa")
        c1, c2 = st.columns(2)
        with c1:
            centro = st.selectbox("Centro", ["Barbearia", "Pessoal"], key="ns_centro")
            categorias = CATEGORIAS_BARBEARIA if centro == "Barbearia" else CATEGORIAS_PESSOAL
            categoria = st.selectbox("Categoria", categorias, key="ns_cat")
            tipo = st.selectbox("Tipo", ["Fixo", "Variável", "Parcela", "Pontual/Investimento"], key="ns_tipo")
            descricao = st.text_input("Descrição", key="ns_desc")
        with c2:
            valor = st.number_input("Valor (R$)", min_value=0.0, step=10.0, format="%.2f", key="ns_valor")
            data = st.date_input("Data", value=datetime.date.today(), key="ns_data")
            forma_pag = st.selectbox("Forma de Pagamento", FORMAS_PAGTO, key="ns_forma")
            status_saida = st.selectbox("Status", ["Pago", "Previsto", "Cancelado"], key="ns_status")

        c3, c4, c5 = st.columns(3)
        with c3:
            parc_atual = st.text_input("Parcela Atual (ex: 1)", key="ns_parc_at", value="")
        with c4:
            parc_total = st.text_input("Total Parcelas (ex: 10)", key="ns_parc_tot", value="")
        with c5:
            obs = st.text_input("Observações", key="ns_obs", value="")

        if st.button("💾 Adicionar Saída", type="primary"):
            if not descricao:
                st.error("Descrição é obrigatória.")
            elif valor <= 0:
                st.error("Valor deve ser maior que zero.")
            else:
                nova = {
                    "Mes": mes_ts,
                    "Data": data.strftime("%d/%m/%Y"),
                    "Centro": centro,
                    "Categoria": categoria,
                    "Tipo": tipo,
                    "Descricao": descricao,
                    "Valor": valor,
                    "Forma Pagamento": forma_pag,
                    "Parcela Atual": parc_atual,
                    "Total Parcelas": parc_total,
                    "Status": status_saida,
                    "Observacoes": obs,
                }
                st.session_state.saidas = pd.concat(
                    [st.session_state.saidas, pd.DataFrame([nova])], ignore_index=True
                )
                st.success(f"✅ {descricao} — {formatar_brl(valor)} adicionado!")
                st.rerun()

    # Tabela de saídas do mês
    st.subheader(f"📋 Saídas de {mes_extenso(mes_ts)}")
    saidas_mes = saidas_do_mes(mes_ts)

    if saidas_mes.empty:
        st.info("Nenhuma saída registrada para este mês ainda.")
    else:
        # Totais
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Barbearia", formatar_brl(saidas_mes[saidas_mes["Centro"] == "Barbearia"]["Valor"].sum()))
        c2.metric("Total Pessoal", formatar_brl(saidas_mes[saidas_mes["Centro"] == "Pessoal"]["Valor"].sum()))
        c3.metric("Total Geral", formatar_brl(saidas_mes["Valor"].sum()))

        # Filtros
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filtro_centro = st.multiselect("Filtrar Centro", ["Barbearia", "Pessoal"],
                                           default=["Barbearia", "Pessoal"], key="fil_centro")
        with col_f2:
            filtro_tipo = st.multiselect("Filtrar Tipo", ["Fixo", "Variável", "Parcela", "Pontual/Investimento"],
                                         default=["Fixo", "Variável", "Parcela", "Pontual/Investimento"], key="fil_tipo")

        df_show = saidas_mes[saidas_mes["Centro"].isin(filtro_centro) & saidas_mes["Tipo"].isin(filtro_tipo)].copy()
        df_show["Valor_fmt"] = df_show["Valor"].apply(formatar_brl)

        st.dataframe(
            df_show[["Data", "Centro", "Categoria", "Tipo", "Descricao", "Valor_fmt", "Forma Pagamento", "Parcela Atual", "Total Parcelas", "Status"]].rename(columns={"Valor_fmt": "Valor"}),
            use_container_width=True, hide_index=True
        )

        # Excluir saída
        with st.expander("🗑 Excluir uma saída"):
            idx_options = df_show.index.tolist()
            if idx_options:
                desc_options = [f"{r['Descricao']} — {formatar_brl(r['Valor'])} ({r.get('Data', '')})"
                                for _, r in df_show.iterrows()]
                selected = st.selectbox("Selecione para excluir", options=range(len(desc_options)),
                                        format_func=lambda i: desc_options[i], key="excluir_idx")
                if st.button("🗑 Confirmar Exclusão", type="secondary"):
                    real_idx = idx_options[selected]
                    st.session_state.saidas = st.session_state.saidas.drop(index=real_idx).reset_index(drop=True)
                    st.success("Saída excluída.")
                    st.rerun()


def aba_historico():
    st.header("📈 Histórico e Evolução")
    hist = st.session_state.historico.copy()
    if hist.empty:
        st.info("Nenhum dado histórico ainda. Feche um mês no Dashboard para começar.")
        return

    hist["_mes"] = hist["Mes"].apply(parse_mes)
    hist = hist.dropna(subset=["_mes"]).sort_values("_mes")
    hist["Mes Label"] = hist["_mes"].apply(mes_extenso)

    # KPIs históricos
    fat_medio = hist["Faturamento"].mean()
    fat_max = hist["Faturamento"].max()
    meses_ok = (hist["Status"] == "OK").sum()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Meses no Histórico", len(hist))
    c2.metric("Faturamento Médio", formatar_brl(fat_medio))
    c3.metric("Melhor Mês", formatar_brl(fat_max))
    c4.metric("Meses com Meta Atingida", f"{meses_ok}/{len(hist)}")

    st.divider()

    # Gráfico evolução faturamento vs meta
    st.subheader("💰 Evolução do Faturamento vs Meta")
    fig_fat = go.Figure()
    fig_fat.add_trace(go.Bar(x=hist["Mes Label"], y=hist["Faturamento"],
                              name="Faturamento", marker_color="#28a745"))
    fig_fat.add_trace(go.Scatter(x=hist["Mes Label"], y=hist["Meta Faturamento"],
                                  name="Meta", line=dict(color="#ffc107", width=2, dash="dash"),
                                  mode="lines+markers"))
    fig_fat.update_layout(height=320, margin=dict(l=0, r=0, t=10, b=0),
                           yaxis_tickprefix="R$ ", legend=dict(orientation="h"))
    st.plotly_chart(fig_fat, use_container_width=True, key="hist_fat")

    # Gráfico despesas vs faturamento
    st.subheader("📊 Despesas vs Faturamento")
    fig_desp = go.Figure()
    fig_desp.add_trace(go.Bar(x=hist["Mes Label"], y=hist["Faturamento"],
                               name="Faturamento", marker_color="#28a745", opacity=0.7))
    fig_desp.add_trace(go.Bar(x=hist["Mes Label"], y=hist["Total Despesas"],
                               name="Total Despesas", marker_color="#dc3545"))
    fig_desp.add_trace(go.Bar(x=hist["Mes Label"], y=hist["Valor Guardado"],
                               name="Guardado", marker_color="#0d6efd"))
    fig_desp.update_layout(barmode="overlay", height=320, margin=dict(l=0, r=0, t=10, b=0),
                            yaxis_tickprefix="R$ ", legend=dict(orientation="h"))
    st.plotly_chart(fig_desp, use_container_width=True, key="hist_desp")

    # % Despesas por mês
    st.subheader("📉 % Despesas sobre Faturamento")
    hist_pct = hist[hist["Faturamento"] > 0].copy()
    hist_pct["% Real"] = hist_pct["Total Despesas"] / hist_pct["Faturamento"] * 100
    fig_pct = go.Figure()
    fig_pct.add_trace(go.Bar(x=hist_pct["Mes Label"], y=hist_pct["% Real"],
                              marker_color=[
                                  "#28a745" if v <= m else "#ffc107" if v <= m * 1.1 else "#dc3545"
                                  for v, m in zip(hist_pct["% Real"], hist_pct["Meta Despesas %"])
                              ],
                              name="% Real"))
    fig_pct.add_trace(go.Scatter(x=hist_pct["Mes Label"], y=hist_pct["Meta Despesas %"],
                                  name="Meta %", line=dict(color="#ffc107", dash="dash"),
                                  mode="lines+markers"))
    fig_pct.update_layout(height=280, margin=dict(l=0, r=0, t=10, b=0),
                           yaxis_ticksuffix="%", legend=dict(orientation="h"))
    st.plotly_chart(fig_pct, use_container_width=True, key="hist_pct")

    # Saldo livre e guardado
    st.subheader("🏦 Saldo Livre e Valor Guardado")
    fig_saldo = go.Figure()
    fig_saldo.add_trace(go.Bar(x=hist["Mes Label"], y=hist["Saldo Livre"],
                                name="Saldo Livre", marker_color="#0dcaf0"))
    fig_saldo.add_trace(go.Bar(x=hist["Mes Label"], y=hist["Valor Guardado"],
                                name="Guardado", marker_color="#0d6efd"))
    fig_saldo.update_layout(barmode="group", height=280, margin=dict(l=0, r=0, t=10, b=0),
                             yaxis_tickprefix="R$ ", legend=dict(orientation="h"))
    st.plotly_chart(fig_saldo, use_container_width=True, key="hist_saldo")

    # Tabela histórico
    st.subheader("📋 Tabela Histórico Completo")
    hist_show = hist.copy()
    for col in ["Faturamento", "Total Despesas", "Despesas Barbearia", "Despesas Pessoais", "Valor Guardado", "Saldo Livre"]:
        if col in hist_show.columns:
            hist_show[col] = hist_show[col].apply(formatar_brl)
    if "% Despesas" in hist_show.columns:
        hist_show["% Despesas"] = hist_show["% Despesas"].apply(lambda x: formatar_pct(float(x) * 100 if float(x) < 1 else float(x)))
    cols_show = ["Mes Label", "Faturamento", "Total Despesas", "Despesas Barbearia", "Despesas Pessoais", "Valor Guardado", "Saldo Livre", "Status"]
    cols_show = [c for c in cols_show if c in hist_show.columns]
    st.dataframe(hist_show[cols_show].rename(columns={"Mes Label": "Mês"}),
                 use_container_width=True, hide_index=True)


def aba_metas():
    st.header("🎯 Metas Mensais")
    st.info("As metas evoluem automaticamente mês a mês. Aqui você pode ajustar individualmente.")

    metas = st.session_state.metas.copy()
    metas["_mes"] = metas["Mes"].apply(parse_mes)
    metas = metas.dropna(subset=["_mes"]).sort_values("_mes")
    metas["Mes Label"] = metas["_mes"].apply(mes_extenso)

    # Gráfico metas faturamento
    fig = px.line(metas, x="Mes Label", y="Meta Faturamento",
                  title="Evolução da Meta de Faturamento", markers=True,
                  color_discrete_sequence=["#28a745"])
    fig.update_layout(height=250, margin=dict(l=0, r=0, t=40, b=0), yaxis_tickprefix="R$ ")
    st.plotly_chart(fig, use_container_width=True, key="meta_fat_chart")

    # Editar metas
    st.subheader("✏️ Editar Metas")
    metas_edit = metas[["Mes Label", "Meta Faturamento", "Meta Despesas %", "Meta Valor Guardado",
                          "Meta Saldo Livre", "Meta Servicos", "Meta Produtos Vendidos"]].copy()
    edited = st.data_editor(
        metas_edit.rename(columns={"Mes Label": "Mês"}),
        use_container_width=True, hide_index=True,
        column_config={
            "Meta Faturamento": st.column_config.NumberColumn("Meta Faturamento (R$)", format="R$ %.2f"),
            "Meta Despesas %": st.column_config.NumberColumn("Meta Despesas (%)", format="%.1f%%"),
            "Meta Valor Guardado": st.column_config.NumberColumn("Meta Guardado (R$)", format="R$ %.2f"),
            "Meta Saldo Livre": st.column_config.NumberColumn("Meta Saldo Livre (R$)", format="R$ %.2f"),
            "Meta Servicos": st.column_config.NumberColumn("Meta Serviços", format="%.0f"),
            "Meta Produtos Vendidos": st.column_config.NumberColumn("Meta Produtos", format="%.0f"),
        },
        key="metas_editor"
    )

    if st.button("💾 Salvar Metas", type="primary"):
        updated = metas.copy()
        for col in ["Meta Faturamento", "Meta Despesas %", "Meta Valor Guardado",
                    "Meta Saldo Livre", "Meta Servicos", "Meta Produtos Vendidos"]:
            if col in edited.columns:
                updated[col] = edited[col].values
        st.session_state.metas = updated.drop(columns=["_mes", "Mes Label"])
        st.success("✅ Metas salvas!")
        st.rerun()

    # Adicionar meta para novo mês
    st.subheader("➕ Adicionar Meta para Novo Mês")
    col1, col2 = st.columns(2)
    with col1:
        novo_mes_meta = st.date_input("Mês", value=datetime.date.today().replace(day=1), key="novo_meta_mes")
        meta_fat = st.number_input("Meta Faturamento", value=9000.0, step=500.0, format="%.2f", key="nm_fat")
        meta_desp = st.number_input("Meta Despesas %", value=40.0, step=1.0, format="%.1f", key="nm_desp")
    with col2:
        meta_guard = st.number_input("Meta Guardado", value=500.0, step=100.0, format="%.2f", key="nm_guard")
        meta_saldo = st.number_input("Meta Saldo Livre", value=4000.0, step=500.0, format="%.2f", key="nm_saldo")
        meta_serv = st.number_input("Meta Serviços", value=450.0, step=10.0, format="%.0f", key="nm_serv")
        meta_prod = st.number_input("Meta Produtos Vendidos", value=30.0, step=5.0, format="%.0f", key="nm_prod")

    if st.button("➕ Adicionar Meta"):
        mes_ts = pd.Timestamp(novo_mes_meta.year, novo_mes_meta.month, 1)
        nova = {
            "Mes": mes_ts,
            "Meta Faturamento": meta_fat,
            "Meta Despesas %": meta_desp,
            "Meta Valor Guardado": meta_guard,
            "Meta Saldo Livre": meta_saldo,
            "Meta Servicos": meta_serv,
            "Meta Produtos Vendidos": meta_prod,
        }
        metas_atual = st.session_state.metas.copy()
        metas_atual["_mes"] = metas_atual["Mes"].apply(parse_mes)
        metas_atual = metas_atual[metas_atual["_mes"] != mes_ts].drop(columns=["_mes"])
        st.session_state.metas = pd.concat([metas_atual, pd.DataFrame([nova])], ignore_index=True)
        st.success(f"✅ Meta para {mes_extenso(mes_ts)} adicionada!")
        st.rerun()


def aba_rankings():
    st.header("🏆 Rankings de Serviços e Produtos")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("✂️ Serviços mais Realizados")
        servicos = st.session_state.servicos
        if servicos.empty or "Realizados" not in servicos.columns:
            st.info("Faça upload do ranking de serviços na aba de Saídas.")
        else:
            top = servicos.head(10)
            fig = px.bar(top, x="Realizados", y="Serviço", orientation="h",
                         color="Realizados", color_continuous_scale="Blues",
                         title="Top 10 Serviços")
            fig.update_layout(height=350, margin=dict(l=0, r=0, t=40, b=0), coloraxis_showscale=False)
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(fig, use_container_width=True, key="rank_serv")
            st.dataframe(top, use_container_width=True, hide_index=True)

    with col2:
        st.subheader("💊 Produtos mais Vendidos")
        produtos = st.session_state.produtos
        if produtos.empty or "Vendidos" not in produtos.columns:
            st.info("Faça upload do ranking de produtos na aba de Saídas.")
        else:
            top = produtos.head(10)
            fig = px.bar(top, x="Vendidos", y="Produto", orientation="h",
                         color="Vendidos", color_continuous_scale="Greens",
                         title="Top 10 Produtos")
            fig.update_layout(height=350, margin=dict(l=0, r=0, t=40, b=0), coloraxis_showscale=False)
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(fig, use_container_width=True, key="rank_prod")
            st.dataframe(top, use_container_width=True, hide_index=True)


def aba_relatorios():
    st.header("📄 Relatórios PDF")
    st.write("Gere e baixe o relatório de fechamento de qualquer mês do histórico.")

    hist = st.session_state.historico.copy()
    hist["_mes"] = hist["Mes"].apply(parse_mes)
    hist = hist.dropna(subset=["_mes"]).sort_values("_mes")

    # Todos os meses com dados
    meses_disponiveis = [label_mes(m) for m in hist["_mes"].tolist()]
    # Incluir mês atual mesmo sem histórico fechado
    mes_at = label_mes(mes_atual_ts())
    if mes_at not in meses_disponiveis:
        meses_disponiveis.append(mes_at)

    if not meses_disponiveis:
        st.info("Nenhum mês disponível. Feche ao menos um mês no Dashboard.")
        return

    mes_sel = st.selectbox("Selecione o mês", sorted(set(meses_disponiveis), key=lambda x: parse_mes(x) or pd.Timestamp.min),
                            format_func=mes_extenso, key="rel_mes_sel")
    mes_ts = parse_mes(mes_sel)
    if mes_ts is None:
        st.error("Mês inválido.")
        return

    resumo = calcular_resumo_mes(mes_ts)
    st.write("**Pré-visualização do relatório:**")
    col1, col2, col3 = st.columns(3)
    col1.metric("Faturamento", formatar_brl(resumo["Faturamento"]))
    col2.metric("Total Despesas", formatar_brl(resumo["Total Despesas"]))
    col3.metric("Saldo Livre", formatar_brl(resumo["Saldo Livre"]))

    pdf_bytes = gerar_pdf(mes_ts)
    st.download_button(
        f"📄 Baixar PDF — {mes_extenso(mes_ts)}",
        data=pdf_bytes,
        file_name=f"relatorio_barbearia_{mes_sel.replace('/', '_')}.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=True,
    )

    # Lista de todos relatórios disponíveis
    if len(meses_disponiveis) > 1:
        st.divider()
        st.subheader("📚 Histórico de Relatórios")
        for mes_label in sorted(set(meses_disponiveis), key=lambda x: parse_mes(x) or pd.Timestamp.min, reverse=True):
            m_ts = parse_mes(mes_label)
            if m_ts is None:
                continue
            res = calcular_resumo_mes(m_ts)
            c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
            c1.write(f"**{mes_extenso(m_ts)}**")
            c2.write(f"Fat: {formatar_brl(res['Faturamento'])}")
            c3.write(f"{cor_status(res['Status'])} {res['Status']}")
            with c4:
                pdf_b = gerar_pdf(m_ts)
                st.download_button("PDF", data=pdf_b,
                                   file_name=f"rel_{mes_label.replace('/', '_')}.pdf",
                                   mime="application/pdf",
                                   key=f"dl_pdf_{mes_label}")


def aba_base():
    st.header("💾 Base de Dados")
    st.write("Gerencie o backup e restauração completa dos dados do painel.")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("⬇️ Exportar Base Atual")
        st.write("Baixe a base completa para guardar o histórico e usar no próximo mês.")
        base_bytes = exportar_base()
        hoje = datetime.date.today().strftime("%Y_%m_%d")
        st.download_button(
            "⬇️ Baixar Base Completa (.xlsx)",
            data=base_bytes,
            file_name=f"base_barbearia_heloisa_{hoje}.xlsx",
            mime="application/xlsx",
            type="primary",
            use_container_width=True,
        )

    with col2:
        st.subheader("⬆️ Importar Base Salva")
        st.write("Restaure uma base exportada anteriormente para recuperar o histórico.")
        f_base = st.file_uploader("Selecione o arquivo base (.xlsx)", type=["xlsx"], key="up_base")
        if f_base and st.button("⬆️ Importar e Restaurar", type="primary"):
            if importar_base(f_base):
                st.success("✅ Base importada com sucesso!")
                st.rerun()

    st.divider()
    st.subheader("🔄 Importar Histórico AppBarber (primeira vez)")
    st.write("Se você vem do painel antigo, importe a base do AppBarber aqui.")
    f_app = st.file_uploader("Base AppBarber (.xlsx)", type=["xlsx"], key="up_appbarber")
    if f_app and st.button("📥 Processar Base AppBarber"):
        try:
            f_app.seek(0)
            xl = pd.ExcelFile(f_app)
            if "Historico" in xl.sheet_names:
                hist_old = xl.parse("Historico")
                hist_old["Mes"] = hist_old["Mes"].apply(parse_mes)
                # Mapear colunas do formato antigo para o novo
                col_map = {
                    "Faturamento Liquido": "Faturamento",
                    "Total Barbearia": "Despesas Barbearia",
                }
                hist_old = hist_old.rename(columns=col_map)
                if "Total Despesas" not in hist_old.columns:
                    hist_old["Total Despesas"] = hist_old.get("Despesas Barbearia", 0) + hist_old.get("Despesas Pessoais", 0)
                if "Saldo Livre" not in hist_old.columns:
                    hist_old["Saldo Livre"] = hist_old.get("Caixa Apos Guardar", 0)
                if "Meta Faturamento" not in hist_old.columns:
                    hist_old["Meta Faturamento"] = 9000.0
                if "Meta Despesas %" not in hist_old.columns:
                    hist_old["Meta Despesas %"] = 42.0
                if "Atingiu Meta" not in hist_old.columns:
                    hist_old["Atingiu Meta"] = False
                if "Status" not in hist_old.columns:
                    hist_old["Status"] = "Atenção"
                for col in COLUNAS_HISTORICO:
                    if col not in hist_old.columns:
                        hist_old[col] = 0.0 if col not in ["Status", "Atingiu Meta"] else False
                st.session_state.historico = hist_old[COLUNAS_HISTORICO].dropna(subset=["Mes"])
                st.success(f"✅ Histórico importado: {len(st.session_state.historico)} meses")
            if "Saidas" in xl.sheet_names or "Despesas" in xl.sheet_names:
                sheet = "Saidas" if "Saidas" in xl.sheet_names else "Despesas"
                saidas_old = xl.parse(sheet)
                saidas_old["Mes"] = saidas_old["Mes"].apply(parse_mes)
                st.session_state.saidas = pd.concat([saidas_default(), saidas_old], ignore_index=True)
                st.success(f"✅ Saídas importadas")
            if "Faturamento" in xl.sheet_names:
                fat_df = xl.parse("Faturamento")
                for _, r in fat_df.iterrows():
                    st.session_state.faturamento_mes[str(r["Mes"])] = float(r.get("Faturamento", 0))
            st.rerun()
        except Exception as e:
            st.error(f"Erro ao processar base AppBarber: {e}")

    # Info de estado atual
    st.divider()
    st.subheader("📊 Estado Atual dos Dados")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Meses no Histórico", len(st.session_state.historico))
    col2.metric("Saídas Registradas", len(st.session_state.saidas))
    col3.metric("Meses c/ Metas", len(st.session_state.metas))
    col4.metric("Meses c/ Faturamento", len(st.session_state.faturamento_mes))


# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Barbearia Heloisa Mazzi",
        page_icon="✂️",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # CSS customizado
    st.markdown("""
    <style>
    .stMetric { background: #f8f9fa; border-radius: 8px; padding: 12px; }
    .stMetric label { font-size: 13px !important; color: #555 !important; }
    [data-testid="stMetricValue"] { font-size: 20px !important; font-weight: bold !important; }
    .block-container { padding-top: 1rem; }
    </style>
    """, unsafe_allow_html=True)

    init_state()

    # Sidebar
    with st.sidebar:
        st.markdown("## ✂️ Barbearia Heloisa Mazzi")
        st.markdown("---")

        # Seleção de mês
        st.markdown("### 📅 Mês de Trabalho")
        meses_opcoes = []
        for m_offset in range(-3, 4):
            today = datetime.date.today()
            m = today.month + m_offset
            y = today.year
            while m < 1:
                m += 12; y -= 1
            while m > 12:
                m -= 12; y += 1
            meses_opcoes.append(label_mes(pd.Timestamp(y, m, 1)))

        mes_sel_label = st.selectbox(
            "Selecione o mês",
            meses_opcoes,
            index=meses_opcoes.index(st.session_state.mes_selecionado) if st.session_state.mes_selecionado in meses_opcoes else 3,
            format_func=mes_extenso,
            key="sidebar_mes"
        )
        st.session_state.mes_selecionado = mes_sel_label
        mes_ts = parse_mes(mes_sel_label)

        # Resumo rápido
        if mes_ts:
            resumo = calcular_resumo_mes(mes_ts)
            st.markdown("---")
            st.markdown("### 📊 Resumo Rápido")
            st.metric("Faturamento", formatar_brl(resumo["Faturamento"]))
            st.metric("Total Despesas", formatar_brl(resumo["Total Despesas"]))
            saldo = resumo["Saldo Livre"]
            st.metric("Saldo Livre", formatar_brl(saldo), delta=f"{'✓' if saldo >= 0 else '✗'}")
            status = resumo["Status"]
            emoji = cor_status(status)
            st.markdown(f"**Status:** {emoji} {status}")

        st.markdown("---")
        st.markdown("### 🔧 Ações Rápidas")
        if st.button("🔒 Fechar Mês Atual", use_container_width=True):
            if mes_ts:
                atualizar_historico(mes_ts)
                st.success("Mês fechado!")
                st.rerun()

        base_bytes = exportar_base()
        st.download_button(
            "⬇️ Baixar Base",
            data=base_bytes,
            file_name=f"base_barbearia_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/xlsx",
            use_container_width=True,
        )

        if mes_ts:
            pdf_bytes = gerar_pdf(mes_ts)
            st.download_button(
                "📄 Baixar PDF do Mês",
                data=pdf_bytes,
                file_name=f"rel_{mes_sel_label.replace('/', '_')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

    # Abas principais
    aba1, aba2, aba3, aba4, aba5, aba6 = st.tabs([
        "📊 Dashboard",
        "💸 Saídas",
        "📈 Histórico",
        "🎯 Metas",
        "🏆 Rankings",
        "💾 Base / PDF",
    ])

    if mes_ts is None:
        st.error("Mês inválido selecionado.")
        return

    with aba1:
        aba_dashboard(mes_ts)
    with aba2:
        aba_saidas(mes_ts)
    with aba3:
        aba_historico()
    with aba4:
        aba_metas()
    with aba5:
        aba_rankings()
    with aba6:
        col_rel, col_base = st.columns([1, 1])
        with col_rel:
            aba_relatorios()
        with col_base:
            aba_base()


if __name__ == "__main__":
    main()
