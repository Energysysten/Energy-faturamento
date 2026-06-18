import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

FATURAMENTO_PATH = "/Users/leonardocarmo/Library/CloudStorage/OneDrive-Pessoal/Energy/FATURAMENTO.xlsx"
CONTROLE_PATH = "/Users/leonardocarmo/Documents/Claude/Projects/Faturamento/Controle_Medicoes.xlsx"

CONTRATO_LABELS = {
    "4600027696": "ÂNCORA MANUT",
    "4600027637": "ÂNCORA OBRAS",
    "4600027590": "LINHA VIVA GO",
    "4600023515": "LINHA VIVA RS",
    "4600019416": "LINHA VIVA PA",
}

CORES = {
    "Faturado": "#22c55e",
    "Pendente": "#ef4444",
    "Em Processamento": "#f59e0b",
}

st.set_page_config(
    page_title="Dashboard Faturamento | Energy",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.kpi-card {
    background: #1e293b;
    border-radius: 12px;
    padding: 20px 24px;
    border-left: 4px solid;
    margin-bottom: 8px;
}
.kpi-value { font-size: 1.8rem; font-weight: 700; color: #f8fafc; margin: 4px 0; }
.kpi-label { font-size: 0.8rem; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.05em; }
.kpi-delta { font-size: 0.85rem; margin-top: 4px; }
.status-badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=300)
def load_faturamento():
    df = pd.read_excel(FATURAMENTO_PATH, sheet_name="index", header=1, usecols="B:S")
    df.columns = [
        "EMPRESA", "CONTRATO_NUM", "CONTRATO_NOME", "GESTOR",
        "OBRA", "CODIGO_CC", "COMPETENCIA", "MUNICIPIO",
        "VALOR_PROVISAO", "VALOR_BRUTO_FAT", "PEDIDO_FOLHA", "NOTA_FISCAL",
        "PROTOCOLO_UPLOAD", "VALOR_BRUTO", "RETENCAO", "IMPOSTOS",
        "VALOR_LIQUIDO", "DATA_RECEBIMENTO",
    ]
    df = df[df["EMPRESA"].notna()].copy()
    df["CONTRATO_NUM"] = df["CONTRATO_NUM"].astype(str).str.strip()
    df["COMPETENCIA"] = pd.to_datetime(df["COMPETENCIA"], errors="coerce")
    df["PERIODO"] = df["COMPETENCIA"].dt.to_period("M").astype(str)
    df["PEDIDO_FOLHA"] = df["PEDIDO_FOLHA"].apply(
        lambda x: str(int(float(x))) if pd.notna(x) and str(x).strip() not in ("", "nan") else None
    )
    df["VALOR_PROVISAO"] = pd.to_numeric(df["VALOR_PROVISAO"], errors="coerce").fillna(0)
    df["VALOR_LIQUIDO"] = pd.to_numeric(df["VALOR_LIQUIDO"], errors="coerce").fillna(0)
    return df


@st.cache_data(ttl=300)
def load_controle():
    df = pd.read_excel(CONTROLE_PATH, sheet_name="Medições")
    df.columns = [
        "DATA_RECEBIMENTO", "N_FOLHA", "N_CONTRATO", "PERIODO_INICIO",
        "PERIODO_FIM", "MUNICIPIO", "FORNECEDOR", "VALOR_TOTAL", "ARQUIVO", "STATUS",
    ]
    df = df[df["N_FOLHA"].notna()].copy()
    df["N_CONTRATO"] = df["N_CONTRATO"].astype(str).str.strip()
    df["N_FOLHA"] = df["N_FOLHA"].astype(str).str.strip()
    df["PERIODO_INICIO"] = pd.to_datetime(df["PERIODO_INICIO"], format="%d/%m/%Y", errors="coerce")
    df["PERIODO"] = df["PERIODO_INICIO"].dt.to_period("M").astype(str)
    df["VALOR_TOTAL"] = pd.to_numeric(df["VALOR_TOTAL"], errors="coerce").fillna(0)
    return df


def fmt_brl(value):
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct_bar(pct):
    color = "#22c55e" if pct >= 80 else "#f59e0b" if pct >= 40 else "#ef4444"
    return f'<div style="background:#334155;border-radius:4px;height:8px;width:100%"><div style="background:{color};border-radius:4px;height:8px;width:{min(pct,100):.0f}%"></div></div>'


def build_merged(fat, ctrl):
    """Merge faturamento (planned) with controle (actual) at contract+period level."""
    fat_agg = (
        fat.groupby(["CONTRATO_NUM", "CONTRATO_NOME", "PERIODO"])
        .agg(PREVISTO=("VALOR_PROVISAO", "sum"), LINHAS=("OBRA", "count"))
        .reset_index()
    )
    ctrl_agg = (
        ctrl.groupby(["N_CONTRATO", "PERIODO"])
        .agg(REALIZADO=("VALOR_TOTAL", "sum"), FOLHAS=("N_FOLHA", "count"))
        .reset_index()
        .rename(columns={"N_CONTRATO": "CONTRATO_NUM"})
    )
    merged = fat_agg.merge(ctrl_agg, on=["CONTRATO_NUM", "PERIODO"], how="outer")
    merged["PREVISTO"] = merged["PREVISTO"].fillna(0)
    merged["REALIZADO"] = merged["REALIZADO"].fillna(0)
    merged["DIFERENCA"] = merged["REALIZADO"] - merged["PREVISTO"]
    merged["PCT"] = merged.apply(
        lambda r: (r["REALIZADO"] / r["PREVISTO"] * 100) if r["PREVISTO"] > 0 else 0, axis=1
    )
    return merged


def enrich_fat_status(fat, ctrl_folhas):
    """Add STATUS column to faturamento rows."""
    def get_status(row):
        if row["PEDIDO_FOLHA"] and row["PEDIDO_FOLHA"] in ctrl_folhas:
            return "Faturado"
        if row["PEDIDO_FOLHA"]:
            return "Em Processamento"
        return "Pendente"
    fat = fat.copy()
    fat["STATUS"] = fat.apply(get_status, axis=1)
    return fat


# ── Load data ──────────────────────────────────────────────────────────────────
try:
    fat = load_faturamento()
    ctrl = load_controle()
except Exception as e:
    st.error(f"Erro ao carregar planilhas: {e}")
    st.stop()

ctrl_folhas = set(ctrl["N_FOLHA"].tolist())
fat = enrich_fat_status(fat, ctrl_folhas)
merged = build_merged(fat, ctrl)

# ── Header ─────────────────────────────────────────────────────────────────────
col_title, col_reload = st.columns([5, 1])
with col_title:
    st.title("⚡ Dashboard de Faturamento")
    st.caption("Energy Construções & Serviços — Previsto vs Realizado")
with col_reload:
    st.write("")
    if st.button("🔄 Recarregar dados"):
        st.cache_data.clear()
        st.rerun()

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(
    ["📊 Visão Geral", "🔍 Previsto vs Realizado", "⚠️ Itens Pendentes", "📥 Medições Recebidas"]
)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — VISÃO GERAL
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    total_previsto = fat["VALOR_PROVISAO"].sum()
    total_realizado = ctrl["VALOR_TOTAL"].sum()
    pct_realizado = (total_realizado / total_previsto * 100) if total_previsto > 0 else 0
    valor_pendente = fat[fat["STATUS"] == "Pendente"]["VALOR_PROVISAO"].sum()
    n_pendentes = (fat["STATUS"] == "Pendente").sum()
    n_faturados = (fat["STATUS"] == "Faturado").sum()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"""
        <div class="kpi-card" style="border-color:#3b82f6">
            <div class="kpi-label">Total Previsto</div>
            <div class="kpi-value">{fmt_brl(total_previsto)}</div>
            <div class="kpi-delta" style="color:#94a3b8">{len(fat)} obras no cronograma</div>
        </div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="kpi-card" style="border-color:#22c55e">
            <div class="kpi-label">Total Realizado</div>
            <div class="kpi-value">{fmt_brl(total_realizado)}</div>
            <div class="kpi-delta" style="color:#22c55e">{len(ctrl)} medições recebidas</div>
        </div>""", unsafe_allow_html=True)
    with k3:
        cor_pct = "#22c55e" if pct_realizado >= 80 else "#f59e0b" if pct_realizado >= 40 else "#ef4444"
        st.markdown(f"""
        <div class="kpi-card" style="border-color:{cor_pct}">
            <div class="kpi-label">% Realizado</div>
            <div class="kpi-value" style="color:{cor_pct}">{pct_realizado:.1f}%</div>
            <div class="kpi-delta" style="color:#94a3b8">{n_faturados} itens com folha confirmada</div>
        </div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""
        <div class="kpi-card" style="border-color:#ef4444">
            <div class="kpi-label">Valor Pendente</div>
            <div class="kpi-value" style="color:#ef4444">{fmt_brl(valor_pendente)}</div>
            <div class="kpi-delta" style="color:#ef4444">{n_pendentes} itens sem folha</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    c_left, c_right = st.columns([3, 2])

    with c_left:
        st.subheader("Previsto vs Realizado por Contrato")
        agg_contrato = merged.groupby(["CONTRATO_NUM", "CONTRATO_NOME"]).agg(
            PREVISTO=("PREVISTO", "sum"), REALIZADO=("REALIZADO", "sum")
        ).reset_index()
        agg_contrato["LABEL"] = agg_contrato["CONTRATO_NOME"].fillna(
            agg_contrato["CONTRATO_NUM"].map(CONTRATO_LABELS).fillna(agg_contrato["CONTRATO_NUM"])
        )
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(
            name="Previsto", x=agg_contrato["LABEL"], y=agg_contrato["PREVISTO"],
            marker_color="#3b82f6", opacity=0.85,
        ))
        fig_bar.add_trace(go.Bar(
            name="Realizado", x=agg_contrato["LABEL"], y=agg_contrato["REALIZADO"],
            marker_color="#22c55e", opacity=0.85,
        ))
        fig_bar.update_layout(
            barmode="group", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", y=1.1), margin=dict(l=0, r=0, t=30, b=0),
            yaxis=dict(tickformat=",.0f", gridcolor="#334155"),
            xaxis=dict(gridcolor="#334155"),
            height=320,
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with c_right:
        st.subheader("Distribuição por Status")
        status_counts = fat["STATUS"].value_counts().reset_index()
        status_counts.columns = ["Status", "Quantidade"]
        fig_pie = px.pie(
            status_counts, names="Status", values="Quantidade",
            color="Status", color_discrete_map=CORES,
            hole=0.55,
        )
        fig_pie.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", y=-0.1),
            margin=dict(l=0, r=0, t=10, b=0), height=320,
        )
        fig_pie.update_traces(textinfo="percent+label")
        st.plotly_chart(fig_pie, use_container_width=True)

    # Monthly evolution
    st.subheader("Evolução Mensal")
    monthly_previsto = fat.groupby("PERIODO")["VALOR_PROVISAO"].sum().reset_index()
    monthly_realizado = ctrl.groupby("PERIODO")["VALOR_TOTAL"].sum().reset_index()
    monthly = monthly_previsto.merge(
        monthly_realizado.rename(columns={"VALOR_TOTAL": "REALIZADO", "PERIODO": "PERIODO"}),
        on="PERIODO", how="outer"
    ).fillna(0).sort_values("PERIODO")
    fig_line = go.Figure()
    fig_line.add_trace(go.Scatter(
        x=monthly["PERIODO"], y=monthly["VALOR_PROVISAO"],
        name="Previsto", mode="lines+markers", line=dict(color="#3b82f6", width=2),
        fill="tozeroy", fillcolor="rgba(59,130,246,0.1)",
    ))
    fig_line.add_trace(go.Scatter(
        x=monthly["PERIODO"], y=monthly["REALIZADO"],
        name="Realizado", mode="lines+markers", line=dict(color="#22c55e", width=2),
        fill="tozeroy", fillcolor="rgba(34,197,94,0.1)",
    ))
    fig_line.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=1.1), margin=dict(l=0, r=0, t=30, b=0),
        yaxis=dict(tickformat=",.0f", gridcolor="#334155"),
        xaxis=dict(gridcolor="#334155"),
        height=280,
    )
    st.plotly_chart(fig_line, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PREVISTO VS REALIZADO (detalhado)
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Análise Detalhada por Contrato / Período")

    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    empresas = ["Todas"] + sorted(fat["EMPRESA"].dropna().unique().tolist())
    contratos = ["Todos"] + sorted(
        [f"{k} — {v}" for k, v in CONTRATO_LABELS.items() if k in fat["CONTRATO_NUM"].values]
    )
    gestores = ["Todos"] + sorted(fat["GESTOR"].dropna().unique().tolist())
    periodos = ["Todos"] + sorted(fat["PERIODO"].dropna().unique().tolist())

    with col_f1:
        sel_empresa = st.selectbox("Empresa", empresas)
    with col_f2:
        sel_contrato = st.selectbox("Contrato", contratos)
    with col_f3:
        sel_gestor = st.selectbox("Gestor", gestores)
    with col_f4:
        sel_periodo = st.selectbox("Período", periodos)

    fat_f = fat.copy()
    if sel_empresa != "Todas":
        fat_f = fat_f[fat_f["EMPRESA"] == sel_empresa]
    if sel_contrato != "Todos":
        num = sel_contrato.split(" — ")[0]
        fat_f = fat_f[fat_f["CONTRATO_NUM"] == num]
    if sel_gestor != "Todos":
        fat_f = fat_f[fat_f["GESTOR"] == sel_gestor]
    if sel_periodo != "Todos":
        fat_f = fat_f[fat_f["PERIODO"] == sel_periodo]

    # Enrich with realizado from controle using folha match
    ctrl_folha_val = ctrl.set_index("N_FOLHA")["VALOR_TOTAL"].to_dict()
    fat_f = fat_f.copy()
    fat_f["REALIZADO"] = fat_f["PEDIDO_FOLHA"].map(ctrl_folha_val).fillna(0)
    fat_f["DIFERENCA"] = fat_f["REALIZADO"] - fat_f["VALOR_PROVISAO"]

    # Summary row
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("Total Previsto (filtro)", fmt_brl(fat_f["VALOR_PROVISAO"].sum()))
    with m2:
        st.metric("Total Realizado (folhas confirmadas)", fmt_brl(fat_f["REALIZADO"].sum()))
    with m3:
        delta = fat_f["REALIZADO"].sum() - fat_f["VALOR_PROVISAO"].sum()
        st.metric("Diferença", fmt_brl(delta), delta=f"{delta:+,.2f}")

    # Table
    display = fat_f[[
        "EMPRESA", "CONTRATO_NOME", "GESTOR", "OBRA", "PERIODO",
        "VALOR_PROVISAO", "REALIZADO", "DIFERENCA", "PEDIDO_FOLHA", "STATUS"
    ]].copy()
    display.columns = [
        "Empresa", "Contrato", "Gestor", "Obra", "Período",
        "Previsto (R$)", "Realizado (R$)", "Diferença (R$)", "Nº Folha", "Status"
    ]
    display["Previsto (R$)"] = display["Previsto (R$)"].apply(fmt_brl)
    display["Realizado (R$)"] = display["Realizado (R$)"].apply(fmt_brl)
    display["Diferença (R$)"] = display["Diferença (R$)"].apply(fmt_brl)

    def highlight_status(row):
        c = CORES.get(row["Status"], "#94a3b8")
        return [f"color: {c}" if col == "Status" else "" for col in row.index]

    st.dataframe(
        display.style.apply(highlight_status, axis=1),
        use_container_width=True, height=450, hide_index=True,
    )

    # Export
    buf = BytesIO()
    fat_f[[
        "EMPRESA", "CONTRATO_NOME", "GESTOR", "OBRA", "PERIODO",
        "VALOR_PROVISAO", "REALIZADO", "DIFERENCA", "PEDIDO_FOLHA", "STATUS"
    ]].to_excel(buf, index=False, sheet_name="Previsto_vs_Realizado")
    st.download_button(
        "📥 Exportar Excel",
        data=buf.getvalue(),
        file_name="previsto_vs_realizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — ITENS PENDENTES
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("⚠️ Obras Previstas sem Medição Recebida")
    st.caption("Ordenado por valor — priorize os maiores para maximizar o faturamento.")

    pendentes = fat[fat["STATUS"] == "Pendente"].copy()
    pendentes = pendentes.sort_values("VALOR_PROVISAO", ascending=False)

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.metric("Quantidade de itens pendentes", len(pendentes))
    with col_p2:
        st.metric("Valor total pendente", fmt_brl(pendentes["VALOR_PROVISAO"].sum()))

    st.markdown("---")

    # Top 10 chart
    top10 = pendentes.head(10).copy()
    top10["LABEL"] = top10["OBRA"].str[:40] + "  (" + top10["CONTRATO_NOME"].fillna("") + ")"
    fig_pend = px.bar(
        top10.sort_values("VALOR_PROVISAO"),
        x="VALOR_PROVISAO", y="LABEL", orientation="h",
        color_discrete_sequence=["#ef4444"],
        labels={"VALOR_PROVISAO": "Valor Previsto (R$)", "LABEL": ""},
        title="Top 10 obras com maior valor pendente",
    )
    fig_pend.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=0, r=0, t=40, b=0), height=350,
        xaxis=dict(tickformat=",.0f", gridcolor="#334155"),
    )
    st.plotly_chart(fig_pend, use_container_width=True)

    # Table
    display_pend = pendentes[[
        "EMPRESA", "CONTRATO_NOME", "GESTOR", "OBRA", "PERIODO",
        "MUNICIPIO", "VALOR_PROVISAO", "CODIGO_CC"
    ]].copy()
    display_pend.columns = [
        "Empresa", "Contrato", "Gestor", "Obra", "Período",
        "Município", "Valor Previsto (R$)", "Código CC"
    ]
    display_pend["Valor Previsto (R$)"] = display_pend["Valor Previsto (R$)"].apply(fmt_brl)
    st.dataframe(display_pend, use_container_width=True, height=400, hide_index=True)

    # Export pendentes
    buf2 = BytesIO()
    pendentes[[
        "EMPRESA", "CONTRATO_NOME", "GESTOR", "OBRA", "PERIODO",
        "MUNICIPIO", "VALOR_PROVISAO", "CODIGO_CC"
    ]].to_excel(buf2, index=False, sheet_name="Pendentes")
    st.download_button(
        "📥 Exportar Lista de Pendentes",
        data=buf2.getvalue(),
        file_name="itens_pendentes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — MEDIÇÕES RECEBIDAS
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("📥 Medições Processadas Automaticamente")
    st.caption("Extraídas via e-mail pelo sistema de automação.")

    m_c1, m_c2, m_c3 = st.columns(3)
    with m_c1:
        st.metric("Total de medições", len(ctrl))
    with m_c2:
        st.metric("Valor total medido", fmt_brl(ctrl["VALOR_TOTAL"].sum()))
    with m_c3:
        if ctrl["PERIODO_INICIO"].notna().any():
            ultima = ctrl["PERIODO_INICIO"].max()
            st.metric("Período mais recente", ultima.strftime("%m/%Y") if pd.notna(ultima) else "—")

    st.markdown("---")

    # By contract
    ctrl_by_cont = (
        ctrl.groupby("N_CONTRATO")
        .agg(TOTAL=("VALOR_TOTAL", "sum"), QTD=("N_FOLHA", "count"))
        .reset_index()
        .sort_values("TOTAL", ascending=False)
    )
    ctrl_by_cont["NOME"] = ctrl_by_cont["N_CONTRATO"].map(CONTRATO_LABELS).fillna(ctrl_by_cont["N_CONTRATO"])
    ctrl_by_cont["LABEL"] = ctrl_by_cont["NOME"] + "\n(" + ctrl_by_cont["N_CONTRATO"] + ")"

    fig_cont = px.bar(
        ctrl_by_cont, x="LABEL", y="TOTAL",
        color_discrete_sequence=["#22c55e"],
        labels={"TOTAL": "Valor Total (R$)", "LABEL": "Contrato"},
        title="Valor total por contrato",
    )
    fig_cont.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=0, r=0, t=40, b=0), height=300,
        yaxis=dict(tickformat=",.0f", gridcolor="#334155"),
    )
    st.plotly_chart(fig_cont, use_container_width=True)

    # Full table
    display_ctrl = ctrl[[
        "DATA_RECEBIMENTO", "N_FOLHA", "N_CONTRATO", "PERIODO",
        "MUNICIPIO", "FORNECEDOR", "VALOR_TOTAL", "STATUS"
    ]].copy()
    display_ctrl["N_CONTRATO_NOME"] = display_ctrl["N_CONTRATO"].map(CONTRATO_LABELS).fillna("")
    display_ctrl["VALOR_TOTAL"] = display_ctrl["VALOR_TOTAL"].apply(fmt_brl)
    display_ctrl = display_ctrl[[
        "DATA_RECEBIMENTO", "N_FOLHA", "N_CONTRATO", "N_CONTRATO_NOME",
        "PERIODO", "MUNICIPIO", "VALOR_TOTAL", "STATUS"
    ]]
    display_ctrl.columns = [
        "Recebido em", "Nº Folha", "Nº Contrato", "Contrato",
        "Período", "Município", "Valor Total (R$)", "Status"
    ]
    st.dataframe(display_ctrl, use_container_width=True, height=450, hide_index=True)
