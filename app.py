import streamlit as st
import pdfplumber
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook
from datetime import date, datetime, timedelta
import re
import io

# ══════════════════════════════════════════════════════════════
# CONFIGURAÇÃO
# ══════════════════════════════════════════════════════════════
STORES = {
    "SIDE":    "West Side / Pompéia",
    "ZONE":    "West Zone / Sonda",
    "PLACE":   "West Place / Girassol",
    "STATION": "West Station / Paulistânia",
}
COLORS = {"SIDE": "#3498db", "ZONE": "#e74c3c", "PLACE": "#2ecc71", "STATION": "#f39c12"}

# Palavras-chave para identificar cada loja no PDF de fechamento
STORE_KEYWORDS = {
    "SIDE":    ["POMPEIA, 1700", "AV. POMPEIA, 1700", "AV POMPEIA 1700", "3672 1466", "3672 1465"],
    "ZONE":    ["CARLOS VICARI", "SONDA POMPEIA", "3675 4295"],
    "PLACE":   ["GIRASSOL"],   # Adicione endereço/telefone quando souber
    "STATION": ["PAULISTANIA, 91", "PAULISTANIA 91", "PAULISTÂNIA", "3675 4094"],
}

# Mapeamento campo → coluna no Excel de fechamento
EXCEL_COL = {
    'pecas': 4, 'servicos': 5, 'fatu': 6, 'apagar': 7,
    'din': 10, 'cheque': 11, 'cc': 12, 'cd': 13,
    'dep': 14, 'ecom': 15, 'outros': 16, 'leitura': 18,
    'agua': 21, 'mercado': 22, 'cafe': 23, 'pedagio': 24,
    'farmacia': 25, 'bolo': 26, 'banco': 27, 'outros_sangria': 28,
    'fundo': 29,
}

st.set_page_config(
    page_title="5ASEC | Fechamento de Caixa",
    page_icon="🧺", layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado
st.markdown("""
<style>
.store-header { font-size: 1.1rem; font-weight: 700; margin-bottom: 0.5rem; }
.ref-box { background: #f0f4ff; border-left: 3px solid #3498db;
           padding: 6px 10px; border-radius: 4px; font-size: 0.82rem; margin: 4px 0; }
.kpi-delta { font-size: 0.75rem; color: #666; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════
def to_float(s):
    if s is None: return None
    try:
        return float(str(s).strip().replace('.', '').replace(',', '.'))
    except:
        return None

def fmt_brl(v):
    if v is None or v == 0: return "—"
    return f"R$ {v:_.2f}".replace('.', ',').replace('_', '.')

def get_val(sd, key):
    v = sd.get(key)
    return float(v) if v is not None else 0.0

def total_sangria(sd):
    keys = ['agua', 'mercado', 'cafe', 'pedagio', 'farmacia', 'bolo', 'banco', 'outros_sangria']
    return sum(get_val(sd, k) for k in keys)

def total_recebido(sd):
    keys = ['din', 'cheque', 'cc', 'cd', 'dep', 'ecom', 'outros']
    return sum(get_val(sd, k) for k in keys)

# ══════════════════════════════════════════════════════════════
# PARSERS
# ══════════════════════════════════════════════════════════════
def parse_br(text, pattern):
    m = re.search(pattern, text, re.IGNORECASE)
    return to_float(m.group(1)) if m else None

def identify_store(text):
    up = text.upper()
    for sid, kws in STORE_KEYWORDS.items():
        if any(k.upper() in up for k in kws):
            return sid
    return None

def parse_pdf_fechamento(file):
    """Parseia PDF de fechamento do sistema 5ASEC Windows."""
    try:
        with pdfplumber.open(file) as pdf:
            text = '\n'.join(p.extract_text() or '' for p in pdf.pages)
    except Exception as e:
        return None, f"Erro ao ler PDF: {e}"
    if not text.strip():
        return None, "PDF sem texto legível (imagem?). Use entrada manual."

    store = identify_store(text)

    # Data do relatório
    m = re.search(r'DATA BASE\s+(\d{2}/\d{2}/\d{4})', text, re.I)
    rep_date = datetime.strptime(m.group(1), '%d/%m/%Y').date() if m else None
    # Seção TOTAL GERAL (entre "TOTAL GERAL" e "V FATURADO")
    m_tg = re.search(r'TOTAL\s+GERAL(.*?)(?:V[\s.]+FATURADO|V FATURADO)', text, re.I | re.S)
    tg = m_tg.group(1) if m_tg else text

    # Seção DEPOSITOS
    m_dep = re.search(r'\*\s*DEPOSITOS(.*?)(?:\*\s*SANGRIAS)', text, re.I | re.S)
    dep_sec = m_dep.group(1) if m_dep else ""

    # Seção SANGRIAS
    m_san = re.search(r'\*\s*SANGRIAS(.*?)(?:TICKETS\s+ANUL|OPERADOR|$)', text, re.I | re.S)
    san_sec = m_san.group(1) if m_san else ""

    d = {
        'store': store, 'date': rep_date,
        'pecas':     parse_br(text, r'QTDE\.\s*DE\s*PE[CÇ]AS\s*:?\s*([\d.,]+)'),
        'servicos':  parse_br(text, r'QTDE\.\s*DE\s*SERV\.\s*:?\s*([\d.,]+)'),
        'fatu':      parse_br(text, r'FATU\.\s*LI[QG]UIDO\s*:?\s*([\d.,]+)'),
        'apagar':    parse_br(text, r'A\s*PAGAR\s*\(DIA\)\s*:?\s*([\d.,]+)'),
        'apagar_ac': parse_br(text, r'A\s*PAGAR\s*\(AC\.\)\s*:?\s*([\d.,]+)'),
        'leitura':   parse_br(text, r'VALOR\s*NO\s*CAIXA\s*:?\s*([\d.,]+)'),
        'outros':    parse_br(text, r'CONSUMO\s*SALDO\s*:?\s*([\d.,]+)') or
                     parse_br(text, r'BAIXA\s*DE\s*SALDO\s*:?\s*([\d.,]+)'),
        'desc_emis': parse_br(text, r'DESCONTO\s*EMIS\.\s*:?\s*([\d.,]+)'),
        # Do TOTAL GERAL (referência — os valores definitivos vêm da Rede/Itaú)
        'cc_ref':    parse_br(tg, r'CART[AÃ]O\s*CR[EÉ]DITO\s*:?\s*([\d.,]+)'),
        'cd_ref':    parse_br(tg, r'CART[AÃ]O\s*D[EÉ]BITO\s*:?\s*([\d.,]+)'),
        'dep_ref':   parse_br(tg, r'DEP[OÓ]SITO\s*:?\s*([\d.,]+)'),
        'din':       parse_br(tg, r'DINHEIRO\s*:?\s*([\d.,]+)'),
        # Seção DEPOSITOS
        'fundo':     parse_br(dep_sec, r'DEPOSITO\s*CAIXA\s*:?\s*([\d.,]+)'),
        # Seção SANGRIAS
        'sangria_total': parse_br(san_sec, r'SANGRIA\s*TOTAL\s*:?\s*([\d.,]+)'),
        'operador':  None,
    }
    op = re.search(r'OPERADOR\s+([A-Z][A-Z\s]+)', text, re.I)
    if op:
        d['operador'] = op.group(1).strip()

    return d, None


def parse_excel_rede(file):
    """
    Parseia o Excel exportado do portal Rede.
    Detecta automaticamente colunas de crédito e débito por valor/nome.
    Retorna: dict {store_id: {cc: float, cd: float}} ou totais gerais.
    """
    try:
        df = pd.read_excel(file, header=None)
    except Exception as e:
        return None, f"Erro ao ler Excel da Rede: {e}"

    # Converte tudo para string para busca
    text_full = df.astype(str).to_string().upper()

    result = {'cc': None, 'cd': None, 'raw_df': df}

    # Tenta detectar linhas com totais de crédito e débito
    # Padrões comuns no relatório Rede:
    cc_patterns = [
        r'(?:CREDITO|CR[EÉ]DITO|CREDIT)[^\d]*([\d.,]+)',
        r'(?:TOTAL\s+CR)[^\d]*([\d.,]+)',
    ]
    cd_patterns = [
        r'(?:DEBITO|D[EÉ]BITO|DEBIT)[^\d]*([\d.,]+)',
        r'(?:TOTAL\s+DB)[^\d]*([\d.,]+)',
    ]

    for pat in cc_patterns:
        v = parse_br(text_full, pat)
        if v and v > 10:   # Filtra valores muito pequenos (cabeçalhos)
            result['cc'] = v
            break

    for pat in cd_patterns:
        v = parse_br(text_full, pat)
        if v and v > 0:
            result['cd'] = v
            break

    return result, None


def read_historical_excel(wb, sheet_name):
    """Lê dados históricos do template de fechamento para gráficos de tendência."""
    if sheet_name not in wb.sheetnames:
        return pd.DataFrame()
    ws = wb[sheet_name]
    ref_date = date(2026, 1, 1)
    rows = []
    for row_num in range(2, 368):
        fatu = ws.cell(row=row_num, column=6).value
        pecas = ws.cell(row=row_num, column=4).value
        if fatu is None and pecas is None:
            continue
        try:
            fatu = float(fatu) if fatu else 0.0
            pecas = int(pecas) if pecas else 0
        except:
            continue
        if fatu == 0 and pecas == 0:
            continue
        day = ref_date + timedelta(days=row_num - 2)
        rows.append({
            'data': day, 'loja': sheet_name,
            'fatu': fatu, 'pecas': pecas,
            'apagar': ws.cell(row=row_num, column=7).value or 0,
            'din': ws.cell(row=row_num, column=10).value or 0,
            'cc': ws.cell(row=row_num, column=12).value or 0,
            'cd': ws.cell(row=row_num, column=13).value or 0,
            'dep': ws.cell(row=row_num, column=14).value or 0,
        })
    return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════
ALL_FIELDS = [
    'pecas', 'servicos', 'fatu', 'apagar', 'din', 'cheque',
    'cc', 'cd', 'dep', 'ecom', 'outros', 'leitura', 'fundo',
    'agua', 'mercado', 'cafe', 'pedagio', 'farmacia', 'bolo',
    'banco', 'outros_sangria',
    '_cc_ref', '_cd_ref', '_dep_ref', '_operador', '_apagar_ac',
]

def empty_store():
    return {f: None for f in ALL_FIELDS}

if 'data' not in st.session_state:
    st.session_state.data = {s: empty_store() for s in STORES}
if 'sel_date' not in st.session_state:
    st.session_state.sel_date = date.today()
if 'hist_wb' not in st.session_state:
    st.session_state.hist_wb = None

# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🧺 5ASEC Fechamento")
    st.divider()

    sel_date = st.date_input("📅 Data do fechamento", value=st.session_state.sel_date, format="DD/MM/YYYY")
    st.session_state.sel_date = sel_date

    st.divider()
    st.markdown("### 📎 Relatórios de Fechamento")
    st.caption("PDFs das lojas (sistema Windows)")
    pdf_files = st.file_uploader(
        "Arraste os PDFs aqui", type=["pdf"],
        accept_multiple_files=True, key="pdf_upload",
        label_visibility="collapsed"
    )
    if pdf_files and st.button("🔄 Processar PDFs", type="primary", use_container_width=True):
        unmatched = []
        for f in pdf_files:
            with st.spinner(f"Processando {f.name}..."):
                d, err = parse_pdf_fechamento(f)
            if err:
                st.error(f"❌ {f.name}: {err}")
                continue
            store = d.get('store')
            if not store:
                unmatched.append((f.name, d))
                continue
            sd = st.session_state.data[store]
            for k in ['pecas', 'servicos', 'fatu', 'apagar', 'din', 'outros', 'leitura', 'fundo']:
                if d.get(k) is not None:
                    sd[k] = d[k]
            sd['_cc_ref'] = d.get('cc_ref')
            sd['_cd_ref'] = d.get('cd_ref')
            sd['_dep_ref'] = d.get('dep_ref')
            sd['_operador'] = d.get('operador')
            sd['_apagar_ac'] = d.get('apagar_ac')
            st.success(f"✅ {STORES[store]}")

        # Lojas não identificadas: pede seleção manual
        for fname, d in unmatched:
            st.warning(f"Loja não identificada: **{fname}**")
            manual_store = st.selectbox(
                "Selecione a loja:", list(STORES.keys()),
                format_func=lambda x: STORES[x], key=f"ms_{fname}"
            )
            if st.button(f"✅ Confirmar", key=f"conf_{fname}"):
                sd = st.session_state.data[manual_store]
                for k in ['pecas', 'servicos', 'fatu', 'apagar', 'din', 'outros', 'leitura', 'fundo']:
                    if d.get(k) is not None:
                        sd[k] = d[k]
                st.rerun()

    st.divider()
    st.markdown("### 💳 Excel da Rede")
    st.caption("Exportado do portal Rede — preenche CC e débito automaticamente")
    rede_file = st.file_uploader("Excel Rede", type=["xlsx", "xls"], key="rede_upload", label_visibility="collapsed")
    if rede_file:
        rede_store = st.selectbox("Qual loja?", list(STORES.keys()), format_func=lambda x: STORES[x], key="rede_store")
        if st.button("📥 Importar Rede", use_container_width=True):
            r, err = parse_excel_rede(rede_file)
            if err:
                st.error(err)
            else:
                sd = st.session_state.data[rede_store]
                if r.get('cc') is not None:
                    sd['cc'] = r['cc']
                    st.success(f"✅ Crédito: {fmt_brl(r['cc'])}")
                if r.get('cd') is not None:
                    sd['cd'] = r['cd']
                    st.success(f"✅ Débito: {fmt_brl(r['cd'])}")
                if r.get('cc') is None and r.get('cd') is None:
                    st.warning("Não foi possível detectar totais automaticamente. Confira o arquivo ou insira manualmente.")
                    # Mostra prévia do Excel para o usuário identificar a estrutura
                    st.dataframe(r['raw_df'].head(20), use_container_width=True)

    st.divider()
    st.markdown("### 📊 Histórico (opcional)")
    st.caption("Carregue o Excel de fechamento para ver tendências")
    hist_file = st.file_uploader("Template de fechamento", type=["xlsx"], key="hist_upload", label_visibility="collapsed")
    if hist_file:
        try:
            st.session_state.hist_wb = load_workbook(hist_file, data_only=True)
            st.success("✅ Histórico carregado")
        except Exception as e:
            st.error(f"Erro: {e}")

    st.divider()
    st.caption("🔵 = valores da Rede  |  🟠 = Itaú/banco")

# ══════════════════════════════════════════════════════════════
# MAIN TABS
# ══════════════════════════════════════════════════════════════
tab1, tab2, tab3 = st.tabs(["📥 Entrada de Dados", "📊 Dashboard", "💾 Exportar"])

# ──────────────────────────────────────────────────────────────
# TAB 1 — ENTRADA DE DADOS
# ──────────────────────────────────────────────────────────────
with tab1:
    st.header(f"Fechamento — {sel_date.strftime('%d/%m/%Y')}")

    def ni(label, key, val=None, is_int=False):
        """Input numérico padronizado."""
        return st.number_input(
            label, min_value=0.0,
            value=float(val) if val is not None else 0.0,
            step=1.0 if is_int else 0.01,
            format="%.0f" if is_int else "%.2f",
            key=key
        )

    for store_id, store_name in STORES.items():
        sd = st.session_state.data[store_id]
        color = COLORS[store_id]

        with st.expander(f"🏪 **{store_name}**", expanded=True):
            if sd.get('_operador'):
                st.caption(f"Operador detectado: **{sd['_operador']}**")

            c1, c2, c3, c4 = st.columns(4)

            with c1:
                st.markdown("**📦 Produção**")
                sd['pecas']    = ni("Nº Peças",    f"{store_id}_pec", sd['pecas'],    is_int=True)
                sd['servicos'] = ni("Nº Serviços", f"{store_id}_srv", sd['servicos'], is_int=True)
                st.markdown("**💰 Faturamento**")
                sd['fatu']   = ni("Fat. Líquido",  f"{store_id}_fatu",   sd['fatu'])
                sd['apagar'] = ni("A Pagar (Dia)", f"{store_id}_apagar", sd['apagar'])
                if sd.get('_apagar_ac'):
                    st.metric("A Pagar Acumulado", fmt_brl(sd['_apagar_ac']))

            with c2:
                st.markdown("**💳 Recebimentos**")
                sd['din']    = ni("Dinheiro",         f"{store_id}_din",  sd['din'])
                sd['cheque'] = ni("Cheque",            f"{store_id}_chq",  sd['cheque'])
                sd['cc']     = ni("Crédito 🔵",        f"{store_id}_cc",   sd['cc'])
                sd['cd']     = ni("Débito 🔵",         f"{store_id}_cd",   sd['cd'])
                sd['dep']    = ni("Depósito/PIX 🟠",   f"{store_id}_dep",  sd['dep'])
                sd['ecom']   = ni("E-commerce",        f"{store_id}_ecom", sd['ecom'])
                sd['outros'] = ni("Outros (Saldo)",    f"{store_id}_out",  sd['outros'])

                # Referências do PDF
                refs = [(sd.get('_cc_ref'), "CC"), (sd.get('_cd_ref'), "CD"), (sd.get('_dep_ref'), "Dep")]
                ref_lines = [f"<b>{label}:</b> {fmt_brl(v)}" for v, label in refs if v]
                if ref_lines:
                    st.markdown(
                        '<div class="ref-box">📋 Referência PDF:<br>' + " &nbsp;|&nbsp; ".join(ref_lines) + "</div>",
                        unsafe_allow_html=True
                    )

            with c3:
                st.markdown("**📊 Caixa**")
                sd['leitura'] = ni("Leitura X",      f"{store_id}_lx",    sd['leitura'])
                sd['fundo']   = ni("Fundo Abertura",  f"{store_id}_fundo", sd['fundo'])
                st.divider()
                t_san = total_sangria(sd)
                t_rec = total_recebido(sd)
                fatu_v = get_val(sd, 'fatu')
                fundo_v = get_val(sd, 'fundo')
                din_v   = get_val(sd, 'din')
                saldo_caixa = fundo_v + din_v - t_san
                st.metric("Total Recebido", fmt_brl(t_rec))
                st.metric("Saldo Final Caixa", fmt_brl(saldo_caixa))
                if fatu_v > 0:
                    pct = t_rec / fatu_v * 100
                    st.metric("% Recebido", f"{pct:.1f}%")

            with c4:
                st.markdown("**🪙 Sangrias**")
                sd['agua']           = ni("💧 Água",    f"{store_id}_agua",    sd['agua'])
                sd['mercado']        = ni("🛒 Mercado", f"{store_id}_mkt",     sd['mercado'])
                sd['cafe']           = ni("☕ Café",    f"{store_id}_cafe",    sd['cafe'])
                sd['pedagio']        = ni("🚗 Pedágio", f"{store_id}_ped",     sd['pedagio'])
                sd['farmacia']       = ni("💊 Farmácia",f"{store_id}_farm",    sd['farmacia'])
                sd['bolo']           = ni("🎂 Bolo",    f"{store_id}_bolo",    sd['bolo'])
                sd['banco']          = ni("🏦 Banco",   f"{store_id}_banco",   sd['banco'])
                sd['outros_sangria'] = ni("➕ Outros",  f"{store_id}_outsan",  sd['outros_sangria'])
                st.metric("Total Sangrias", fmt_brl(t_san))

# ──────────────────────────────────────────────────────────────
# TAB 2 — DASHBOARD
# ──────────────────────────────────────────────────────────────
with tab2:
    st.header(f"📊 Dashboard Gerencial — {sel_date.strftime('%d/%m/%Y')}")

    # ── KPIs consolidados ──
    tot_fatu   = sum(get_val(st.session_state.data[s], 'fatu')   for s in STORES)
    tot_pecas  = sum(get_val(st.session_state.data[s], 'pecas')  for s in STORES)
    tot_apagar = sum(get_val(st.session_state.data[s], 'apagar') for s in STORES)
    tot_din    = sum(get_val(st.session_state.data[s], 'din')    for s in STORES)
    tot_cc     = sum(get_val(st.session_state.data[s], 'cc')     for s in STORES)
    tot_cd     = sum(get_val(st.session_state.data[s], 'cd')     for s in STORES)
    tot_dep    = sum(get_val(st.session_state.data[s], 'dep')    for s in STORES)
    tot_out    = sum(get_val(st.session_state.data[s], 'outros') for s in STORES)
    tot_rec    = tot_din + tot_cc + tot_cd + tot_dep + tot_out
    ticket_med = tot_fatu / tot_pecas if tot_pecas > 0 else 0
    pct_rec    = tot_rec / tot_fatu * 100 if tot_fatu > 0 else 0

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("🏭 Faturamento Total", fmt_brl(tot_fatu))
    k2.metric("👕 Total Peças", f"{int(tot_pecas):,}".replace(',', '.'))
    k3.metric("💰 Recebido Hoje", fmt_brl(tot_rec),
              delta=f"{pct_rec:.1f}% do faturado" if tot_fatu > 0 else None)
    k4.metric("⏳ A Receber (Dia)", fmt_brl(tot_apagar))
    k5.metric("🎫 Ticket Médio", fmt_brl(ticket_med))

    st.divider()

    # ── Gráficos do dia ──
    chart_rows = []
    for s in STORES:
        sd = st.session_state.data[s]
        chart_rows.append({
            'Loja': STORES[s].split(' / ')[0],
            'Faturamento': get_val(sd, 'fatu'),
            'Recebido': total_recebido(sd),
            'A Pagar': get_val(sd, 'apagar'),
            'Peças': get_val(sd, 'pecas'),
        })
    df_day = pd.DataFrame(chart_rows)

    col_a, col_b = st.columns([3, 2])

    with col_a:
        fig_bar = px.bar(
            df_day.melt(id_vars='Loja', value_vars=['Faturamento', 'Recebido', 'A Pagar']),
            x='Loja', y='value', color='variable', barmode='group',
            title='Faturamento × Recebido × A Pagar por Loja',
            labels={'value': 'R$', 'variable': ''},
            color_discrete_map={'Faturamento': '#3498db', 'Recebido': '#2ecc71', 'A Pagar': '#e74c3c'},
        )
        fig_bar.update_layout(height=380, legend=dict(orientation='h', y=1.12))
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_b:
        pay_labels = ['Dinheiro', 'Crédito 🔵', 'Débito 🔵', 'Dep/PIX 🟠', 'Outros']
        pay_vals   = [tot_din, tot_cc, tot_cd, tot_dep, tot_out]
        filt = [(l, v) for l, v in zip(pay_labels, pay_vals) if v > 0]
        if filt:
            labels_f, vals_f = zip(*filt)
            fig_pie = go.Figure(go.Pie(
                labels=list(labels_f), values=list(vals_f), hole=0.42,
                marker_colors=['#2ecc71', '#3498db', '#9b59b6', '#e67e22', '#95a5a6'],
            ))
            fig_pie.update_layout(title='Formas de Pagamento', height=380,
                                  legend=dict(orientation='h', y=-0.1))
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("Preencha os recebimentos para ver o gráfico de pagamentos.")

    # ── Tabela resumo ──
    st.subheader("📋 Resumo por Loja")
    tbl = []
    for s in STORES:
        sd = st.session_state.data[s]
        tbl.append({
            'Loja': STORES[s],
            'Peças': int(get_val(sd, 'pecas')),
            'Fat. Líquido': fmt_brl(get_val(sd, 'fatu')),
            'A Pagar Dia': fmt_brl(get_val(sd, 'apagar')),
            'Dinheiro': fmt_brl(get_val(sd, 'din')),
            'Crédito 🔵': fmt_brl(get_val(sd, 'cc')),
            'Débito 🔵': fmt_brl(get_val(sd, 'cd')),
            'Dep/PIX 🟠': fmt_brl(get_val(sd, 'dep')),
            'Leitura X': fmt_brl(get_val(sd, 'leitura')),
            'Saldo Caixa': fmt_brl(get_val(sd, 'fundo') + get_val(sd, 'din') - total_sangria(sd)),
        })
    st.dataframe(pd.DataFrame(tbl).set_index('Loja'), use_container_width=True)

    # ── Histórico (se template carregado) ──
    if st.session_state.hist_wb:
        st.divider()
        st.subheader("📈 Tendência — Últimos 30 dias")
        hist_dfs = []
        for s in STORES:
            df_h = read_historical_excel(st.session_state.hist_wb, s)
            if not df_h.empty:
                df_h['loja_label'] = STORES[s].split(' / ')[0]
                hist_dfs.append(df_h)

        if hist_dfs:
            df_all = pd.concat(hist_dfs)
            df_all = df_all[df_all['data'] <= sel_date]
            cutoff = sel_date - timedelta(days=30)
            df_30 = df_all[df_all['data'] >= cutoff]

            if not df_30.empty:
                fig_line = px.line(
                    df_30, x='data', y='fatu', color='loja_label',
                    title='Faturamento diário por loja (últimos 30 dias)',
                    labels={'data': 'Data', 'fatu': 'Faturamento (R$)', 'loja_label': 'Loja'},
                    color_discrete_map={
                        'West Side': '#3498db', 'West Zone': '#e74c3c',
                        'West Place': '#2ecc71', 'West Station': '#f39c12'
                    },
                )
                fig_line.update_layout(height=350)
                st.plotly_chart(fig_line, use_container_width=True)

                # Consolidado semanal
                df_30['semana'] = df_30['data'].apply(lambda d: d.isocalendar()[1])
                df_week = df_30.groupby(['semana', 'loja_label'])['fatu'].sum().reset_index()
                fig_week = px.bar(
                    df_week, x='semana', y='fatu', color='loja_label', barmode='stack',
                    title='Faturamento semanal consolidado',
                    labels={'semana': 'Semana do ano', 'fatu': 'R$', 'loja_label': 'Loja'},
                )
                fig_week.update_layout(height=320)
                st.plotly_chart(fig_week, use_container_width=True)
        else:
            st.info("Sem dados históricos preenchidos no template carregado.")

# ──────────────────────────────────────────────────────────────
# TAB 3 — EXPORTAR
# ──────────────────────────────────────────────────────────────
with tab3:
    st.header("💾 Exportar Fechamento")

    st.subheader("1. Preencher o template Excel")
    template_file = st.file_uploader(
        "Carregue o arquivo 'Fechamento de Caixa - 2026.xlsx'",
        type=['xlsx'], key='tpl_upload'
    )

    if template_file:
        st.success("✅ Template carregado. Clique abaixo para gerar o arquivo preenchido.")
        if st.button("📥 Gerar Excel Preenchido", type="primary"):
            # Calcula linha da data selecionada
            ref_date = date(2026, 1, 1)
            row_num = (sel_date - ref_date).days + 2

            try:
                wb = load_workbook(template_file)
                sheet_map = {'SIDE': 'SIDE', 'ZONE': 'ZONE', 'PLACE': 'PLACE', 'STATION': 'STATION'}

                for store_id, sheet_name in sheet_map.items():
                    if sheet_name not in wb.sheetnames:
                        continue
                    ws = wb[sheet_name]
                    sd = st.session_state.data[store_id]

                    field_col = {
                        'pecas': 4, 'servicos': 5, 'fatu': 6, 'apagar': 7,
                        'din': 10, 'cheque': 11, 'cc': 12, 'cd': 13,
                        'dep': 14, 'ecom': 15, 'outros': 16, 'leitura': 18,
                        'agua': 21, 'mercado': 22, 'cafe': 23, 'pedagio': 24,
                        'farmacia': 25, 'bolo': 26, 'banco': 27, 'outros_sangria': 28,
                        'fundo': 29,
                    }
                    for field, col_idx in field_col.items():
                        val = sd.get(field)
                        if val is not None and float(val) != 0:
                            ws.cell(row=row_num, column=col_idx).value = float(val)

                    # Corrige fórmula AD no SIDE (template tem +0-J em vez de +J)
                    if store_id == 'SIDE':
                        ws.cell(row=row_num, column=30).value = f'=AC{row_num}+J{row_num}-T{row_num}'

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                st.download_button(
                    label="⬇️ Baixar Excel",
                    data=buf,
                    file_name=f"Fechamento_{sel_date.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success(f"✅ Linha {row_num} preenchida para {sel_date.strftime('%d/%m/%Y')}")
            except Exception as e:
                st.error(f"Erro ao gerar Excel: {e}")

    st.divider()
    st.subheader("2. Exportar como CSV")
    rows_csv = []
    for s in STORES:
        sd = st.session_state.data[s]
        row = {'data': sel_date.isoformat(), 'loja': STORES[s]}
        for f in ['pecas', 'servicos', 'fatu', 'apagar', 'din', 'cheque', 'cc', 'cd',
                  'dep', 'ecom', 'outros', 'leitura', 'fundo',
                  'agua', 'mercado', 'cafe', 'pedagio', 'farmacia', 'bolo', 'banco', 'outros_sangria']:
            row[f] = sd.get(f) or ''
        rows_csv.append(row)
    csv_str = pd.DataFrame(rows_csv).to_csv(index=False, sep=';', decimal=',')
    st.download_button(
        label="⬇️ Baixar CSV",
        data=csv_str.encode('utf-8-sig'),
        file_name=f"Fechamento_{sel_date.strftime('%Y%m%d')}.csv",
        mime='text/csv'
    )

    st.divider()
    st.subheader("3. Limpar sessão")
    if st.button("🗑️ Limpar todos os dados", type="secondary"):
        st.session_state.data = {s: empty_store() for s in STORES}
        st.rerun()
