import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────
#  Configuration des colonnes
# ──────────────────────────────────────────────
COLUMNS_BY_TYPE = {
    "Facturation Achat": {
        "extract_only": False,
        "columns": ["Date", "Products", "Quantity ordered", "Price", "Order number"],
        "rename": {
            "Price":        "Prix Achat HT",
            "Order number": "N° Facture",
        },
    },
    "Facturation Client": {
        "extract_only": True,
        "columns": [
            "Submit date", "Object code", "Object fullname", "Document number",
            "Delivery date", "Transport fees", "Product code", "Product",
            "Product type", "Type of article", "Quantity", "Sale unit code",
            "Position brutto value", "Position VAT rate", "PSA value", "PSA",
        ],
        "rename": {
            "Submit date":           "Date/Heure Emission",
            "Object code":           "Code Client",
            "Object fullname":       "Nom Client",
            "Document number":       "N° Facture",
            "Delivery date":         "Date Livraison",
            "Transport fees":        "Frais de port",
            "Product code":          "Code Produit",
            "Product":               "Désignation Produit",
            "Product type":          "Type Produit",
            "Quantity":              "Quantité Commandée",
            "Position brutto value": "Valeur Brutto",
            "Position VAT rate":     "Taux de TVA",
            "PSA value":             "Valeur PSA",
            "PSA":                   "PSA",
        },
        "extra_columns": ["N° Camion", "Chauffeur", "Vendeur", "Tournée"],
    },
}

INDEX_SHEET_NAME = "Données"
TRACE_COLS       = ["Fichier source", "Date d'import"]

# ──────────────────────────────────────────────
#  Helpers de Style
# ──────────────────────────────────────────────
def _header_style(ws, nb_cols, row=1):
    for col_idx in range(1, nb_cols + 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill      = PatternFill("solid", start_color="2E4057")
        cell.alignment = Alignment(horizontal="center", vertical="center")

def _border():
    thin = Side(style="thin", color="BFBFBF")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def apply_auto_width(ws):
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row=row, column=col_idx).value
            if val:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

# ──────────────────────────────────────────────
#  Rapport Global
# ──────────────────────────────────────────────
def build_report_sheet(wb, all_results):
    ws = wb.create_sheet("Rapport Global", 0)
    header_fill = PatternFill("solid", start_color="2E4057")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    
    headers = ["Fichier", "Feuille", "Cols", "Lignes Tot.", "Uniques", "Doublons", "Mode"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    current_row = 2
    for file_res in all_results:
        filename = file_res["filename"]
        for r in file_res["sheets"]:
            ws.cell(row=current_row, column=1, value=filename)
            ws.cell(row=current_row, column=2, value=r["sheet_name"])
            ws.cell(row=current_row, column=3, value=r["total_columns"])
            ws.cell(row=current_row, column=4, value=r["total_rows"])
            ws.cell(row=current_row, column=5, value=r["unique_rows"])
            ws.cell(row=current_row, column=6, value=r["duplicate_rows"])
            ws.cell(row=current_row, column=7, value="Extraction" if r["extract_only"] else "Dédoublonnage")
            for col in range(1, 8):
                ws.cell(row=current_row, column=col).border = _border()
            current_row += 1

    apply_auto_width(ws)
    ws.sheet_view.showGridLines = False

# ──────────────────────────────────────────────
#  Gestion de l'Index Cumulatif
# ──────────────────────────────────────────────
def update_index_streamlit(all_results, file_type, existing_index_file=None):
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    frames = []
    for file_res in all_results:
        source_filename = file_res["filename"]
        for r in file_res["sheets"]:
            df = r["df_result"].copy()
            df.insert(0, TRACE_COLS[0], source_filename)
            df.insert(1, TRACE_COLS[1], now_str)
            frames.append(df)
    
    if not frames: return None
    df_new = pd.concat(frames, ignore_index=True)
    
    if existing_index_file:
        wb_idx = load_workbook(existing_index_file)
        if INDEX_SHEET_NAME in wb_idx.sheetnames:
            ws_idx = wb_idx[INDEX_SHEET_NAME]
            existing_headers = [ws_idx.cell(row=1, column=c).value for c in range(1, ws_idx.max_column + 1) if ws_idx.cell(row=1, column=c).value]
            for col in existing_headers:
                if col not in df_new.columns: df_new[col] = ""
            df_new = df_new[[c for c in existing_headers if c in df_new.columns]]
            next_row = ws_idx.max_row + 1
        else:
            ws_idx = wb_idx.create_sheet(INDEX_SHEET_NAME)
            next_row = 1
    else:
        wb_idx = Workbook()
        ws_idx = wb_idx.active
        ws_idx.title = INDEX_SHEET_NAME
        next_row = 1

    border = _border()
    alt_fill = PatternFill("solid", start_color="EAF0FB")
    white_fill = PatternFill("solid", start_color="FFFFFF")
    trace_fill = PatternFill("solid", start_color="E8F5E9")

    if next_row == 1:
        for col_idx, col_name in enumerate(df_new.columns, start=1):
            ws_idx.cell(row=1, column=col_idx, value=col_name)
        _header_style(ws_idx, len(df_new.columns), row=1)
        next_row = 2

    for row_offset, row_data in enumerate(df_new.itertuples(index=False)):
        current_row = next_row + row_offset
        fill = alt_fill if current_row % 2 == 0 else white_fill
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws_idx.cell(row=current_row, column=col_idx, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
            if col_idx <= len(TRACE_COLS):
                cell.font = Font(name="Arial", size=9, italic=True, color="555555")
                cell.fill = trace_fill
            else:
                cell.font = Font(name="Arial", size=10)
                cell.fill = fill

    apply_auto_width(ws_idx)
    ws_idx.sheet_view.showGridLines = False
    output = io.BytesIO()
    wb_idx.save(output)
    return output.getvalue()

# ──────────────────────────────────────────────
#  Logique de traitement
# ──────────────────────────────────────────────
def filter_columns(df, sheet_name, file_type):
    if file_type is None or file_type not in COLUMNS_BY_TYPE:
        return df, [], list(df.columns)
    config = COLUMNS_BY_TYPE[file_type]
    expected = [c for c in config["columns"]]
    rename_map = config.get("rename", {})
    available = [c for c in expected if c in df.columns]
    missing = [c for c in expected if c not in df.columns]
    df_out = df[available].rename(columns=rename_map)
    for extra_col in config.get("extra_columns", []):
        df_out[extra_col] = ""
    return df_out, missing, list(df_out.columns)

def process_multiple_files(uploaded_files, file_type, index_file=None):
    all_results = []
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for uploaded_file in uploaded_files:
            uploaded_file.seek(0)
            excel_data = pd.read_excel(uploaded_file, sheet_name=None)
            file_sheets_res = []
            for sheet_name, df in excel_data.items():
                df_filtered, missing_cols, exported_cols = filter_columns(df, sheet_name, file_type)
                extract_only = COLUMNS_BY_TYPE[file_type].get("extract_only", False) if file_type in COLUMNS_BY_TYPE else False
                
                if extract_only:
                    df_result = df_filtered.reset_index(drop=True)
                    unique_rows, duplicate_rows = len(df), 0
                else:
                    df_result = df_filtered.drop_duplicates()
                    unique_rows = len(df_result)
                    duplicate_rows = len(df) - unique_rows
                
                sheet_export_name = f"{uploaded_file.name[:15]}_{sheet_name[:10]}"
                df_result.to_excel(writer, sheet_name=sheet_export_name, index=False)
                
                file_sheets_res.append({
                    "sheet_name": sheet_name, "total_columns": len(df.columns),
                    "total_rows": len(df), "unique_rows": unique_rows,
                    "duplicate_rows": duplicate_rows, "extract_only": extract_only,
                    "df_result": df_result
                })
            all_results.append({"filename": uploaded_file.name, "sheets": file_sheets_res})

    output.seek(0)
    wb = load_workbook(output)
    build_report_sheet(wb, all_results)
    
    for sheet_name in wb.sheetnames:
        if sheet_name != "Rapport Global":
            ws = wb[sheet_name]
            _header_style(ws, ws.max_column)
            apply_auto_width(ws)

    final_output = io.BytesIO()
    wb.save(final_output)
    
    index_data = update_index_streamlit(all_results, file_type, index_file)
    
    return final_output.getvalue(), index_data, all_results


# ══════════════════════════════════════════════
#  CSS GLOBAL — Design système
# ══════════════════════════════════════════════
st.set_page_config(
    page_title="Excel Batch Processor Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
/* ── Imports ── */
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

/* ── Variables ── */
:root {
    --navy:      #1B2A4A;
    --navy-mid:  #2E4470;
    --blue:      #3A6FD8;
    --blue-light:#5B8FEE;
    --accent:    #00C49A;
    --warn:      #F5A623;
    --danger:    #E05252;
    --bg:        #F4F6FB;
    --surface:   #FFFFFF;
    --border:    #DDE3F0;
    --text:      #1B2A4A;
    --muted:     #7A889E;
    --radius:    12px;
    --shadow:    0 2px 12px rgba(27,42,74,0.08);
    --shadow-md: 0 6px 24px rgba(27,42,74,0.12);
}

/* ── Base ── */
html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
    color: var(--text);
}
.stApp { background: var(--bg); }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: var(--navy) !important;
    border-right: 1px solid var(--navy-mid);
}
section[data-testid="stSidebar"] * {
    color: #E8EDF8 !important;
}
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stFileUploader label {
    color: #A8B8D8 !important;
    font-size: 0.78rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.07em !important;
    text-transform: uppercase !important;
}
section[data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div {
    background: rgba(255,255,255,0.08) !important;
    border: 1px solid rgba(255,255,255,0.15) !important;
    border-radius: 8px !important;
    color: #fff !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background: rgba(255,255,255,0.05) !important;
    border: 1px dashed rgba(255,255,255,0.2) !important;
    border-radius: 8px !important;
}

/* ── Step cards ── */
.step-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.4rem 1.6rem;
    margin-bottom: 1.2rem;
    box-shadow: var(--shadow);
    transition: box-shadow 0.2s;
}
.step-card:hover { box-shadow: var(--shadow-md); }

.step-header {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin-bottom: 0.6rem;
}
.step-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 28px; height: 28px;
    background: var(--blue);
    color: #fff;
    border-radius: 50%;
    font-size: 0.78rem;
    font-weight: 700;
    flex-shrink: 0;
}
.step-badge.done { background: var(--accent); }
.step-title {
    font-weight: 700;
    font-size: 1rem;
    color: var(--navy);
}
.step-desc {
    font-size: 0.82rem;
    color: var(--muted);
    margin-left: 2.75rem;
    line-height: 1.5;
}

/* ── Mode pill ── */
.mode-pill {
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    padding: 0.3rem 0.85rem;
    border-radius: 999px;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.03em;
}
.mode-pill.extract  { background: #EAF9F5; color: #00A882; border: 1px solid #B2EDDF; }
.mode-pill.dedupe   { background: #EEF3FF; color: #3A6FD8; border: 1px solid #BDD0F8; }
.mode-pill.standard { background: #F3F5FB; color: #7A889E; border: 1px solid #DDE3F0; }

/* ── Info callout ── */
.info-box {
    background: #EEF3FF;
    border-left: 3px solid var(--blue);
    border-radius: 0 8px 8px 0;
    padding: 0.75rem 1rem;
    font-size: 0.83rem;
    color: #2E4470;
    margin: 0.5rem 0;
}
.warn-box {
    background: #FFF8EE;
    border-left: 3px solid var(--warn);
    border-radius: 0 8px 8px 0;
    padding: 0.75rem 1rem;
    font-size: 0.83rem;
    color: #7A4E00;
    margin: 0.5rem 0;
}

/* ── Metric cards ── */
.metric-row { display: flex; gap: 0.75rem; flex-wrap: wrap; margin: 0.75rem 0; }
.metric-card {
    flex: 1;
    min-width: 110px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 0.85rem 1rem;
    text-align: center;
}
.metric-num {
    font-size: 1.6rem;
    font-weight: 700;
    color: var(--navy);
    font-family: 'DM Mono', monospace;
    line-height: 1.1;
}
.metric-num.accent { color: var(--accent); }
.metric-num.warn   { color: var(--warn); }
.metric-label {
    font-size: 0.72rem;
    color: var(--muted);
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-top: 0.2rem;
}

/* ── File tag ── */
.file-tag {
    display: inline-flex;
    align-items: center;
    gap: 0.35rem;
    background: #EEF3FF;
    color: #2E4470;
    border: 1px solid #BDD0F8;
    border-radius: 6px;
    padding: 0.22rem 0.65rem;
    font-size: 0.78rem;
    font-weight: 500;
    margin: 0.15rem;
    font-family: 'DM Mono', monospace;
}

/* ── Result table ── */
.result-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.84rem;
    margin-top: 0.5rem;
}
.result-table th {
    background: var(--navy);
    color: #fff;
    padding: 0.6rem 0.9rem;
    text-align: left;
    font-weight: 600;
    font-size: 0.75rem;
    letter-spacing: 0.05em;
    text-transform: uppercase;
}
.result-table th:first-child { border-radius: 8px 0 0 0; }
.result-table th:last-child  { border-radius: 0 8px 0 0; }
.result-table td {
    padding: 0.55rem 0.9rem;
    border-bottom: 1px solid var(--border);
    color: var(--text);
}
.result-table tr:hover td { background: #F4F6FB; }
.result-table tr:last-child td { border-bottom: none; }
.badge-ok  { color: var(--accent); font-weight: 600; }
.badge-dup { color: var(--warn);   font-weight: 600; }

/* ── Download block ── */
.dl-card {
    background: linear-gradient(135deg, #1B2A4A 0%, #2E4470 100%);
    border-radius: var(--radius);
    padding: 1.4rem 1.6rem;
    color: #fff;
}
.dl-title {
    font-size: 0.9rem;
    font-weight: 700;
    margin-bottom: 0.25rem;
    color: #fff;
}
.dl-sub {
    font-size: 0.78rem;
    color: #A8B8D8;
    margin-bottom: 1rem;
}

/* ── Buttons override ── */
.stButton > button {
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    transition: all 0.18s !important;
}
.stDownloadButton > button {
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
}

/* ── Section label ── */
.section-label {
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--muted);
    margin-bottom: 0.5rem;
}

/* ── Divider ── */
hr.soft { border: none; border-top: 1px solid var(--border); margin: 1.2rem 0; }

/* ── Success banner ── */
.success-banner {
    background: linear-gradient(135deg, #00C49A22, #00C49A11);
    border: 1px solid #00C49A55;
    border-radius: var(--radius);
    padding: 1rem 1.4rem;
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin-bottom: 1.2rem;
}
.success-icon { font-size: 1.6rem; }
.success-text { font-weight: 600; color: #007A62; font-size: 0.95rem; }
.success-sub  { font-size: 0.8rem; color: #009E80; margin-top: 0.1rem; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  Session State
# ══════════════════════════════════════════════
if 'results'           not in st.session_state: st.session_state.results           = None
if 'processed_data'    not in st.session_state: st.session_state.processed_data    = None
if 'index_data'        not in st.session_state: st.session_state.index_data        = None
if 'current_file_type' not in st.session_state: st.session_state.current_file_type = None

# ══════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style="padding:0.5rem 0 1.4rem 0;">
        <div style="font-size:1.35rem;font-weight:800;color:#fff;letter-spacing:-0.02em;">
            📊 Batch Processor
        </div>
        <div style="font-size:0.78rem;color:#7A98CC;margin-top:0.2rem;">
            Traitement Excel professionnel
        </div>
    </div>
    <hr style="border:none;border-top:1px solid rgba(255,255,255,0.1);margin-bottom:1.2rem;">
    """, unsafe_allow_html=True)

    # ── Étape 1 ──
    st.markdown("""
    <div class="section-label" style="color:#7A98CC;">Étape 1 · Mode de traitement</div>
    """, unsafe_allow_html=True)
    
    file_type = st.selectbox(
        "Type de fichier",
        options=[None] + list(COLUMNS_BY_TYPE.keys()),
        format_func=lambda x: "⚙️  Standard (toutes colonnes)" if x is None else (
            "🛒  Facturation Achat" if x == "Facturation Achat" else "👤  Facturation Client"
        ),
        label_visibility="collapsed",
        help="Choisissez le profil qui correspond à vos fichiers source."
    )

    # Explication du mode sélectionné
    if file_type is None:
        st.markdown('<div class="mode-pill standard">⚙️ Standard — toutes colonnes conservées</div>', unsafe_allow_html=True)
        st.markdown('<div class="info-box" style="margin-top:0.6rem;">Les doublons seront supprimés automatiquement. Toutes les colonnes du fichier source sont exportées.</div>', unsafe_allow_html=True)
    elif file_type == "Facturation Achat":
        st.markdown('<div class="mode-pill dedupe">🔄 Mode Dédoublonnage actif</div>', unsafe_allow_html=True)
        st.markdown('<div class="info-box" style="margin-top:0.6rem;">5 colonnes ciblées · Les lignes identiques sont supprimées.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="mode-pill extract">📤 Mode Extraction actif</div>', unsafe_allow_html=True)
        st.markdown('<div class="info-box" style="margin-top:0.6rem;">16 colonnes ciblées · Toutes les lignes sont conservées.</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Étape 2 ──
    st.markdown('<div class="section-label" style="color:#7A98CC;">Étape 2 · Index cumulatif (optionnel)</div>', unsafe_allow_html=True)
    index_file = st.file_uploader(
        "Charger l'index existant",
        type=["xlsx"],
        label_visibility="collapsed",
        help="Si vous avez déjà un fichier index d'imports précédents, chargez-le ici. Les nouvelles données y seront ajoutées sans écraser l'existant."
    )
    if index_file:
        st.markdown(f'<div style="margin-top:0.4rem;"><span class="file-tag">📎 {index_file.name}</span></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div style="font-size:0.76rem;color:#7A98CC;margin-top:0.3rem;">💡 Sans index, un nouveau fichier sera créé.</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<hr style="border:none;border-top:1px solid rgba(255,255,255,0.1);">', unsafe_allow_html=True)

    if st.button("🔄 Réinitialiser l'application", use_container_width=True):
        st.session_state.results           = None
        st.session_state.processed_data    = None
        st.session_state.index_data        = None
        st.session_state.current_file_type = None
        st.rerun()

# ══════════════════════════════════════════════
#  CORPS PRINCIPAL
# ══════════════════════════════════════════════

# ── En-tête ──
st.markdown("""
<div style="margin-bottom:1.8rem;">
    <h1 style="font-size:1.75rem;font-weight:800;color:#1B2A4A;margin-bottom:0.2rem;letter-spacing:-0.02em;">
        Traitement groupé de fichiers Excel
    </h1>
    <p style="font-size:0.9rem;color:#7A889E;margin:0;">
        Chargez un ou plusieurs fichiers, configurez le mode dans le panneau gauche, puis lancez le traitement en un clic.
    </p>
</div>
""", unsafe_allow_html=True)

# ── Étape 3 : Chargement des fichiers ──
st.markdown("""
<div class="step-card">
    <div class="step-header">
        <span class="step-badge">3</span>
        <span class="step-title">Charger vos fichiers Excel</span>
    </div>
    <div class="step-desc">Formats acceptés : <strong>.xlsx</strong> · Plusieurs fichiers simultanément · Chaque feuille est traitée séparément.</div>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Déposez vos fichiers ici ou cliquez pour parcourir",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help="Vous pouvez sélectionner plusieurs fichiers en maintenant Ctrl (Windows) ou ⌘ (Mac).",
)

# Aperçu des fichiers chargés
if uploaded_files:
    tags = "".join([f'<span class="file-tag">📄 {f.name}</span>' for f in uploaded_files])
    st.markdown(f"""
    <div style="margin:0.75rem 0 0.25rem 0;">
        <span style="font-size:0.78rem;font-weight:600;color:#7A889E;text-transform:uppercase;letter-spacing:0.06em;">
            {len(uploaded_files)} fichier(s) sélectionné(s)
        </span>
    </div>
    <div style="margin-bottom:0.5rem;">{tags}</div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Bouton de traitement ──
col_btn, col_info = st.columns([2, 3])
with col_btn:
    process_disabled = not bool(uploaded_files)
    launch = st.button(
        f"🚀  Lancer le traitement — {len(uploaded_files) if uploaded_files else 0} fichier(s)",
        use_container_width=True,
        disabled=process_disabled,
        type="primary",
    )
with col_info:
    if process_disabled:
        st.markdown('<div class="warn-box" style="margin:0;">⬅️ Chargez au moins un fichier Excel pour activer le traitement.</div>', unsafe_allow_html=True)
    else:
        mode_label = "Standard" if file_type is None else file_type
        st.markdown(f'<div class="info-box" style="margin:0;">✅ Prêt · Mode <strong>{mode_label}</strong> · Cliquez sur le bouton pour démarrer.</div>', unsafe_allow_html=True)

# ── Traitement ──
if launch:
    with st.spinner(f"Analyse de {len(uploaded_files)} fichier(s) en cours…"):
        try:
            processed_data, index_data, all_results = process_multiple_files(uploaded_files, file_type, index_file)
            st.session_state.processed_data    = processed_data
            st.session_state.index_data        = index_data
            st.session_state.results           = all_results
            st.session_state.current_file_type = file_type
        except Exception as e:
            st.error(f"❌ Une erreur est survenue : {e}")

# ══════════════════════════════════════════════
#  RÉSULTATS
# ══════════════════════════════════════════════
if st.session_state.results:
    st.markdown("<hr class='soft'>", unsafe_allow_html=True)

    # ── Bannière succès ──
    total_rows  = sum(s["total_rows"]     for f in st.session_state.results for s in f["sheets"])
    total_uniq  = sum(s["unique_rows"]    for f in st.session_state.results for s in f["sheets"])
    total_dups  = sum(s["duplicate_rows"] for f in st.session_state.results for s in f["sheets"])
    nb_files    = len(st.session_state.results)
    nb_sheets   = sum(len(f["sheets"]) for f in st.session_state.results)

    st.markdown(f"""
    <div class="success-banner">
        <div class="success-icon">✅</div>
        <div>
            <div class="success-text">Traitement terminé avec succès</div>
            <div class="success-sub">{nb_files} fichier(s) · {nb_sheets} feuille(s) analysée(s) · {total_rows:,} lignes traitées</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Métriques globales ──
    st.markdown('<div class="section-label">Vue d\'ensemble</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="metric-row">
        <div class="metric-card">
            <div class="metric-num">{nb_files}</div>
            <div class="metric-label">Fichiers</div>
        </div>
        <div class="metric-card">
            <div class="metric-num">{nb_sheets}</div>
            <div class="metric-label">Feuilles</div>
        </div>
        <div class="metric-card">
            <div class="metric-num">{total_rows:,}</div>
            <div class="metric-label">Lignes totales</div>
        </div>
        <div class="metric-card">
            <div class="metric-num accent">{total_uniq:,}</div>
            <div class="metric-label">Lignes uniques</div>
        </div>
        <div class="metric-card">
            <div class="metric-num warn">{total_dups:,}</div>
            <div class="metric-label">Doublons supprimés</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Résumé détaillé + Téléchargements ──
    col_table, col_dl = st.columns([3, 2])

    with col_table:
        st.markdown('<div class="section-label">Détail par feuille</div>', unsafe_allow_html=True)

        # Construction du tableau HTML
        rows_html = ""
        for f_res in st.session_state.results:
            for s_res in f_res["sheets"]:
                mode_html = (
                    '<span class="mode-pill extract" style="font-size:0.72rem;padding:0.15rem 0.55rem;">📤 Extraction</span>'
                    if s_res["extract_only"]
                    else '<span class="mode-pill dedupe" style="font-size:0.72rem;padding:0.15rem 0.55rem;">🔄 Dédoublonnage</span>'
                )
                dup_class = "badge-dup" if s_res["duplicate_rows"] > 0 else "badge-ok"
                rows_html += f"""
                <tr>
                    <td><span class="file-tag" style="font-size:0.75rem;">{f_res['filename']}</span></td>
                    <td>{s_res['sheet_name']}</td>
                    <td style="text-align:center;font-family:'DM Mono',monospace;">{s_res['total_rows']:,}</td>
                    <td style="text-align:center;" class="badge-ok">{s_res['unique_rows']:,}</td>
                    <td style="text-align:center;" class="{dup_class}">{s_res['duplicate_rows']:,}</td>
                    <td>{mode_html}</td>
                </tr>"""

        st.markdown(f"""
        <div style="background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;box-shadow:var(--shadow);">
            <table class="result-table">
                <thead>
                    <tr>
                        <th>Fichier</th><th>Feuille</th>
                        <th style="text-align:center;">Total</th>
                        <th style="text-align:center;">Uniques</th>
                        <th style="text-align:center;">Doublons</th>
                        <th>Mode</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
        """, unsafe_allow_html=True)

    with col_dl:
        st.markdown('<div class="section-label">Télécharger les résultats</div>', unsafe_allow_html=True)
        
        type_suffix   = "Standard" if st.session_state.current_file_type is None else st.session_state.current_file_type.replace(" ", "_")
        export_name   = f"traitement_{type_suffix}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        index_name    = f"index_{type_suffix}.xlsx"
        mode_label    = st.session_state.current_file_type if st.session_state.current_file_type else "Standard"

        st.markdown("""
        <div class="dl-card">
            <div class="dl-title">📦 Fichier consolidé</div>
            <div class="dl-sub">Toutes les feuilles traitées + rapport global</div>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label="💾  Télécharger le fichier consolidé",
            data=st.session_state.processed_data,
            file_name=export_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        if st.session_state.index_data:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(f"""
            <div class="dl-card" style="background:linear-gradient(135deg,#005940 0%,#007A62 100%);">
                <div class="dl-title">📂 Index cumulatif — {mode_label}</div>
                <div class="dl-sub">Historique enrichi avec cet import</div>
            </div>
            """, unsafe_allow_html=True)

            st.download_button(
                label=f"📂  Télécharger l'index {mode_label}",
                data=st.session_state.index_data,
                file_name=index_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.markdown("""
        <div style="margin-top:1rem;background:#F4F6FB;border:1px solid #DDE3F0;border-radius:8px;padding:0.75rem 1rem;font-size:0.78rem;color:#7A889E;">
            💡 <strong>Astuce :</strong> Vos résultats restent disponibles jusqu'à la réinitialisation. Vous pouvez télécharger les fichiers à tout moment.
        </div>
        """, unsafe_allow_html=True)