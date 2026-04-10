import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────
#  Configuration des colonnes (Mise à jour : MOMO)
# ──────────────────────────────────────────────
COLUMNS_BY_TYPE = {
    "Facturation Achat": {
        "extract_only": False,
        "description": "📦 **Extraction + Déduplication** : Garde uniquement les lignes uniques pour vos achats.",
        "columns": ["Date", "Products", "Quantity ordered", "Price", "Order number"],
        "rename": {
            "Price":        "Prix Achat HT",
            "Order number": "N° Facture",
        },
    },
    "Facturation Client": {
        "extract_only": True,
        "description": "👥 **Extraction seule** : Prépare vos données clients sans supprimer de lignes.",
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
    "Momo": {
        "extract_only": True,
        "description": "📱 **Traitement Momo (CSV/Excel)** : Extraction et renommage spécifique des flux mobiles.",
        "columns": ["Id", "Date", "Status", "Type", "From", "To name", "Amount", "Balance"],
        "rename": {
            "Id":      "N° Identification",
            "From":    "Provenance",
            "To name": "To handler name",
        },
        "extra_columns": ["Vendeur", "Compte", "Tournée"],
    }
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

def load_data(uploaded_file):
    """Charge un fichier CSV ou Excel en dictionnaire de DataFrames."""
    filename = uploaded_file.name
    if filename.endswith('.csv'):
        # Tenter plusieurs séparateurs courants pour le CSV
        try:
            df = pd.read_csv(uploaded_file, sep=None, engine='python')
        except:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=',')
        return {"Données_CSV": df}
    else:
        return pd.read_excel(uploaded_file, sheet_name=None)

def process_multiple_files(uploaded_files, file_type, index_file=None):
    all_results = []
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for uploaded_file in uploaded_files:
            uploaded_file.seek(0)
            excel_data = load_data(uploaded_file)
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
                
                sheet_export_name = f"{uploaded_file.name[:15]}_{sheet_name[:10]}".replace('.', '_')
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

# ──────────────────────────────────────────────
#  Interface Streamlit (UX Optimisée)
# ──────────────────────────────────────────────
st.set_page_config(page_title="Excel & CSV Processor Pro", page_icon="📈", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { border-radius: 8px; height: 3em; font-weight: bold; }
    .step-box { background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #2E4057; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .step-title { color: #2E4057; font-size: 1.2em; font-weight: bold; margin-bottom: 10px; }
    </style>
    """, unsafe_allow_html=True)

if 'results' not in st.session_state: st.session_state.results = None
if 'processed_data' not in st.session_state: st.session_state.processed_data = None
if 'index_data' not in st.session_state: st.session_state.index_data = None
if 'current_file_type' not in st.session_state: st.session_state.current_file_type = None

st.title("📈 Excel & CSV Processor Pro")
st.markdown("Solution professionnelle pour le traitement par lots : **Achat, Client et Momo**.")

with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/microsoft-excel-2019.png", width=80)
    st.header("🛠️ Options")
    if st.button("🔄 Réinitialiser l'outil", use_container_width=True):
        st.session_state.results = None
        st.session_state.processed_data = None
        st.session_state.index_data = None
        st.session_state.current_file_type = None
        st.rerun()

st.markdown('<div class="step-box"><div class="step-title">1️⃣ Configuration du Traitement</div>', unsafe_allow_html=True)
col_cfg1, col_cfg2 = st.columns(2)
with col_cfg1:
    file_type = st.selectbox(
        "Type de données à traiter", 
        options=[None] + list(COLUMNS_BY_TYPE.keys()), 
        format_func=lambda x: "✨ Traitement Standard" if x is None else x
    )
    if file_type:
        st.info(COLUMNS_BY_TYPE[file_type]["description"])

with col_cfg2:
    index_file = st.file_uploader("📂 Index existant (Optionnel)", type=["xlsx"])
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="step-box"><div class="step-title">2️⃣ Chargement des Fichiers (Excel ou CSV)</div>', unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "Glissez vos fichiers ici (.xlsx, .csv)", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True
)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_files:
    if st.button("🚀 Lancer le Traitement Groupé", use_container_width=True, type="primary"):
        with st.spinner("Analyse en cours..."):
            try:
                processed_data, index_data, all_results = process_multiple_files(uploaded_files, file_type, index_file)
                st.session_state.processed_data = processed_data
                st.session_state.index_data = index_data
                st.session_state.results = all_results
                st.session_state.current_file_type = file_type
                st.balloons()
            except Exception as e:
                st.error(f"⚠️ Erreur : {e}")

if st.session_state.results:
    st.markdown('<div class="step-box"><div class="step-title">3️⃣ Résultats & Téléchargements</div>', unsafe_allow_html=True)
    res_col1, res_col2 = st.columns([2, 1])
    with res_col1:
        st.markdown("### 📋 Rapport d'analyse")
        report_data = []
        for f_res in st.session_state.results:
            for s_res in f_res["sheets"]:
                report_data.append({
                    "Fichier Source": f_res["filename"], "Feuille": s_res["sheet_name"],
                    "Lignes": s_res["total_rows"], "Doublons": f"❌ {s_res['duplicate_rows']}" if s_res['duplicate_rows'] > 0 else "✅ 0"
                })
        st.dataframe(pd.DataFrame(report_data), use_container_width=True, hide_index=True)
    with res_col2:
        st.markdown("### 📥 Téléchargements")
        type_suffix = "Standard" if st.session_state.current_file_type is None else st.session_state.current_file_type.replace(" ", "_")
        st.download_button(
            label="💾 Télécharger le fichier consolidé",
            data=st.session_state.processed_data,
            file_name=f"traitement_{type_suffix}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        if st.session_state.index_data:
            st.download_button(
                label=f"📂 Télécharger l'index {st.session_state.current_file_type if st.session_state.current_file_type else 'Standard'}", 
                data=st.session_state.index_data, 
                file_name=f"index_{type_suffix}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True
            )
    st.markdown('</div>', unsafe_allow_html=True)
