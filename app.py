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
    
    # Générer l'index
    index_data = update_index_streamlit(all_results, file_type, index_file)
    
    return final_output.getvalue(), index_data, all_results

# ──────────────────────────────────────────────
#  Interface Streamlit avec Session State
# ──────────────────────────────────────────────
st.set_page_config(page_title="Excel Batch Processor Pro", page_icon="📊", layout="wide")

if 'results' not in st.session_state:
    st.session_state.results = None
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'index_data' not in st.session_state:
    st.session_state.index_data = None
if 'current_file_type' not in st.session_state:
    st.session_state.current_file_type = None

st.title("📊 Excel Batch Processor & Extractor")
st.markdown("Solution professionnelle pour le traitement par lots avec **mémorisation des résultats**.")

with st.sidebar:
    st.header("1. Configuration")
    file_type = st.selectbox("Type de traitement", options=[None] + list(COLUMNS_BY_TYPE.keys()), format_func=lambda x: "Standard" if x is None else x)
    
    st.header("2. Index Cumulatif")
    index_file = st.file_uploader("Charger l'index actuel (.xlsx)", type=["xlsx"])
    
    if st.button("🔄 Réinitialiser l'application"):
        st.session_state.results = None
        st.session_state.processed_data = None
        st.session_state.index_data = None
        st.session_state.current_file_type = None
        st.rerun()

st.subheader("📁 Charger vos fichiers Excel")
uploaded_files = st.file_uploader("Sélectionnez un ou plusieurs fichiers Excel", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and st.button("🚀 Lancer le traitement groupé", use_container_width=True):
    with st.spinner(f"Traitement de {len(uploaded_files)} fichiers en cours..."):
        try:
            processed_data, index_data, all_results = process_multiple_files(uploaded_files, file_type, index_file)
            
            st.session_state.processed_data = processed_data
            st.session_state.index_data = index_data
            st.session_state.results = all_results
            st.session_state.current_file_type = file_type
            st.success("✅ Traitement terminé !")
        except Exception as e:
            st.error(f"Erreur : {e}")

if st.session_state.results:
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("### 📋 Résumé Global")
        report_data = []
        for f_res in st.session_state.results:
            for s_res in f_res["sheets"]:
                report_data.append({
                    "Fichier": f_res["filename"], 
                    "Feuille": s_res["sheet_name"],
                    "Lignes": s_res["total_rows"], 
                    "Doublons": s_res["duplicate_rows"]
                })
        st.table(pd.DataFrame(report_data))
    
    with col2:
        st.markdown("### 📥 Téléchargements")
        # Nom de fichier dynamique pour le fichier consolidé
        type_suffix = "Standard" if st.session_state.current_file_type is None else st.session_state.current_file_type.replace(" ", "_")
        
        st.download_button(
            label="💾 Télécharger le fichier consolidé",
            data=st.session_state.processed_data,
            file_name=f"traitement_{type_suffix}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        if st.session_state.index_data:
            # Nom de fichier dynamique pour l'index
            index_filename = f"index_{type_suffix}.xlsx"
            st.download_button(
                label=f"📂 Télécharger l'index {st.session_state.current_file_type if st.session_state.current_file_type else 'Standard'}", 
                data=st.session_state.index_data, 
                file_name=index_filename, 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True
            )
        st.info("💡 Vos résultats sont mémorisés. Vous pouvez cliquer sur les boutons de téléchargement sans perdre l'affichage.")
