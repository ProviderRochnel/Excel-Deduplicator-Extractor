import streamlit as st
import pandas as pd
import io
import os
import re
import numpy as np
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
try:
    from mega import Mega
except ImportError:
    Mega = None

# ──────────────────────────────────────────────
#  Configuration des colonnes & Métadonnées
# ──────────────────────────────────────────────
COLUMNS_BY_TYPE = {
    "Facturation Achat": {
        "extract_only": False,
        "icon": "📦",
        "description": "Optimisation des achats : extraction des données clés et suppression automatique des doublons de facturation.",
        "tooltip": "Idéal pour consolider vos bons de commande et éviter les doubles paiements.",
        "columns": ["Date", "Products", "Quantity ordered", "Price", "Order number"],
        "rename": {"Price": "Prix Achat HT", "Order number": "N° Facture"},
    },
    "Facturation Client": {
        "extract_only": True,
        "icon": "👥",
        "description": "Préparation des données clients : extraction exhaustive pour analyse de facturation et logistique.",
        "tooltip": "Préparez vos fichiers pour la livraison et le suivi commercial.",
        "columns": [
            "Submit date", "Object code", "Object fullname", "Document number",
            "Delivery date", "Transport fees", "Product code", "Product",
            "Product type", "Type of article", "Quantity", "Sale unit code",
            "Position brutto value", "Position VAT rate", "PSA value", "PSA",
        ],
        "rename": {
            "Submit date": "Date/Heure Emission", "Object code": "Code Client", "Object fullname": "Nom Client",
            "Document number": "N° Facture", "Delivery date": "Date Livraison", "Transport fees": "Frais de port",
            "Product code": "Code Produit", "Product": "Désignation Produit", "Product type": "Type Produit",
            "Quantity": "Quantité Commandée", "Position brutto value": "Valeur Brutto", "Position VAT rate": "Taux de TVA",
            "PSA value": "Valeur PSA", "PSA": "PSA",
        },
        "extra_columns": ["N° Camion", "Chauffeur", "Vendeur", "Tournée"],
    },
    "Momo": {
        "extract_only": True,
        "icon": "📱",
        "description": "Traitement des flux mobiles (Momo) : conversion CSV/Excel et normalisation des transactions.",
        "tooltip": "Traitez vos rapports d'opérations mobiles en un clic.",
        "columns": ["Id", "Date", "Status", "Type", "From", "To name", "Amount", "Balance"],
        "rename": {"Id": "N° Identification", "From": "Provenance", "To name": "To handler name"},
        "extra_columns": ["Vendeur", "Compte", "Tournée"],
    },
    "FICHIER DES RISTOURNES": {
        "extract_only": True,
        "icon": "💰",
        "description": "Calcul des ristournes : récapitulatif structuré des remises quotidiennes par magasin.",
        "tooltip": "Générez vos états de remises sur ventes en quelques secondes.",
        "columns": ["Code", "Store name", "Start date", "Document number", "Product code", "Product name", "Family", "CODE RABATES", "RISTOURNE"],
        "rename": {},
        "extra_columns": [],
    },
    "BASE MAGASIN": {
        "extract_only": True,
        "icon": "🏪",
        "description": "Mouvements de stock : analyse fine des opérations magasin (Chargement, Déchargement, etc.).",
        "tooltip": "Découpe automatiquement vos documents en Opérations et Comptes.",
        "columns": ["Product", "Change date", "Prev. quantity", "Cur. quantity", "Position record date", "Change source", "Document"],
        "rename": {
            "Product": "Product", "Change date": "Date Opération", "Prev. quantity": "Quantité Initiale",
            "Cur. quantity": "Quantité Résultante", "Position record date": "Journée", "Change source": "Référence", "Document": "Opération"
        },
        "extra_columns": ["Compte", "Responsable", "Camion", "Chauffeur"],
    },
    "STOCKS DES PRODUITS": {
        "extract_only": True,
        "icon": "📊",
        "description": "Inventaire financier : valorisation des stocks avec arrondis comptables et traduction de nature.",
        "tooltip": "Valorisez vos stocks au PA HT avec arrondis automatiques.",
        "columns": ["Code", "Name", "Category", "Price", "Qty", "Totalvalue"],
        "rename": {"Code": "Code", "Name": "Nom du Produit", "Category": "Nature", "Price": "Prix Achat HT", "Qty": "Qty", "Totalvalue": "Valeur au PA HT"},
        "extra_columns": ["Conditionnement", "PA TTC", "Valeur au PA TTC", "Prix Vente HT", "Prix de Vente 2%", "Valeur au PV TTC"],
    },
    "BASE CLIENTS": {
        "extract_only": True,
        "icon": "🏘️",
        "description": "Base CRM : extraction des paramètres clients (localisation, fiscalité, classification).",
        "tooltip": "Exportez votre base client avec ses coordonnées et typologies.",
        "columns": [
            "Code", "Name", "Street", "OBJ_CONTACT|BusinessPhone", "OBJ_PARAM_NC",
            "OBJ_PARAM_ADDITIONAL_TAX", "OBJ_PARAM_DISTANCE", "OBJ_PARAM_TYPOLOGIE",
            "OBJ_PARAM_ITINERAIRE", "Creation Date"
        ],
        "rename": {
            "Street": "Localisation", "OBJ_CONTACT|BusinessPhone": "Numéro Téléphone", "OBJ_PARAM_NC": "Numéro d'Identification Unique",
            "OBJ_PARAM_ADDITIONAL_TAX": "Statut Fiscal", "OBJ_PARAM_DISTANCE": "Position par rapport au Centre",
            "OBJ_PARAM_TYPOLOGIE": "TYPOLOGIE", "OBJ_PARAM_ITINERAIRE": "Classification",
        },
        "extra_columns": ["Route", "Tournée"],
    },
    "FICHIER DES INVENTAIRES": {
        "extract_only": True,
        "icon": "📝",
        "description": "Clôture journalière : nettoyage des stocks nuls et calcul des écarts d'inventaire.",
        "tooltip": "Supprime automatiquement les lignes sans stock pour un inventaire clair.",
        "columns": ["Date", "Code Article", "Nom Article", "Conditionnement", "Quantité Initiale", "Quantité Comptée", "Valeur Ecart"],
        "rename": {},
        "extra_columns": ["Valeur Stock PV TTC 2%"],
    }
}

# Constantes & Mappings
INDEX_SHEET_NAME = "Données"
TRACE_COLS = ["Fichier source", "Date d'import"]
BASE_MAGASIN_OPS = {"LR": "Chargement", "UC": "Déchargement", "DO": "Sorties directes de stock", "DI": "Entrées Directes de stock", "BO": "Achat ou Commande Appro", "DE": "Retour Emballages SABC"}
BASE_MAGASIN_COMPTES = {"S02004": "R1", "S02003": "R2", "S02005": "R3", "S02006": "R4"}
STOCKS_NATURES = {"BEER": "BIÈRES", "BG": "BOISSONS RAFRAICHISSANTE SANS ALCOOL", "AM": "ALCOOL MIX", "EAU": "EAUX MINERALES NATURELLES", "EMB": "PACKAGES"}

# ──────────────────────────────────────────────
#  Gestion Mega.nz
# ──────────────────────────────────────────────
def get_mega_client():
    if Mega is None:
        st.error("Le module 'mega.py' n'est pas installé.")
        return None
    if "mega_client" not in st.session_state:
        try:
            email = st.secrets.get("MEGA_EMAIL")
        except Exception:
            email = None
        try:
            password = st.secrets.get("MEGA_PASSWORD")
        except Exception:
            password = None

        email = email or st.sidebar.text_input("Email Mega", type="default")
        password = password or st.sidebar.text_input("Mot de passe Mega", type="password")

        if email and password:
            try:
                mega = Mega()
                st.session_state.mega_client = mega.login(email, password)
            except Exception as e:
                st.error(f"Erreur de connexion Mega : {e}")
                return None
        else:
            return None
    return st.session_state.mega_client

def upload_to_mega(file_content, filename, folder_name="DataHub_Indexes"):
    import tempfile, uuid
    m = get_mega_client()
    if not m:
        return False

    folder = m.find(folder_name)
    if not folder:
        folder = m.create_folder(folder_name)

    existing = m.find(filename)
    if existing:
        m.destroy(existing[0])

    # Dossier temporaire unique (compatible Windows et Linux)
    tmp_dir = os.path.join(tempfile.gettempdir(), f"mega_{uuid.uuid4().hex}")
    os.makedirs(tmp_dir, exist_ok=True)
    temp_path = os.path.join(tmp_dir, filename)

    try:
        with open(temp_path, "wb") as f:
            f.write(file_content)
        m.upload(temp_path, folder[0])
        return True
    except Exception as e:
        st.error(f"Erreur upload Mega : {e}")
        return False
    finally:
        try:
            import shutil
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass

def download_from_mega(filename):
    import tempfile, uuid
    m = get_mega_client()
    if not m:
        return None
    try:
        files = m.get_files()
        target = next((f for f in files.values() if f.get('a') and f['a'].get('n') == filename), None)
        if not target:
            return None

        # Dossier temporaire unique (compatible Windows et Linux)
        tmp_dir = os.path.join(tempfile.gettempdir(), f"mega_{uuid.uuid4().hex}")
        os.makedirs(tmp_dir, exist_ok=True)

        m.download(target, tmp_dir)

        downloaded_path = os.path.join(tmp_dir, filename)
        if not os.path.exists(downloaded_path):
            return None

        with open(downloaded_path, "rb") as f:
            content = f.read()

        return content
    except Exception as e:
        st.error(f"Erreur téléchargement Mega : {e}")
        return None
    finally:
        try:
            import shutil
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass

def list_mega_indexes(folder_name="DataHub_Indexes"):
    m = get_mega_client()
    if not m: return []
    folder = m.find(folder_name)
    if not folder: return []

    files = m.get_files()
    index_files = []
    for f_id, f_info in files.items():
        if f_info['p'] == folder[0] and f_info['t'] == 0:
            index_files.append(f_info['a']['n'])
    return index_files

# ──────────────────────────────────────────────
#  Styles Excel & Helpers
# ──────────────────────────────────────────────
def _header_style(ws, nb_cols, row=1):
    for col_idx in range(1, nb_cols + 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", start_color="2E4057")
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
            if val: max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 50)

# ──────────────────────────────────────────────
#  Logique de Traitement (Core)
# ──────────────────────────────────────────────
def filter_columns(df, sheet_name, file_type):
    if file_type is None or file_type not in COLUMNS_BY_TYPE: return df, [], list(df.columns)
    config = COLUMNS_BY_TYPE[file_type]
    expected = [c for c in config["columns"]]
    rename_map = config.get("rename", {})
    if file_type == "STOCKS DES PRODUITS" and "Code" not in df.columns and len(df.columns) >= 7: df["Code"] = df.iloc[:, 6]
    available = [c for c in expected if c in df.columns]
    missing = [c for c in expected if c not in df.columns]
    df_out = df[available].copy()
    if file_type == "FICHIER DES INVENTAIRES":
        if "Quantité Initiale" in df_out.columns and "Quantité Comptée" in df_out.columns:
            df_out = df_out[~((df_out["Quantité Initiale"] == 0) & (df_out["Quantité Comptée"] == 0))]
    elif file_type == "BASE MAGASIN":
        if "Position record date" in df_out.columns: df_out["Position record date"] = pd.to_datetime(df_out["Position record date"], errors='coerce').dt.strftime('%d/%m/%Y')
        if "Document" in df_out.columns:
            def extract_op_compte(doc_str):
                doc_str = str(doc_str); op_val = ""; compte_val = ""
                for code, label in BASE_MAGASIN_OPS.items():
                    if code in doc_str: op_val = label; break
                for code, label in BASE_MAGASIN_COMPTES.items():
                    if code in doc_str: compte_val = label; break
                return pd.Series([op_val, compte_val])
            df_out[["Document", "Compte"]] = df_out["Document"].apply(extract_op_compte)
        else: df_out["Compte"] = ""
    elif file_type == "STOCKS DES PRODUITS":
        if "Category" in df_out.columns: df_out["Category"] = df_out["Category"].map(STOCKS_NATURES).fillna(df_out["Category"])
        for col in ["Price", "Totalvalue"]:
            if col in df_out.columns: df_out[col] = pd.to_numeric(df_out[col], errors='coerce').apply(lambda x: np.ceil(x) if pd.notnull(x) else x)
    df_out = df_out.rename(columns=rename_map)
    for extra_col in config.get("extra_columns", []):
        if extra_col not in df_out.columns: df_out[extra_col] = ""
    return df_out, missing, list(df_out.columns)

def load_data(uploaded_file):
    filename = uploaded_file.name
    if filename.endswith('.csv'):
        try: df = pd.read_csv(uploaded_file, sep=None, engine='python')
        except: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, sep=',')
        return {"Données_CSV": df}
    else: return pd.read_excel(uploaded_file, sheet_name=None)

def process_multiple_files(uploaded_files, file_type):
    all_results = []
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for uploaded_file in uploaded_files:
            uploaded_file.seek(0); excel_data = load_data(uploaded_file)
            file_sheets_res = []
            for sheet_name, df in excel_data.items():
                df_filtered, missing_cols, exported_cols = filter_columns(df, sheet_name, file_type)
                extract_only = COLUMNS_BY_TYPE[file_type].get("extract_only", False) if file_type in COLUMNS_BY_TYPE else False
                if extract_only: df_result = df_filtered.reset_index(drop=True); unique_rows, duplicate_rows = len(df_result), 0
                else: df_result = df_filtered.drop_duplicates(); unique_rows = len(df_result); duplicate_rows = len(df) - unique_rows
                safe_name = uploaded_file.name.split('.')[0][:15]
                sheet_export_name = f"{safe_name}_{sheet_name[:10]}".replace(' ', '_').replace('.', '_')
                df_result.to_excel(writer, sheet_name=sheet_export_name, index=False)
                file_sheets_res.append({"sheet_name": sheet_name, "total_columns": len(df.columns), "total_rows": len(df), "unique_rows": unique_rows, "duplicate_rows": duplicate_rows, "extract_only": extract_only, "df_result": df_result})
            all_results.append({"filename": uploaded_file.name, "sheets": file_sheets_res})
    output.seek(0); wb = load_workbook(output); ws_rep = wb.create_sheet("Rapport Global", 0)
    headers = ["Fichier", "Feuille", "Cols", "Lignes Tot.", "Uniques", "Doublons", "Mode"]
    for i, h in enumerate(headers, 1):
        cell = ws_rep.cell(row=1, column=i, value=h); cell.fill = PatternFill("solid", start_color="2E4057"); cell.font = Font(name="Arial", bold=True, color="FFFFFF"); cell.alignment = Alignment(horizontal="center")
    curr_row = 2
    for file_res in all_results:
        for r in file_res["sheets"]:
            vals = [file_res["filename"], r["sheet_name"], r["total_columns"], r["total_rows"], r["unique_rows"], r["duplicate_rows"], "Extraction" if r["extract_only"] else "Dédoublonnage"]
            for c, v in enumerate(vals, 1): cell = ws_rep.cell(row=curr_row, column=c, value=v); cell.border = _border()
            curr_row += 1
    apply_auto_width(ws_rep)
    for sheet_name in wb.sheetnames:
        if sheet_name != "Rapport Global": ws = wb[sheet_name]; _header_style(ws, ws.max_column); apply_auto_width(ws)
    final_output = io.BytesIO(); wb.save(final_output)
    return final_output.getvalue(), all_results

def update_index_content(all_results, existing_index_content=None):
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    frames = []
    for file_res in all_results:
        for r in file_res["sheets"]:
            df = r["df_result"].copy(); df.insert(0, TRACE_COLS[0], file_res["filename"]); df.insert(1, TRACE_COLS[1], now_str); frames.append(df)
    if not frames: return None
    df_new = pd.concat(frames, ignore_index=True)

    if existing_index_content:
        wb_idx = load_workbook(io.BytesIO(existing_index_content))
        if INDEX_SHEET_NAME not in wb_idx.sheetnames: ws_idx = wb_idx.create_sheet(INDEX_SHEET_NAME); next_row = 1
        else: ws_idx = wb_idx[INDEX_SHEET_NAME]; next_row = ws_idx.max_row + 1
    else: wb_idx = Workbook(); ws_idx = wb_idx.active; ws_idx.title = INDEX_SHEET_NAME; next_row = 1

    if next_row == 1:
        for c, col in enumerate(df_new.columns, 1): ws_idx.cell(row=1, column=c, value=col)
        _header_style(ws_idx, len(df_new.columns)); next_row = 2

    for r_idx, row in enumerate(df_new.itertuples(index=False)):
        curr_r = next_row + r_idx
        for c_idx, val in enumerate(row, 1):
            cell = ws_idx.cell(row=curr_r, column=c_idx, value=val); cell.border = _border()
            if c_idx <= 2: cell.font = Font(name="Arial", size=9, italic=True, color="555555"); cell.fill = PatternFill("solid", start_color="E8F5E9")

    apply_auto_width(ws_idx); output = io.BytesIO(); wb_idx.save(output); return output.getvalue()

# ──────────────────────────────────────────────
#  Interface Streamlit (Premium UI/UX)
# ──────────────────────────────────────────────
st.set_page_config(page_title="DataHub Pro | Mega Cloud", page_icon="📈", layout="wide")
st.markdown("""<style>
    .main { background-color: #f0f2f6; }
    .stApp { background-color: #f0f2f6; }
    .stButton>button { border-radius: 12px; height: 3.5em; font-weight: 700; transition: all 0.3s; }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .step-card { background-color: #ffffff; padding: 25px; border-radius: 15px; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-top: 5px solid #2E4057; }
    .step-header { display: flex; align-items: center; margin-bottom: 15px; }
    .step-number { background-color: #2E4057; color: white; border-radius: 50%; width: 30px; height: 30px; display: flex; align-items: center; justify-content: center; font-weight: bold; margin-right: 15px; }
    .step-title { color: #2E4057; font-size: 1.3em; font-weight: bold; }
    .stat-card { background-color: #f8f9fa; padding: 15px; border-radius: 10px; text-align: center; border: 1px solid #e9ecef; }
    .stat-val { font-size: 1.8em; font-weight: bold; color: #2E4057; }
    .stat-label { font-size: 0.9em; color: #6c757d; }
    .footer { text-align: center; color: #adb5bd; padding: 40px 0; font-size: 0.85em; }
    .nav-button { border-radius: 10px; padding: 10px; text-align: center; background-color: #ffffff; cursor: pointer; border: 1px solid #e9ecef; }
</style>""", unsafe_allow_html=True)

# Initialisation Session State
if 'results' not in st.session_state: st.session_state.results = None
if 'processed_data' not in st.session_state: st.session_state.processed_data = None
if 'current_file_type' not in st.session_state: st.session_state.current_file_type = None
if 'page' not in st.session_state: st.session_state.page = "Traitement"

# Sidebar
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/microsoft-excel-2019.png", width=60)
    st.markdown("### ⚙️ Centre de Contrôle")

    st.markdown("---")
    if st.button("🚀 Traitement des fichiers", use_container_width=True): st.session_state.page = "Traitement"; st.rerun()
    if st.button("📂 Bibliothèque des Index (Mega)", use_container_width=True): st.session_state.page = "Index"; st.rerun()

    st.markdown("---")
    st.markdown("### ☁️ Connexion Mega.nz")
    get_mega_client()

    if st.button("🔄 Réinitialiser la session", use_container_width=True):
        for key in list(st.session_state.keys()): del st.session_state[key]
        st.rerun()

# ──────────────────────────────────────────────
#  PAGE : TRAITEMENT
# ──────────────────────────────────────────────
if st.session_state.page == "Traitement":
    st.markdown("<h1>🚀 Traitement Intelligent</h1>", unsafe_allow_html=True)

    col_main1, col_main2 = st.columns([1, 1])

    with col_main1:
        st.markdown('<div class="step-card"><div class="step-header"><div class="step-number">1</div><div class="step-title">Configuration</div></div>', unsafe_allow_html=True)
        file_type = st.selectbox("Module métier", options=[None] + list(COLUMNS_BY_TYPE.keys()), format_func=lambda x: "✨ Traitement Standard" if x is None else f"{COLUMNS_BY_TYPE[x]['icon']} {x}")
        if file_type: st.markdown(f"*{COLUMNS_BY_TYPE[file_type]['description']}*")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_main2:
        st.markdown('<div class="step-card"><div class="step-header"><div class="step-number">2</div><div class="step-title">Source de Données</div></div>', unsafe_allow_html=True)
        uploaded_files = st.file_uploader("Glissez vos fichiers Excel ou CSV ici", type=["xlsx", "csv"], accept_multiple_files=True)
        if uploaded_files:
            if st.button("🚀 Lancer l'analyse", use_container_width=True, type="primary"):
                with st.spinner("Traitement en cours..."):
                    try:
                        p_data, results = process_multiple_files(uploaded_files, file_type)
                        st.session_state.processed_data = p_data; st.session_state.results = results; st.session_state.current_file_type = file_type; st.balloons()
                    except Exception as e: st.error(f"⚠️ Erreur : {str(e)}")
        st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.results:
        st.markdown('<div class="step-card"><div class="step-header"><div class="step-number">3</div><div class="step-title">Tableau de Bord & Actions Cloud</div></div>', unsafe_allow_html=True)

        total_in = sum(sum(s['total_rows'] for s in f['sheets']) for f in st.session_state.results)
        total_out = sum(sum(s['unique_rows'] for s in f['sheets']) for f in st.session_state.results)
        s1, s2, s3 = st.columns(3)
        with s1: st.markdown(f'<div class="stat-card"><div class="stat-val">{total_in}</div><div class="stat-label">Lignes Entrantes</div></div>', unsafe_allow_html=True)
        with s2: st.markdown(f'<div class="stat-card"><div class="stat-val" style="color:#28a745;">{total_out}</div><div class="stat-label">Lignes Uniques</div></div>', unsafe_allow_html=True)
        with s3: st.markdown(f'<div class="stat-card"><div class="stat-val" style="color:#dc3545;">{total_in - total_out}</div><div class="stat-label">Doublons Éliminés</div></div>', unsafe_allow_html=True)

        st.markdown("---")

        c_left, c_right = st.columns([2, 1])
        with c_left:
            st.markdown("### 📋 Aperçu")
            tab_data = [{"Fichier": f["filename"], "Uniques": sum(s['unique_rows'] for s in f['sheets'])} for f in st.session_state.results]
            st.dataframe(pd.DataFrame(tab_data), use_container_width=True, hide_index=True)

        with c_right:
            st.markdown("### 📥 Actions")
            suffix = "Standard" if st.session_state.current_file_type is None else st.session_state.current_file_type.replace(" ", "_")
            st.download_button("💾 Exporter Fichier Traité", data=st.session_state.processed_data, file_name=f"Export_{suffix}.xlsx", use_container_width=True)

            st.markdown("---")
            st.markdown("#### ☁️ Centralisation Cloud")
            if st.button("🔗 Fusionner avec l'Index Central (Mega)", use_container_width=True, type="secondary"):
                if not get_mega_client():
                    st.warning("Veuillez vous connecter à Mega dans la barre latérale.")
                else:
                    with st.spinner("Fusion sur le Cloud Mega..."):
                        idx_filename = f"index_{suffix}.xlsx"
                        existing_content = download_from_mega(idx_filename)
                        new_index_content = update_index_content(st.session_state.results, existing_content)
                        if upload_to_mega(new_index_content, idx_filename):
                            st.success(f"Fusion réussie dans {idx_filename} sur Mega !")
                        else:
                            st.error("Erreur lors de l'upload vers Mega.")
        st.markdown('</div>', unsafe_allow_html=True)

# ──────────────────────────────────────────────
#  PAGE : BIBLIOTHÈQUE DES INDEX
# ──────────────────────────────────────────────
elif st.session_state.page == "Index":
    st.markdown("<h1>📂 Bibliothèque des Index (Mega Cloud)</h1>", unsafe_allow_html=True)

    if not get_mega_client():
        st.info("Veuillez vous connecter à Mega.nz via la barre latérale pour accéder à vos index centralisés.")
    else:
        with st.spinner("Chargement de la bibliothèque Mega..."):
            indexes = list_mega_indexes()

        if not indexes:
            st.warning("Aucun index trouvé sur Mega.nz. Effectuez un traitement et fusionnez les données pour créer votre premier index.")
        else:
            st.markdown(f"Retrouvez ici tous vos index centralisés. Vous pouvez les visualiser ou les télécharger.")

            for idx_name in sorted(indexes):
                with st.expander(f"📄 {idx_name}", expanded=False):
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.markdown(f"**Nom du fichier** : `{idx_name}`")
                        st.markdown(f"**Catégorie** : {idx_name.replace('index_', '').replace('.xlsx', '').replace('_', ' ')}")
                    with col2:
                        content = download_from_mega(idx_name)
                        if content:
                            st.download_button(f"📥 Télécharger", data=content, file_name=idx_name, key=idx_name, use_container_width=True)

st.markdown('<div class="footer">DataHub Pro v3.0 | Cloud Sync Mega.nz | © 2024</div>', unsafe_allow_html=True)
