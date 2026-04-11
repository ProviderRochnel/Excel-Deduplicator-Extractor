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

# ──────────────────────────────────────────────
#  Configuration des colonnes
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
    },
    "FICHIER DES RISTOURNES": {
        "extract_only": True,
        "description": "💰 **Fichier des Ristournes** : Extraction et organisation des données de ristournes quotidiennes.",
        "columns": [
            "Code", "Store name", "Start date", "Document number",
            "Product code", "Product name", "Family", "CODE RABATES", "RISTOURNE"
        ],
        "rename": {},
        "extra_columns": [],
    },
    "BASE MAGASIN": {
        "extract_only": True,
        "description": "🏪 **Base Magasin** : Gestion des mouvements de stock avec découpage de l'opération et du compte.",
        "columns": [
            "Product", "Change date", "Prev. quantity", "Cur. quantity",
            "Position record date", "Change source", "Document"
        ],
        "rename": {
            "Product":              "Product",
            "Change date":          "Date Opération",
            "Prev. quantity":       "Quantité Initiale",
            "Cur. quantity":        "Quantité Résultante",
            "Position record date": "Journée",
            "Change source":        "Référence",
            "Document":             "Opération"
        },
        "extra_columns": ["Compte", "Responsable", "Camion", "Chauffeur"],
    },
    "STOCKS DES PRODUITS": {
        "extract_only": True,
        "description": "📊 **Stocks des Produits** : Analyse des stocks (Quantité/Valeur) avec traduction de nature et calculs financiers.",
        "columns": ["Code", "Name", "Category", "Price", "Qty", "Totalvalue"],
        "rename": {
            "Code":       "Code",
            "Name":       "Nom du Produit",
            "Category":   "Nature",
            "Price":      "Prix Achat HT",
            "Qty":        "Qty",
            "Totalvalue": "Valeur au PA HT"
        },
        "extra_columns": [
            "Conditionnement", "PA TTC", "Valeur au PA TTC",
            "Prix Vente HT", "Prix de Vente 2%", "Valeur au PV TTC"
        ],
    },
    "BASE CLIENTS": {
        "extract_only": True,
        "description": "🏘️ **Base Clients** : Récapitulatif complet des paramètres clients avec mapping métier.",
        "columns": [
            "Code", "Name", "Street", "OBJ_CONTACT|BusinessPhone",
            "OBJ_PARAM_NC", "OBJ_PARAM_ADDITIONAL_TAX", "OBJ_PARAM_DISTANCE",
            "OBJ_PARAM_TYPOLOGIE", "OBJ_PARAM_ITINERAIRE", "Creation Date"
        ],
        "rename": {
            "Street":                    "Localisation",
            "OBJ_CONTACT|BusinessPhone": "Numéro Téléphone",
            "OBJ_PARAM_NC":              "Numéro d'Identification Unique",
            "OBJ_PARAM_ADDITIONAL_TAX":  "Statut Fiscal",
            "OBJ_PARAM_DISTANCE":        "Position par rapport au Centre",
            "OBJ_PARAM_TYPOLOGIE":       "TYPOLOGIE",
            "OBJ_PARAM_ITINERAIRE":      "Classification",
        },
        "extra_columns": ["Route", "Tournée"],
    },
    "FICHIER DES INVENTAIRE": {
        "extract_only": True,
        "description": "📋 **Inventaire Journalier** : Stock compté à la clôture — élimination des lignes sans opération, ajout Valeur Stock PV TTC 2%.",
        # Colonnes attendues APRÈS pré-traitement (noms normalisés)
        "columns": [
            "Date", "Type Article", "Code Article", "Nom Article",
            "Conditionnement", "Quantité Initiale", "Quantité Comptée",
            "Quantité Ecart", "Valeur Ecart"
        ],
        "rename": {},           # déjà normalisés dans le pré-traitement
        "extra_columns": ["Valeur Stock PV TTC 2%"],
    },
}

# Constantes de traçabilité
INDEX_SHEET_NAME = "Données"
TRACE_COLS       = ["Fichier source", "Date d'import"]

# Mapping spécifique BASE MAGASIN
BASE_MAGASIN_OPS = {
    "LR": "Chargement", "UC": "Déchargement", "DO": "Sorties directes de stock",
    "DI": "Entrées Directes de stock", "BO": "Achat ou Commande Appro",
    "DE": "Retour Emballages SABC"
}
BASE_MAGASIN_COMPTES = {
    "S02004": "R1", "S02003": "R2", "S02005": "R3", "S02006": "R4"
}

# Mapping spécifique STOCKS DES PRODUITS
STOCKS_NATURES = {
    "BEER": "BIÈRES",
    "BG":   "BOISSONS RAFRAICHISSANTE SANS ALCOOL",
    "AM":   "ALCOOL MIX",
    "EAU":  "EAUX MINERALES NATURELLES",
    "EMB":  "PACKAGES"
}

# ──────────────────────────────────────────────
#  Pré-traitement spécifique INVENTAIRE
# ──────────────────────────────────────────────
def preprocess_inventory(raw_df):
    """
    Transforme le format brut Eleader en DataFrame exploitable.
    Gestion robuste des erreurs pour éviter les plantages.
    """
    try:
        # Validation initiale
        if raw_df is None or raw_df.empty:
            raise ValueError("DataFrame vide ou None")
        
        if raw_df.shape[0] < 8:
            raise ValueError(f"Format invalide : nécessite au moins 8 lignes, trouvé {raw_df.shape[0]}")
        
        # ── 1. Extraire la date avec gestion d'erreur ────
        date_str = ""
        try:
            if raw_df.shape[0] > 3:
                cell_val = str(raw_df.iloc[3, 0])
                m = re.search(r'\d{2}/\d{2}/\d{4}', cell_val)
                date_str = m.group(0) if m else "Date Inconnue"
        except Exception:
            date_str = "Date Inconnue"

        # ── 2. Extraire et nettoyer les en-têtes avec validation ────────
        try:
            if raw_df.shape[0] <= 7:
                raise ValueError("Pas assez de lignes pour les en-têtes")
            
            raw_headers = raw_df.iloc[7].tolist()
            clean_headers = [
                str(h).replace('\n', ' ').strip() if pd.notna(h) else f"_col{i}"
                for i, h in enumerate(raw_headers)
            ]
            
            if len(clean_headers) < 13:
                raise ValueError(f"Pas assez de colonnes : {len(clean_headers)} < 13")
                
        except Exception as e:
            raise ValueError(f"Erreur traitement en-têtes : {e}")

        # ── 3. Construire le DataFrame de données avec validation ─
        try:
            data = raw_df.iloc[8:].copy()
            if data.empty:
                raise ValueError("Aucune donnée trouvée après la ligne 8")
            
            data.columns = clean_headers
            data = data.reset_index(drop=True)
        except Exception as e:
            raise ValueError(f"Erreur construction DataFrame : {e}")

        # ── 4. Sélectionner et renommer les colonnes avec validation ────────
        try:
            required_positions = [0, 2, 3, 6, 7, 8, 10, 12]
            for pos in required_positions:
                if pos >= len(clean_headers):
                    raise ValueError(f"Position de colonne {pos} hors limites")
            
            col_map = {
                clean_headers[0]:  "Type Article",
                clean_headers[2]:  "Code Article",
                clean_headers[3]:  "Nom Article",
                clean_headers[6]:  "Conditionnement",
                clean_headers[7]:  "Quantité Initiale",
                clean_headers[8]:  "Quantité Comptée",
                clean_headers[10]: "Quantité Ecart",
                clean_headers[12]: "Valeur Ecart",
            }
            
            # Vérifier que toutes les colonnes existent
            missing_cols = [col for col in col_map.keys() if col not in data.columns]
            if missing_cols:
                raise ValueError(f"Colonnes manquantes : {missing_cols}")
            
            data = data[list(col_map.keys())].rename(columns=col_map)
            
        except Exception as e:
            raise ValueError(f"Erreur mapping colonnes : {e}")

        # ── 5. Forward-fill "Type Article" avec validation ─────────────
        try:
            if "Type Article" in data.columns:
                data["Type Article"] = data["Type Article"].replace(["", "nan", "None"], np.nan)
                data["Type Article"] = data["Type Article"].ffill()
        except Exception as e:
            raise ValueError(f"Erreur traitement Type Article : {e}")

        # ── 6. Supprimer les lignes "TOTAL" avec gestion d'erreur ─
        try:
            if "Code Article" in data.columns:
                data = data[data["Code Article"].astype(str).str.strip().str.upper() != "TOTAL"]
            if "Type Article" in data.columns:
                data = data[data["Type Article"].astype(str).str.strip().str.upper() != "TOTAL"]
        except Exception as e:
            raise ValueError(f"Erreur suppression lignes TOTAL : {e}")

        # ── 7. Convertir les colonnes numériques avec validation ─────────────────────────
        numeric_cols = ["Quantité Initiale", "Quantité Comptée", "Quantité Ecart", "Valeur Ecart"]
        for col in numeric_cols:
            try:
                if col in data.columns:
                    data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0)
            except Exception as e:
                raise ValueError(f"Erreur conversion colonne {col} : {e}")

        # ── 8. Supprimer les lignes sans opération avec validation ────────────────
        try:
            if "Quantité Initiale" in data.columns and "Quantité Comptée" in data.columns:
                before_count = len(data)
                data = data[~((data["Quantité Initiale"] == 0) & (data["Quantité Comptée"] == 0))]
                after_count = len(data)
                
                # Log silencieux du nombre de lignes supprimées
                if before_count > after_count:
                    pass  # {before_count - after_count} lignes sans opération supprimées
        except Exception as e:
            raise ValueError(f"Erreur filtrage lignes sans opération : {e}")

        # ── 9. Ajouter la colonne Date avec validation ──────────────
        try:
            data.insert(0, "Date", date_str)
        except Exception as e:
            raise ValueError(f"Erreur ajout colonne Date : {e}")

        # Validation finale
        if data.empty:
            raise ValueError("DataFrame final vide après traitement")
        
        data = data.reset_index(drop=True)
        return data
        
    except Exception as e:
        # Créer un DataFrame minimal en cas d'erreur critique
        error_msg = f"Erreur pré-traitement inventaire : {e}"
        return pd.DataFrame({"Erreur": [error_msg]})


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
        cell.fill = header_fill; cell.font = header_font
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
            existing_headers = [
                ws_idx.cell(row=1, column=c).value
                for c in range(1, ws_idx.max_column + 1)
                if ws_idx.cell(row=1, column=c).value
            ]
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
    alt_fill   = PatternFill("solid", start_color="EAF0FB")
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

    config     = COLUMNS_BY_TYPE[file_type]
    expected   = list(config["columns"])
    rename_map = config.get("rename", {})

    # Mapper la colonne 7 vers 'Code' si nécessaire (pour STOCKS DES PRODUITS)
    if file_type == "STOCKS DES PRODUITS" and "Code" not in df.columns and len(df.columns) >= 7:
        try:
            df["Code"] = df.iloc[:, 6]
        except Exception as e:
            st.warning(f"⚠️ Impossible de mapper la colonne 7 vers 'Code' : {e}")

    available = [c for c in expected if c in df.columns]
    missing   = [c for c in expected if c not in df.columns]

    df_out = df[available].copy()

    # ── Logique BASE MAGASIN ──────────────────────────────────────────
    if file_type == "BASE MAGASIN":
        if "Position record date" in df_out.columns:
            df_out["Position record date"] = pd.to_datetime(
                df_out["Position record date"], errors='coerce'
            ).dt.strftime('%d/%m/%Y')
        if "Document" in df_out.columns:
            def extract_op_compte(doc_str):
                doc_str = str(doc_str)
                op_val = ""; compte_val = ""
                for code, label in BASE_MAGASIN_OPS.items():
                    if code in doc_str: op_val = label; break
                for code, label in BASE_MAGASIN_COMPTES.items():
                    if code in doc_str: compte_val = label; break
                return pd.Series([op_val, compte_val])
            df_out[["Document", "Compte"]] = df_out["Document"].apply(extract_op_compte)
        else:
            df_out["Compte"] = ""

    # ── Logique STOCKS DES PRODUITS ───────────────────────────────────
    elif file_type == "STOCKS DES PRODUITS":
        if "Category" in df_out.columns:
            df_out["Category"] = df_out["Category"].map(STOCKS_NATURES).fillna(df_out["Category"])
        if "Price" in df_out.columns:
            df_out["Price"] = pd.to_numeric(df_out["Price"], errors='coerce').apply(
                lambda x: np.ceil(x) if pd.notnull(x) else x)
        if "Totalvalue" in df_out.columns:
            df_out["Totalvalue"] = pd.to_numeric(df_out["Totalvalue"], errors='coerce').apply(
                lambda x: np.ceil(x) if pd.notnull(x) else x)

    # ── Logique FICHIER DES INVENTAIRE ───────────────────────────────
    # (Le pré-traitement lourd est déjà fait dans preprocess_inventory.
    #  Ici on s'assure juste que les colonnes numériques sont bien typées.)
    elif file_type == "FICHIER DES INVENTAIRE":
        for col in ["Quantité Initiale", "Quantité Comptée", "Quantité Ecart", "Valeur Ecart"]:
            if col in df_out.columns:
                df_out[col] = pd.to_numeric(df_out[col], errors='coerce').fillna(0)

    # ── Renommage ─────────────────────────────────────────────────────
    df_out = df_out.rename(columns=rename_map)

    # ── Colonnes extra vides ──────────────────────────────────────────
    for extra_col in config.get("extra_columns", []):
        if extra_col not in df_out.columns:
            df_out[extra_col] = ""

    return df_out, missing, list(df_out.columns)


def load_data(uploaded_file, file_type=None):
    """
    Charge un fichier uploadé et retourne un dict {sheet_name: DataFrame}.
    Gestion robuste des erreurs et validation des fichiers.
    """
    try:
        if uploaded_file is None:
            raise ValueError("Fichier None fourni")
        
        filename = uploaded_file.name
        
        # Validation du fichier
        if not filename:
            raise ValueError("Nom de fichier vide")
        
        file_size = len(uploaded_file.getvalue())
        if file_size == 0:
            raise ValueError("Fichier vide")
        
        if file_size > 50 * 1024 * 1024:  # 50MB limite
            raise ValueError(f"Fichier trop volumineux : {file_size/1024/1024:.1f}MB > 50MB")

        # Reset position pour lecture
        uploaded_file.seek(0)

        if filename.endswith('.csv'):
            return _load_csv_with_fallback(uploaded_file)
        
        elif filename.endswith(('.xlsx', '.xls')):
            if file_type == "FICHIER DES INVENTAIRE":
                return _load_inventory_excel(uploaded_file)
            else:
                return _load_standard_excel(uploaded_file)
        
        else:
            raise ValueError(f"Type de fichier non supporté : {filename}")
            
    except Exception as e:
        error_msg = f"Erreur chargement fichier {filename if 'filename' in locals() else 'inconnu'} : {e}"
        return {"Erreur": pd.DataFrame({"Erreur": [error_msg]})}


def _load_csv_with_fallback(uploaded_file):
    """Charge CSV avec plusieurs méthodes de fallback"""
    try:
        # Essai avec détection automatique
        df = pd.read_csv(uploaded_file, sep=None, engine='python', encoding='utf-8')
        return {"Données_CSV": df}
    except Exception:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=',', encoding='utf-8')
            return {"Données_CSV": df}
        except Exception:
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
                return {"Données_CSV": df}
            except Exception:
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, sep=',', encoding='latin-1')
                    return {"Données_CSV": df}
                except Exception as e:
                    raise ValueError(f"Impossible de lire le CSV : {e}")


def _load_inventory_excel(uploaded_file):
    """Charge Excel pour inventaire avec pré-traitement"""
    try:
        uploaded_file.seek(0)
        raw_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None, dtype=str)
        
        if not raw_sheets:
            raise ValueError("Aucune feuille trouvée dans le fichier Excel")
        
        result = {}
        for sheet_name, raw_df in raw_sheets.items():
            try:
                processed_df = preprocess_inventory(raw_df)
                
                # Vérifier si le pré-traitement a généré une erreur
                if "Erreur" in processed_df.columns:
                    st.warning(f"⚠️ Erreur pré-traitement feuille « {sheet_name} » : {processed_df['Erreur'].iloc[0]}")
                    result[sheet_name] = processed_df
                else:
                    result[sheet_name] = processed_df
                    
            except Exception as e:
                error_df = pd.DataFrame({"Erreur": [f"Échec traitement feuille {sheet_name} : {e}"]})
                result[sheet_name] = error_df
                st.warning(f"⚠️ Feuille « {sheet_name} » ignorée : {e}")
        
        return result
        
    except Exception as e:
        raise ValueError(f"Erreur lecture Excel inventaire : {e}")


def _load_standard_excel(uploaded_file):
    """Charge Excel standard"""
    try:
        uploaded_file.seek(0)
        sheets = pd.read_excel(uploaded_file, sheet_name=None)
        
        if not sheets:
            raise ValueError("Aucune feuille trouvée")
        
        return sheets
        
    except Exception as e:
        raise ValueError(f"Erreur lecture Excel standard : {e}")


def process_multiple_files(uploaded_files, file_type, index_file=None):
    """
    Traite plusieurs fichiers avec gestion robuste des erreurs.
    """
    try:
        if not uploaded_files:
            raise ValueError("Aucun fichier fourni")
        
        all_results = []
        output = io.BytesIO()
        
        # Validation du type de fichier
        if file_type and file_type not in COLUMNS_BY_TYPE:
            raise ValueError(f"Type de fichier inconnu : {file_type}")

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for file_idx, uploaded_file in enumerate(uploaded_files):
                try:
                    # Validation du fichier individuel
                    if uploaded_file is None:
                        st.error(f"Fichier {file_idx + 1} est None")
                        continue
                    
                    uploaded_file.seek(0)
                    excel_data = load_data(uploaded_file, file_type)
                    
                    if not excel_data:
                        st.warning(f"Aucune donnée trouvée dans {uploaded_file.name}")
                        continue
                    
                    file_sheets_res = []
                    
                    for sheet_name, df in excel_data.items():
                        try:
                            # Validation du DataFrame
                            if df is None or df.empty:
                                st.warning(f"Feuille vide : {sheet_name}")
                                continue
                            
                            # Vérifier si c'est une feuille d'erreur
                            if "Erreur" in df.columns:
                                st.error(f"Erreur dans feuille {sheet_name} : {df['Erreur'].iloc[0]}")
                                continue
                            
                            df_filtered, missing_cols, exported_cols = filter_columns(df, sheet_name, file_type)
                            
                            if df_filtered.empty:
                                st.warning(f"Aucune colonne valide trouvée dans {sheet_name}")
                                continue
                            
                            extract_only = (
                                COLUMNS_BY_TYPE[file_type].get("extract_only", False)
                                if file_type in COLUMNS_BY_TYPE else False
                            )

                            if extract_only:
                                df_result = df_filtered.reset_index(drop=True)
                                unique_rows = len(df)
                                duplicate_rows = 0
                            else:
                                df_result = df_filtered.drop_duplicates()
                                unique_rows = len(df_result)
                                duplicate_rows = len(df) - unique_rows

                            # Validation du nom de feuille Excel
                            safe_name = _sanitize_sheet_name(uploaded_file.name.split('.')[0][:15])
                            sheet_export_name = f"{safe_name}_{_sanitize_sheet_name(sheet_name[:10])}"
                            
                            # Écriture avec validation
                            try:
                                df_result.to_excel(writer, sheet_name=sheet_export_name, index=False)
                            except Exception as e:
                                st.error(f"Erreur écriture feuille {sheet_export_name} : {e}")
                                continue

                            file_sheets_res.append({
                                "sheet_name": sheet_name,
                                "total_columns": len(df.columns),
                                "total_rows": len(df),
                                "unique_rows": unique_rows,
                                "duplicate_rows": duplicate_rows,
                                "extract_only": extract_only,
                                "df_result": df_result,
                            })
                            
                        except Exception as e:
                            st.error(f"Erreur traitement feuille {sheet_name} : {e}")
                            continue

                    all_results.append({"filename": uploaded_file.name, "sheets": file_sheets_res})
                    
                except Exception as e:
                    st.error(f"Erreur fichier {uploaded_file.name} : {e}")
                    continue

        # Validation finale
        if not all_results:
            raise ValueError("Aucun résultat généré")
        
        # Génération du fichier final
        output.seek(0)
        wb = load_workbook(output)
        build_report_sheet(wb, all_results)
        
        for sheet_name in wb.sheetnames:
            if sheet_name != "Rapport Global":
                try:
                    ws = wb[sheet_name]
                    _header_style(ws, ws.max_column)
                    apply_auto_width(ws)
                except Exception as e:
                    st.warning(f"Erreur style feuille {sheet_name} : {e}")

        final_output = io.BytesIO()
        wb.save(final_output)
        
        # Génération de l'index
        try:
            index_data = update_index_streamlit(all_results, file_type, index_file)
        except Exception as e:
            st.warning(f"Erreur génération index : {e}")
            index_data = None
        
        return final_output.getvalue(), index_data, all_results
        
    except Exception as e:
        st.error(f"Erreur globale traitement : {e}")
        # Retourner un résultat vide en cas d'erreur critique
        empty_output = io.BytesIO()
        empty_wb = Workbook()
        empty_wb.remove(empty_wb.active)
        empty_wb.save(empty_output)
        return empty_output.getvalue(), None, []


def _sanitize_sheet_name(name):
    """Nettoie le nom de feuille pour Excel"""
    import re
    # Remplacer les caractères invalides pour Excel
    sanitized = re.sub(r'[\\/*?:[\]]+', '_', name)
    sanitized = re.sub(r'\s+', '_', sanitized.strip())
    return sanitized[:31]  # Limite Excel de 31 caractères


# ──────────────────────────────────────────────
#  Interface Streamlit
# ──────────────────────────────────────────────
st.set_page_config(page_title="Excel & CSV Processor Pro", page_icon="📈", layout="wide")
st.markdown("""
<style>
.main { background-color: #f8f9fa; }
.stButton>button { border-radius: 8px; height: 3em; font-weight: bold; }
.step-box {
    background-color: #ffffff; padding: 20px; border-radius: 10px;
    border-left: 5px solid #2E4057; margin-bottom: 20px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}
.step-title { color: #2E4057; font-size: 1.2em; font-weight: bold; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

if 'results'            not in st.session_state: st.session_state.results            = None
if 'processed_data'     not in st.session_state: st.session_state.processed_data     = None
if 'index_data'         not in st.session_state: st.session_state.index_data         = None
if 'current_file_type'  not in st.session_state: st.session_state.current_file_type  = None

st.title("Excel & CSV Processor Pro")
st.markdown("Gestion par lots : **Achat, Client, Momo, Ristournes, Magasin, Stocks, Base Clients et Inventaire**.")

with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/microsoft-excel-2019.png", width=80)
    st.header("🛠️ Options")
    if st.button("Réinitialiser l'outil", use_container_width=True):
        st.session_state.results            = None
        st.session_state.processed_data     = None
        st.session_state.index_data         = None
        st.session_state.current_file_type  = None
        st.rerun()

st.markdown('<div class="step-box"><div class="step-title">Configuration du Traitement</div>', unsafe_allow_html=True)
col_cfg1, col_cfg2 = st.columns(2)
with col_cfg1:
    file_type = st.selectbox(
        "Type de données à traiter",
        options=[None] + list(COLUMNS_BY_TYPE.keys()),
        format_func=lambda x: "Traitement Standard" if x is None else x
    )
    if file_type:
        st.info(COLUMNS_BY_TYPE[file_type]["description"])

        # Afficher un rappel pour l'inventaire
        if file_type == "FICHIER DES INVENTAIRE":
            st.caption(
                "Format Eleader attendu : date en ligne 3, en-têtes en ligne 7, "
                "données à partir de la ligne 8. Les lignes sans opération "
                "(Qté Initiale = 0 et Qté Comptée = 0) sont automatiquement supprimées."
            )

with col_cfg2:
    index_file = st.file_uploader("📂 Index existant (Optionnel)", type=["xlsx"])
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="step-box"><div class="step-title">Chargement des Fichiers</div>', unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "Glissez vos fichiers ici (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True
)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_files:
    if st.button("🚀 Lancer le Traitement Groupé", use_container_width=True, type="primary"):
        with st.spinner("Calculs et analyse en cours..."):
            try:
                processed_data, index_data, all_results = process_multiple_files(
                    uploaded_files, file_type, index_file
                )
                st.session_state.processed_data    = processed_data
                st.session_state.index_data        = index_data
                st.session_state.results           = all_results
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
                    "Fichier Source": f_res["filename"],
                    "Feuille":        s_res["sheet_name"],
                    "Lignes":         s_res["total_rows"],
                    "Doublons":       f"❌ {s_res['duplicate_rows']}" if s_res['duplicate_rows'] > 0 else "✅ 0"
                })
        st.dataframe(pd.DataFrame(report_data), use_container_width=True, hide_index=True)

    with res_col2:
        st.markdown("### 📥 Téléchargements")
        type_suffix = (
            "Standard" if st.session_state.current_file_type is None
            else st.session_state.current_file_type.replace(" ", "_")
        )
        st.download_button(
            label="💾 Télécharger le fichier consolidé",
            data=st.session_state.processed_data,
            file_name=f"traitement_{type_suffix}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        if st.session_state.index_data:
            st.download_button(
                label=f"📂 Télécharger l'index {st.session_state.current_file_type or 'Standard'}",
                data=st.session_state.index_data,
                file_name=f"index_{type_suffix}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    st.markdown('</div>', unsafe_allow_html=True)
