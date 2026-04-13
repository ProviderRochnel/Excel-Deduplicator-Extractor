"""
DataHub Pro - Application de Traitement de Données Excel
Version optimisée avec structure modulaire et maintenabilité
"""

# ============================================================================
# IMPORTATIONS
# ============================================================================
import streamlit as st
import pandas as pd
import io
import os
import re
import numpy as np
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================================
# CONFIGURATION ET CONSTANTES
# ============================================================================

# Configuration des traitements par catégorie
COLUMNS_CONFIG = {
    "Facturation Achat": {
        "extract_only": False,
        "description": "Optimisation des achats : extraction des données clés et suppression automatique des doublons de facturation.",
        "tooltip": "Idéal pour consolider vos bons de commande et éviter les doubles paiements.",
        "columns": ["Date", "Products", "Quantity ordered", "Price", "Order number"],
        "rename": {"Price": "Prix Achat HT", "Order number": "N° Facture"},
    },
    "Facturation Client": {
        "extract_only": True,
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
        "description": "Traitement des flux mobiles (Momo) : conversion CSV/Excel et normalisation des transactions.",
        "tooltip": "Traitez vos rapports d'opérations mobiles en un clic.",
        "columns": ["Id", "Date", "Status", "Type", "From", "To name", "Amount", "Balance"],
        "rename": {"Id": "N° Identification", "From": "Provenance", "To name": "To handler name"},
        "extra_columns": ["Vendeur", "Compte", "Tournée"],
    },
    "FICHIER DES RISTOURNES": {
        "extract_only": True,
        "description": "Calcul des ristournes : récapitulatif structuré des remises quotidiennes par magasin.",
        "tooltip": "Générez vos états de remises sur ventes en quelques secondes.",
        "columns": ["Code", "Store name", "Start date", "Document number", "Product code", "Product name", "Family", "CODE RABATES", "RISTOURNE"],
        "rename": {},
        "extra_columns": [],
    },
    "BASE MAGASIN": {
        "extract_only": False,
        "description": "Mouvements de stock : analyse fine des opérations magasin (Chargement, Déchargement, etc.) et suppression automatique des doublons.",
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
        "description": "Inventaire financier : valorisation des stocks avec arrondis comptables et traduction de nature.",
        "tooltip": "Valorisez vos stocks au PA HT avec arrondis automatiques.",
        "columns": ["Code", "Name", "Category", "Price", "Qty", "Totalvalue"],
        "rename": {"Code": "Code", "Name": "Nom du Produit", "Category": "Nature", "Price": "Prix Achat HT", "Qty": "Qty", "Totalvalue": "Valeur au PA HT"},
        "extra_columns": ["Conditionnement", "PA TTC", "Valeur au PA TTC", "Prix Vente HT", "Prix de Vente 2%", "Valeur au PV TTC"],
    },
    "BASE CLIENTS": {
        "extract_only": True,
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
        "description": "Clôture journalière : nettoyage des stocks nuls et calcul des écarts d'inventaire.",
        "tooltip": "Supprime automatiquement les lignes sans stock pour un inventaire clair.",
        "columns": ["Date", "Code Article", "Nom Article", "Conditionnement", "Quantité Initiale", "Quantité Comptée", "Valeur Ecart"],
        "rename": {},
        "extra_columns": ["Valeur Stock PV TTC 2%"],
    }
}

# Configuration des tâches quotidiennes
DAILY_TASKS_CONFIG = {
    "Inventaire": {
        "category": "FICHIER DES INVENTAIRES",
        "icon": "",
        "description": "Clôture journalière des stocks",
        "priority": 1
    },
    "Données Magasin": {
        "category": "BASE MAGASIN", 
        "icon": "",
        "description": "Mouvements de stock du jour",
        "priority": 2
    },
    "Ristourne": {
        "category": "FICHIER DES RISTOURNES",
        "icon": "", 
        "description": "Calcul des ristournes quotidiennes",
        "priority": 3
    },
    "Vente": {
        "category": "Facturation Client",
        "icon": "",
        "description": "Facturation clients du jour", 
        "priority": 4
    },
    "Achat": {
        "category": "Facturation Achat",
        "icon": "",
        "description": "Facturation achats du jour",
        "priority": 5
    }
}

# Constantes système
INDEX_SHEET_NAME = "Données"
TRACE_COLS = ["Fichier source", "Date d'import"]
BASE_MAGASIN_OPS = {"LR": "Chargement", "UC": "Déchargement", "DO": "Sorties directes de stock", "DI": "Entrées Directes de stock", "BO": "Achat ou Commande Appro", "DE": "Retour Emballages SABC"}
BASE_MAGASIN_COMPTES = {"S02004": "R1", "S02003": "R2", "S02005": "R3", "S02006": "R4"}
STOCKS_NATURES = {"BEER": "BIÈRES", "BG": "BOISSONS RAFRAICHISSANTE SANS ALCOOL", "AM": "ALCOOL MIX", "EAU": "EAUX MINERALES NATURELLES", "EMB": "PACKAGES"}

# Styles et couleurs
PRIMARY_COLOR = "2E4057"
SUCCESS_COLOR = "28a745"
WARNING_COLOR = "ffc107"
DANGER_COLOR = "dc3545"
LIGHT_COLOR = "f8f9fa"


# ============================================================================
# UTILITAIRES SYSTÈME
# ============================================================================

class FileManager:
    """Gestion centralisée des opérations sur fichiers"""
    
    @staticmethod
    def get_index_folder() -> Path:
        """Crée et retourne le chemin du dossier DataHub_Index dans le dossier utilisateur"""
        user_folder = Path.home()
        index_folder = user_folder / "DataHub_Index"
        index_folder.mkdir(exist_ok=True)
        return index_folder
    
    @staticmethod
    def save_index_locally(file_content: bytes, filename: str) -> bool:
        """Sauvegarde le contenu de l'index dans un fichier"""
        try:
            index_folder = FileManager.get_index_folder()
            file_path = index_folder / filename
            with open(file_path, "wb") as f:
                f.write(file_content)
            return True
        except Exception as e:
            st.error(f"Erreur lors de la sauvegarde : {e}")
            return False
    
    @staticmethod
    def get_local_index(filename: str) -> Optional[bytes]:
        """Récupère le contenu d'un index depuis le dossier utilisateur"""
        try:
            index_folder = FileManager.get_index_folder()
            file_path = index_folder / filename
            if file_path.exists():
                with open(file_path, "rb") as f:
                    return f.read()
            return None
        except Exception as e:
            st.error(f"Erreur lors de la lecture : {e}")
            return None
    
    @staticmethod
    def list_local_indexes() -> List[str]:
        """Liste tous les index disponibles dans le dossier utilisateur"""
        try:
            index_folder = FileManager.get_index_folder()
            return [f.name for f in index_folder.glob("*.xlsx")]
        except Exception as e:
            st.error(f"Erreur lors de la liste des fichiers : {e}")
            return []


class DailyTaskManager:
    """Gestion des tâches quotidiennes obligatoires"""
    
    @staticmethod
    def get_task_file() -> Path:
        """Retourne le chemin du fichier de suivi des tâches quotidiennes"""
        return FileManager.get_index_folder() / "daily_tasks.json"
    
    @staticmethod
    def load_tasks() -> Dict[str, Any]:
        """Charge les tâches quotidiennes depuis le fichier JSON"""
        try:
            task_file = DailyTaskManager.get_task_file()
            if task_file.exists():
                with open(task_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception:
            return {}
    
    @staticmethod
    def save_tasks(tasks_data: Dict[str, Any]) -> bool:
        """Sauvegarde les tâches quotidiennes dans le fichier JSON"""
        try:
            task_file = DailyTaskManager.get_task_file()
            with open(task_file, 'w', encoding='utf-8') as f:
                json.dump(tasks_data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            st.error(f"Erreur sauvegarde tâches : {e}")
            return False
    
    @staticmethod
    def get_today_key() -> str:
        """Retourne la clé pour aujourd'hui (format YYYY-MM-DD)"""
        return datetime.now().strftime("%Y-%m-%d")
    
    @staticmethod
    def is_task_completed(task_name: str) -> bool:
        """Vérifie si une tâche est complétée aujourd'hui"""
        tasks_data = DailyTaskManager.load_tasks()
        today_key = DailyTaskManager.get_today_key()
        return tasks_data.get(today_key, {}).get(task_name, False)
    
    @staticmethod
    def mark_task_completed(task_name: str) -> bool:
        """Marque une tâche comme complétée aujourd'hui"""
        tasks_data = DailyTaskManager.load_tasks()
        today_key = DailyTaskManager.get_today_key()
        
        if today_key not in tasks_data:
            tasks_data[today_key] = {}
        
        tasks_data[today_key][task_name] = True
        return DailyTaskManager.save_tasks(tasks_data)
    
    @staticmethod
    def get_progress() -> Tuple[int, int]:
        """Retourne le progrès quotidien (tâches complétées / total)"""
        completed = sum(1 for task_name in DAILY_TASKS_CONFIG.keys() 
                       if DailyTaskManager.is_task_completed(task_name))
        total = len(DAILY_TASKS_CONFIG)
        return completed, total


class ExcelStyler:
    """Gestion des styles Excel"""
    
    @staticmethod
    def apply_header_style(ws, nb_cols: int, row: int = 1) -> None:
        """Applique le style d'en-tête"""
        for col_idx in range(1, nb_cols + 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill("solid", start_color=PRIMARY_COLOR, end_color=PRIMARY_COLOR)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
    @staticmethod
    def get_border() -> Border:
        """Retourne un bordure standard"""
        thin = Side(style="thin", color="BFBFBF")
        return Border(left=thin, right=thin, top=thin, bottom=thin)
    
    @staticmethod
    def apply_auto_width(ws) -> None:
        """Ajuste automatiquement la largeur des colonnes"""
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row in range(1, ws.max_row + 1):
                val = ws.cell(row=row, column=col_idx).value
                if val:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[col_letter].width = min(max_len + 4, 50)


# ============================================================================
# LOGIQUE DE TRAITEMENT DES DONNÉES
# ============================================================================

class DataProcessor:
    """Classe principale pour le traitement des données"""
    
    @staticmethod
    def filter_columns(df: pd.DataFrame, sheet_name: str, file_type: Optional[str]) -> Tuple[pd.DataFrame, List[str], List[str]]:
        """Filtre et transforme les colonnes selon le type de fichier"""
        if file_type is None or file_type not in COLUMNS_CONFIG:
            return df, [], list(df.columns)
        
        config = COLUMNS_CONFIG[file_type]
        expected = config["columns"]
        rename_map = config.get("rename", {})
        
        # Cas spécial pour STOCKS DES PRODUITS
        if file_type == "STOCKS DES PRODUITS" and "Code" not in df.columns and len(df.columns) >= 7:
            df["Code"] = df.iloc[:, 6]
        
        # Colonnes disponibles
        available = [c for c in expected if c in df.columns]
        missing = [c for c in expected if c not in df.columns]
        df_out = df[available].copy()
        
        # Traitements spécifiques par type
        if file_type == "FICHIER DES INVENTAIRES":
            df_out = DataProcessor._process_inventaires(df_out)
        elif file_type == "BASE MAGASIN":
            df_out = DataProcessor._process_base_magasin(df_out)
        elif file_type == "STOCKS DES PRODUITS":
            df_out = DataProcessor._process_stocks(df_out)
        
        # Renommage et colonnes supplémentaires
        df_out = df_out.rename(columns=rename_map)
        for extra_col in config.get("extra_columns", []):
            if extra_col not in df_out.columns:
                df_out[extra_col] = ""
        
        return df_out, missing, list(df_out.columns)
    
    @staticmethod
    def _process_inventaires(df: pd.DataFrame) -> pd.DataFrame:
        """Traitement spécifique pour les inventaires"""
        if "Quantité Initiale" in df.columns and "Quantité Comptée" in df.columns:
            return df[~((df["Quantité Initiale"] == 0) & (df["Quantité Comptée"] == 0))]
        return df
    
    @staticmethod
    def _process_base_magasin(df: pd.DataFrame) -> pd.DataFrame:
        """Traitement spécifique pour la base magasin"""
        if "Position record date" in df.columns:
            df["Position record date"] = pd.to_datetime(df["Position record date"], errors='coerce').dt.strftime('%d/%m/%Y')
        
        if "Document" in df.columns:
            def extract_op_compte(doc_str):
                doc_str = str(doc_str)
                op_val = ""
                compte_val = ""
                for code, label in BASE_MAGASIN_OPS.items():
                    if code in doc_str:
                        op_val = label
                        break
                for code, label in BASE_MAGASIN_COMPTES.items():
                    if code in doc_str:
                        compte_val = label
                        break
                return pd.Series([op_val, compte_val])
            
            df[["Document", "Compte"]] = df["Document"].apply(extract_op_compte)
        else:
            df["Compte"] = ""
        
        # Suppression des doublons sur les colonnes métier clés
        dedup_cols = [c for c in ["Product", "Change date", "Prev. quantity", "Cur. quantity", "Change source", "Document"] if c in df.columns]
        if dedup_cols:
            df = df.drop_duplicates(subset=dedup_cols).reset_index(drop=True)
        
        return df
    
    @staticmethod
    def _process_stocks(df: pd.DataFrame) -> pd.DataFrame:
        """Traitement spécifique pour les stocks"""
        if "Category" in df.columns:
            df["Category"] = df["Category"].map(STOCKS_NATURES).fillna(df["Category"])
        
        for col in ["Price", "Totalvalue"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').apply(lambda x: np.ceil(x) if pd.notnull(x) else x)
        
        return df
    
    @staticmethod
    def load_data(uploaded_file) -> Dict[str, pd.DataFrame]:
        """Charge les données depuis un fichier uploadé"""
        filename = uploaded_file.name
        if filename.endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file, sep=None, engine='python')
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep=',')
            return {"Données_CSV": df}
        else:
            return pd.read_excel(uploaded_file, sheet_name=None)
    
    @staticmethod
    def process_multiple_files(uploaded_files: List, file_type: Optional[str]) -> Tuple[bytes, List[Dict]]:
        """Traite plusieurs fichiers et génère le rapport"""
        all_results = []
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for uploaded_file in uploaded_files:
                uploaded_file.seek(0)
                excel_data = DataProcessor.load_data(uploaded_file)
                file_sheets_res = []
                
                for sheet_name, df in excel_data.items():
                    df_filtered, missing_cols, exported_cols = DataProcessor.filter_columns(df, sheet_name, file_type)
                    
                    if file_type in COLUMNS_CONFIG:
                        extract_only = COLUMNS_CONFIG[file_type].get("extract_only", False)
                    else:
                        extract_only = False
                    
                    if extract_only:
                        df_result = df_filtered.reset_index(drop=True)
                        unique_rows, duplicate_rows = len(df_result), 0
                    elif file_type == "BASE MAGASIN":
                        # Le dédoublonnage est déjà effectué dans _process_base_magasin
                        df_result = df_filtered.reset_index(drop=True)
                        unique_rows = len(df_result)
                        duplicate_rows = len(df) - unique_rows
                    else:
                        df_result = df_filtered.drop_duplicates()
                        unique_rows = len(df_result)
                        duplicate_rows = len(df) - unique_rows
                    
                    safe_name = uploaded_file.name.split('.')[0][:15]
                    sheet_export_name = f"{safe_name}_{sheet_name[:10]}".replace(' ', '_').replace('.', '_')
                    df_result.to_excel(writer, sheet_name=sheet_export_name, index=False)
                    
                    file_sheets_res.append({
                        "sheet_name": sheet_name,
                        "total_columns": len(df.columns),
                        "total_rows": len(df),
                        "unique_rows": unique_rows,
                        "duplicate_rows": duplicate_rows,
                        "extract_only": extract_only,
                        "df_result": df_result
                    })
                
                all_results.append({"filename": uploaded_file.name, "sheets": file_sheets_res})
        
        # Création du rapport global
        output.seek(0)
        wb = load_workbook(output)
        ws_rep = wb.create_sheet("Rapport Global", 0)
        
        headers = ["Fichier", "Feuille", "Cols", "Lignes Tot.", "Uniques", "Doublons", "Mode"]
        for i, h in enumerate(headers, 1):
            cell = ws_rep.cell(row=1, column=i, value=h)
            cell.fill = PatternFill("solid", start_color=PRIMARY_COLOR, end_color=PRIMARY_COLOR)
            cell.font = Font(name="Arial", bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")
        
        curr_row = 2
        for file_res in all_results:
            for r in file_res["sheets"]:
                vals = [
                    file_res["filename"], r["sheet_name"], r["total_columns"], 
                    r["total_rows"], r["unique_rows"], r["duplicate_rows"], 
                    "Extraction" if r["extract_only"] else "Dédoublonnage"
                ]
                for c, v in enumerate(vals, 1):
                    cell = ws_rep.cell(row=curr_row, column=c, value=v)
                    cell.border = ExcelStyler.get_border()
                curr_row += 1
        
        ExcelStyler.apply_auto_width(ws_rep)
        
        # Appliquer les styles à toutes les feuilles
        for sheet_name in wb.sheetnames:
            if sheet_name != "Rapport Global":
                ws = wb[sheet_name]
                ExcelStyler.apply_header_style(ws, ws.max_column)
                ExcelStyler.apply_auto_width(ws)
        
        final_output = io.BytesIO()
        wb.save(final_output)
        return final_output.getvalue(), all_results


class IndexManager:
    """Gestion des index centralisés"""
    
    @staticmethod
    def update_index_content(all_results: List[Dict], existing_index_content: Optional[bytes] = None) -> Tuple[Optional[bytes], int, int]:
        """
        Fusionne les nouveaux résultats dans l'index existant.
        Règle : une ligne est ignorée si son contenu (hors colonnes de traçabilité)
        existe déjà dans l'index.
        """
        now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        # Lire l'index existant
        df_existing = None
        if existing_index_content:
            try:
                df_existing = pd.read_excel(
                    io.BytesIO(existing_index_content),
                    sheet_name=INDEX_SHEET_NAME,
                    dtype=str
                )
            except Exception:
                df_existing = None
        
        # Construire le DataFrame des nouvelles lignes
        frames_new = []
        for file_res in all_results:
            for r in file_res["sheets"]:
                df = r["df_result"].copy().astype(str)
                df.insert(0, TRACE_COLS[0], file_res["filename"])
                df.insert(1, TRACE_COLS[1], now_str)
                frames_new.append(df)
        
        if not frames_new:
            return None, 0, 0
        
        df_new = pd.concat(frames_new, ignore_index=True)
        
        # Déduplication par contenu (hors colonnes de traçabilité)
        data_cols = [c for c in df_new.columns if c not in TRACE_COLS]
        added_rows = 0
        skipped_rows = 0
        
        if df_existing is not None and not df_existing.empty:
            # Aligner les colonnes
            all_cols = list(dict.fromkeys(list(df_existing.columns) + list(df_new.columns)))
            df_existing = df_existing.reindex(columns=all_cols).astype(str)
            df_new = df_new.reindex(columns=all_cols).astype(str)
            
            # Colonnes de comparaison
            compare_cols = [c for c in data_cols if c in df_existing.columns]
            
            # Créer une clé de comparaison
            existing_keys = set(
                df_existing[compare_cols].fillna("").apply(
                    lambda row: "||".join(row.values), axis=1
                )
            )
            
            def is_new_row(row):
                key = "||".join(row[compare_cols].fillna("").values)
                return key not in existing_keys
            
            mask_new = df_new.apply(is_new_row, axis=1)
            skipped_rows = (~mask_new).sum()
            added_rows = mask_new.sum()
            df_to_add = df_new[mask_new]
            
            df_final = pd.concat([df_existing, df_to_add], ignore_index=True)
        else:
            df_final = df_new
            added_rows = len(df_new)
        
        if added_rows == 0:
            return None, skipped_rows, added_rows
        
        # Écrire dans un classeur Excel formaté
        wb_idx = Workbook()
        ws_idx = wb_idx.active
        ws_idx.title = INDEX_SHEET_NAME
        
        for c, col in enumerate(df_final.columns, 1):
            ws_idx.cell(row=1, column=c, value=col)
        ExcelStyler.apply_header_style(ws_idx, len(df_final.columns))
        
        for r_idx, row in enumerate(df_final.itertuples(index=False), start=2):
            for c_idx, val in enumerate(row, 1):
                cell_val = "" if (isinstance(val, float) and np.isnan(val)) else val
                cell = ws_idx.cell(row=r_idx, column=c_idx, value=cell_val)
                cell.border = ExcelStyler.get_border()
                if c_idx <= 2:
                    cell.font = Font(name="Arial", size=9, italic=True, color="555555")
                    cell.fill = PatternFill("solid", start_color="E8F5E9", end_color="E8F5E9")
        
        ExcelStyler.apply_auto_width(ws_idx)
        output = io.BytesIO()
        wb_idx.save(output)
        return output.getvalue(), skipped_rows, added_rows


# ============================================================================
# INTERFACE UTILISATEUR
# ============================================================================

class UIComponents:
    """Composants UI réutilisables"""
    
    @staticmethod
    def setup_page_config() -> None:
        """Configure la page Streamlit"""
        st.set_page_config(
            page_title="DataHub Pro | Traitement Local", 
            layout="wide"
        )
    
    @staticmethod
    def apply_custom_styles() -> None:
        """Applique les styles CSS personnalisés"""
        st.markdown("""
        <style>
            .main { background-color: #f0f2f6; }
            .stApp { background-color: #f0f2f6; }
            .stButton>button { 
                border-radius: 12px; 
                height: 3.5em; 
                font-weight: 700; 
                transition: all 0.3s; 
            }
            .stButton>button:hover { 
                transform: translateY(-2px); 
                box-shadow: 0 4px 12px rgba(0,0,0,0.1); 
            }
            .step-card { 
                background-color: #ffffff; 
                padding: 25px; 
                border-radius: 15px; 
                margin-bottom: 25px; 
                box-shadow: 0 4px 6px rgba(0,0,0,0.05); 
                border-top: 5px solid #2E4057; 
            }
            .step-header { 
                display: flex; 
                align-items: center; 
                margin-bottom: 15px; 
            }
            .step-number { 
                background-color: #2E4057; 
                color: white; 
                border-radius: 50%; 
                width: 30px; 
                height: 30px; 
                display: flex; 
                align-items: center; 
                justify-content: center; 
                font-weight: bold; 
                margin-right: 15px; 
            }
            .step-title { 
                color: #2E4057; 
                font-size: 1.3em; 
                font-weight: bold; 
            }
            .stat-card { 
                background-color: #f8f9fa; 
                padding: 15px; 
                border-radius: 10px; 
                text-align: center; 
                border: 1px solid #e9ecef; 
            }
            .stat-val { 
                font-size: 1.8em; 
                font-weight: bold; 
                color: #2E4057; 
            }
            .stat-label { 
                font-size: 0.9em; 
                color: #6c757d; 
            }
            .footer { 
                text-align: center; 
                color: #adb5bd; 
                padding: 40px 0; 
                font-size: 0.85em; 
            }
        </style>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def render_sidebar() -> None:
        """Affiche la barre latérale"""
        with st.sidebar:
            st.markdown("### Centre de Contrôle")
            
            # Navigation
            st.markdown("---")
            if st.button("Traitement des fichiers", use_container_width=True):
                st.session_state.page = "Traitement"
                st.rerun()
            if st.button("Tâches Quotidiennes", use_container_width=True):
                st.session_state.page = "Taches"
                st.rerun()
            if st.button("Bibliothèque des Index", use_container_width=True):
                st.session_state.page = "Index"
                st.rerun()
            
            # Information stockage
            st.markdown("---")
            st.markdown("### Stockage Utilisateur")
            
            try:
                index_folder = FileManager.get_index_folder()
                st.success(f"Dossier prêt : `{index_folder.name}`")
                st.caption(f"Chemin complet : {index_folder}")
            except Exception as e:
                st.error(f"Erreur dossier : {e}")
            
            st.info("Les index sont sauvegardés dans le dossier **DataHub_Index** de votre répertoire utilisateur.")
            
            if st.button("Réinitialiser la session", use_container_width=True):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()


class DataHubApp:
    """Application principale DataHub"""
    
    def __init__(self):
        self.processor = DataProcessor()
        self.index_manager = IndexManager()
        self.task_manager = DailyTaskManager()
        self.file_manager = FileManager()
        self.ui = UIComponents()
    
    def initialize_session(self) -> None:
        """Initialise l'état de la session"""
        if 'results' not in st.session_state:
            st.session_state.results = None
        if 'processed_data' not in st.session_state:
            st.session_state.processed_data = None
        if 'current_file_type' not in st.session_state:
            st.session_state.current_file_type = None
        if 'page' not in st.session_state:
            st.session_state.page = "Traitement"
        
        # Création proactive du dossier
        try:
            FileManager.get_index_folder()
        except Exception as e:
            st.error(f"Erreur lors de la création du dossier de stockage : {e}")
    
    def render_processing_page(self) -> None:
        """Affiche la page de traitement"""
        st.markdown("<h1>Traitement Intelligent</h1>", unsafe_allow_html=True)
        
        col_main1, col_main2 = st.columns([1, 1])
        
        with col_main1:
            st.markdown("""
            <div class="step-card">
                <div class="step-header">
                    <div class="step-number">1</div>
                    <div class="step-title">Configuration</div>
                </div>
            """, unsafe_allow_html=True)
            
            file_type = st.selectbox(
                "Module métier", 
                options=[None] + list(COLUMNS_CONFIG.keys()),
                format_func=lambda x: "Traitement Standard" if x is None else x
            )
            
            if file_type:
                st.markdown(f"*{COLUMNS_CONFIG[file_type]['description']}*")
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col_main2:
            st.markdown("""
            <div class="step-card">
                <div class="step-header">
                    <div class="step-number">2</div>
                    <div class="step-title">Source de Données</div>
                </div>
            """, unsafe_allow_html=True)
            
            uploaded_files = st.file_uploader(
                "Glissez vos fichiers Excel ou CSV ici", 
                type=["xlsx", "csv"], 
                accept_multiple_files=True
            )
            
            if uploaded_files:
                if st.button("Lancer l'analyse", use_container_width=True, type="primary"):
                    with st.spinner("Traitement en cours..."):
                        try:
                            p_data, results = self.processor.process_multiple_files(uploaded_files, file_type)
                            st.session_state.processed_data = p_data
                            st.session_state.results = results
                            st.session_state.current_file_type = file_type
                            
                            # Marquer automatiquement la tâche quotidienne
                            if file_type:
                                for task_name, task_config in DAILY_TASKS_CONFIG.items():
                                    if task_config["category"] == file_type:
                                        if self.task_manager.mark_task_completed(task_name):
                                            st.success(f"Tâche quotidienne '{task_name}' marquée comme complétée !")
                                        break
                            
                            st.balloons()
                        except Exception as e:
                            st.error(f"Erreur : {str(e)}")
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Affichage des résultats
        if st.session_state.results:
            self._render_results_section()
    
    def _render_results_section(self) -> None:
        """Affiche la section des résultats"""
        st.markdown("""
        <div class="step-card">
            <div class="step-header">
                <div class="step-number">3</div>
                <div class="step-title">Tableau de Bord & Actions</div>
            </div>
        """, unsafe_allow_html=True)
        
        # Statistiques
        total_in = sum(sum(s['total_rows'] for s in f['sheets']) for f in st.session_state.results)
        total_out = sum(sum(s['unique_rows'] for s in f['sheets']) for f in st.session_state.results)
        
        s1, s2, s3 = st.columns(3)
        with s1:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-val">{total_in}</div>
                <div class="stat-label">Lignes Entrantes</div>
            </div>
            """, unsafe_allow_html=True)
        
        with s2:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-val" style="color:{SUCCESS_COLOR};">{total_out}</div>
                <div class="stat-label">Lignes Uniques</div>
            </div>
            """, unsafe_allow_html=True)
        
        with s3:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-val" style="color:{DANGER_COLOR};">{total_in - total_out}</div>
                <div class="stat-label">Doublons Éliminés</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        c_left, c_right = st.columns([2, 1])
        
        with c_left:
            st.markdown("### Aperçu")
            tab_data = [
                {"Fichier": f["filename"], "Uniques": sum(s['unique_rows'] for s in f['sheets'])} 
                for f in st.session_state.results
            ]
            st.dataframe(pd.DataFrame(tab_data), use_container_width=True, hide_index=True)
        
        with c_right:
            st.markdown("### Actions")
            suffix = "Standard" if st.session_state.current_file_type is None else st.session_state.current_file_type.replace(" ", "_")
            st.download_button(
                "Exporter Fichier Traité", 
                data=st.session_state.processed_data, 
                file_name=f"Export_{suffix}.xlsx", 
                use_container_width=True
            )
            
            st.markdown("---")
            st.markdown("#### Centralisation Locale")
            if st.button("Fusionner avec l'Index Local", use_container_width=True, type="secondary"):
                with st.spinner("Fusion de l'index local..."):
                    idx_filename = f"index_{suffix}.xlsx"
                    existing_content = self.file_manager.get_local_index(idx_filename)
                    new_index_content, skipped, added = self.index_manager.update_index_content(
                        st.session_state.results, existing_content
                    )
                    
                    if new_index_content is None:
                        st.warning(f"Aucune nouvelle ligne à ajouter - {skipped} ligne(s) déjà présentes.")
                    elif self.file_manager.save_index_locally(new_index_content, idx_filename):
                        st.success(f"Fusion réussie : **{added}** nouvelle(s) ligne(s) ajoutée(s), **{skipped}** doublon(s) ignoré(s).")
                    else:
                        st.error("Erreur lors de la sauvegarde locale.")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    def render_daily_tasks_page(self) -> None:
        """Affiche la page des tâches quotidiennes"""
        st.markdown("<h1>Tâches Quotidiennes Obligatoires</h1>", unsafe_allow_html=True)
        
        # Afficher la date et le progrès
        today_str = datetime.now().strftime("%d/%m/%Y")
        completed, total = self.task_manager.get_progress()
        progress_percent = (completed / total) * 100 if total > 0 else 0
        
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.markdown(f"### Date : {today_str}")
        with col2:
            st.markdown(f"### {completed}/{total}")
        with col3:
            st.markdown(f"### {progress_percent:.0f}%")
        
        st.progress(progress_percent / 100)
        st.markdown("---")
        
        # Afficher les tâches
        st.markdown("## Liste des Tâches")
        
        for task_name, task_config in sorted(DAILY_TASKS_CONFIG.items(), key=lambda x: x[1]["priority"]):
            is_completed = self.task_manager.is_task_completed(task_name)
            
            col_task, col_status = st.columns([4, 1])
            
            with col_task:
                if is_completed:
                    st.markdown(f"""
                    <div style="background-color: #d4edda; border-left: 4px solid {SUCCESS_COLOR}; padding: 15px; border-radius: 5px; margin: 10px 0;">
                        <h4 style="margin: 0; color: #155724;">
                            {task_config['icon']} <s>{task_name}</s>
                        </h4>
                        <p style="margin: 5px 0 0 0; color: #155724; font-size: 0.9em;">
                            {task_config['description']} ? <strong>Catégorie: {task_config['category']}</strong>
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div style="background-color: {LIGHT_COLOR}; border-left: 4px solid #6c757d; padding: 15px; border-radius: 5px; margin: 10px 0;">
                        <h4 style="margin: 0; color: #495057;">
                            {task_config['icon']} {task_name}
                        </h4>
                        <p style="margin: 5px 0 0 0; color: #6c757d; font-size: 0.9em;">
                            {task_config['description']} ? <strong>Catégorie: {task_config['category']}</strong>
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
            
            with col_status:
                if is_completed:
                    st.success("✅")
                else:
                    st.warning("⚠️")
        
        st.markdown("---")
        
        # Statistiques et instructions
        col_info1, col_info2 = st.columns(2)
        
        with col_info1:
            st.markdown("### Statistiques")
            if completed == total:
                st.success("Toutes les tâches du jour sont complétées !")
            else:
                remaining = total - completed
                st.info(f"Il reste {remaining} tâche(s) à compléter aujourd'hui.")
        
        with col_info2:
            st.markdown("### Instructions")
            st.markdown("""
            - Traitez les fichiers dans la section **Traitement**
            - Sélectionnez la catégorie correspondante
            - La tâche se coche automatiquement
            - Le suivi est sauvegardé quotidiennement
            """)
        
        # Bouton de navigation
        if completed < total:
            st.markdown("---")
            if st.button("Aller au Traitement", use_container_width=True, type="primary"):
                st.session_state.page = "Traitement"
                st.rerun()
    
    def render_index_page(self) -> None:
        """Affiche la page de la bibliothèque d'index"""
        st.markdown("<h1>Bibliothèque des Index (Utilisateur)</h1>", unsafe_allow_html=True)
        
        try:
            index_folder = self.file_manager.get_index_folder()
            st.success(f"Dossier de stockage prêt : `{index_folder}`")
            st.caption(f"Chemin complet : {index_folder}")
            
            indexes = self.file_manager.list_local_indexes()
        except Exception as e:
            st.error(f"Impossible d'accéder au dossier de stockage : {e}")
            indexes = []
        
        if not indexes:
            st.warning("Aucun index trouvé dans votre dossier utilisateur. Effectuez un traitement et fusionnez les données pour créer votre premier index.")
        else:
            st.markdown("Retrouvez ici tous vos index sauvegardés dans votre dossier utilisateur. Vous pouvez les télécharger.")
            
            for idx_name in sorted(indexes):
                with st.expander(f"{idx_name}", expanded=False):
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.markdown(f"**Nom du fichier** : `{idx_name}`")
                        st.markdown(f"**Catégorie** : {idx_name.replace('index_', '').replace('.xlsx', '').replace('_', ' ')}")
                    with col2:
                        content = self.file_manager.get_local_index(idx_name)
                        if content:
                            st.download_button(
                                f"Télécharger", 
                                data=content, 
                                file_name=idx_name, 
                                key=idx_name, 
                                use_container_width=True
                            )
    
    def run(self) -> None:
        """Point d'entrée principal de l'application"""
        # Configuration
        self.ui.setup_page_config()
        self.ui.apply_custom_styles()
        
        # Initialisation
        self.initialize_session()
        
        # Navigation
        self.ui.render_sidebar()
        
        # Routage des pages
        if st.session_state.page == "Traitement":
            self.render_processing_page()
        elif st.session_state.page == "Taches":
            self.render_daily_tasks_page()
        elif st.session_state.page == "Index":
            self.render_index_page()
        
        # Footer
        st.markdown('<div class="footer">DataHub Pro v3.0 | Stockage Local | © 2024</div>', unsafe_allow_html=True)


# ============================================================================
# POINT D'ENTRÉE
# ============================================================================

if __name__ == "__main__":
    app = DataHubApp()
    app.run()
