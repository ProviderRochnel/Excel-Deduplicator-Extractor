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

# Configuration des tâches quotidiennes avec icônes améliorées
DAILY_TASKS_CONFIG = {
    "Inventaire": {
        "category": "FICHIER DES INVENTAIRES",
        "icon": "📋",
        "description": "Clôture journalière des stocks",
        "priority": 1
    },
    "Données Magasin": {
        "category": "BASE MAGASIN", 
        "icon": "🏪",
        "description": "Mouvements de stock du jour",
        "priority": 2
    },
    "Ristourne": {
        "category": "FICHIER DES RISTOURNES",
        "icon": "💰", 
        "description": "Calcul des ristournes quotidiennes",
        "priority": 3
    },
    "Vente": {
        "category": "Facturation Client",
        "icon": "🤝",
        "description": "Facturation clients du jour", 
        "priority": 4
    },
    "Achat": {
        "category": "Facturation Achat",
        "icon": "🛒",
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
PRIMARY_COLOR = "#2E4057"
SUCCESS_COLOR = "#28a745"
WARNING_COLOR = "#ffc107"
DANGER_COLOR = "#dc3545"
LIGHT_COLOR = "#f8f9fa"


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
        
        # Vérifier si le dossier existe déjà
        if index_folder.exists():
            if index_folder.is_dir():
                # Le dossier existe déjà, ne rien faire
                pass
            else:
                # Un fichier existe avec ce nom, le supprimer
                index_folder.unlink(missing_ok=True)
                index_folder.mkdir(exist_ok=True)
        else:
            # Créer le dossier
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
    
    @staticmethod
    def delete_index(filename: str) -> bool:
        """Supprime un index spécifique du dossier utilisateur"""
        try:
            index_folder = FileManager.get_index_folder()
            file_path = index_folder / filename
            if file_path.exists():
                file_path.unlink()
                return True
            return False
        except Exception as e:
            st.error(f"Erreur lors de la suppression : {e}")
            return False


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
            cell.fill = PatternFill("solid", start_color=PRIMARY_COLOR.replace("#", ""))
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
                file_results = {"filename": uploaded_file.name, "sheets": []}
                try:
                    sheets_dict = DataProcessor.load_data(uploaded_file)
                    for sheet_name, df in sheets_dict.items():
                        total_rows = len(df)
                        df_proc, missing, final_cols = DataProcessor.filter_columns(df, sheet_name, file_type)
                        
                        # Ajout des colonnes de traçabilité
                        df_proc[TRACE_COLS[0]] = uploaded_file.name
                        df_proc[TRACE_COLS[1]] = datetime.now().strftime("%d/%m/%Y %H:%M")
                        
                        unique_rows = len(df_proc)
                        
                        # Écriture dans Excel
                        safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', f"{uploaded_file.name[:20]}_{sheet_name[:10]}")
                        df_proc.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                        
                        # Application des styles
                        ws = writer.sheets[safe_sheet_name]
                        ExcelStyler.apply_header_style(ws, len(df_proc.columns))
                        ExcelStyler.apply_auto_width(ws)
                        
                        file_results["sheets"].append({
                            "name": sheet_name,
                            "total_rows": total_rows,
                            "unique_rows": unique_rows,
                            "missing_cols": missing
                        })
                except Exception as e:
                    file_results["error"] = str(e)
                
                all_results.append(file_results)
        
        return output.getvalue(), all_results


class IndexManager:
    """Gestion des index consolidés"""
    
    @staticmethod
    def merge_to_index(processed_bytes: bytes, category: str) -> Tuple[bool, str]:
        """Fusionne les nouvelles données avec l'index local existant"""
        try:
            filename = f"index_{category.replace(' ', '_')}.xlsx"
            new_data_io = io.BytesIO(processed_bytes)
            new_df_dict = pd.read_excel(new_data_io, sheet_name=None)
            
            # Combiner toutes les feuilles du nouveau traitement en un seul DF
            new_df = pd.concat(new_df_dict.values(), ignore_index=True)
            
            # Récupérer l'index existant
            existing_content = FileManager.get_local_index(filename)
            if existing_content:
                existing_df = pd.read_excel(io.BytesIO(existing_content))
                final_df = pd.concat([existing_df, new_df], ignore_index=True)
            else:
                final_df = new_df
            
            # Supprimer les doublons sur l'index global
            # On garde la dernière occurrence pour mettre à jour les données si besoin
            subset_cols = [c for c in final_df.columns if c not in TRACE_COLS]
            final_df = final_df.drop_duplicates(subset=subset_cols, keep='last')
            
            # Sauvegarder
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name=INDEX_SHEET_NAME)
                ws = writer.sheets[INDEX_SHEET_NAME]
                ExcelStyler.apply_header_style(ws, len(final_df.columns))
                ExcelStyler.apply_auto_width(ws)
            
            if FileManager.save_index_locally(output.getvalue(), filename):
                return True, filename
            return False, ""
        except Exception as e:
            st.error(f"Erreur fusion index : {e}")
            return False, ""


# ============================================================================
# COMPOSANTS UI
# ============================================================================

class UIComponents:
    """Gestion des composants d'interface Streamlit"""
    
    @staticmethod
    def setup_page_config() -> None:
        """Configure les paramètres de la page"""
        st.set_page_config(
            page_title="DataHub Pro | Excel Optimizer",
            page_icon="📊",
            layout="wide",
            initial_sidebar_state="expanded"
        )
    
    @staticmethod
    def apply_custom_styles() -> None:
        """Applique le CSS personnalisé"""
        st.markdown(f"""
        <style>
            .main {{ background-color: #f0f2f6; }}
            .stApp {{ background-color: #f0f2f6; }}
            .stButton>button {{ 
                border-radius: 12px; 
                height: 3.5em; 
                font-weight: 700; 
                transition: all 0.3s; 
            }}
            .stButton>button:hover {{ 
                transform: translateY(-2px); 
                box-shadow: 0 4px 12px rgba(0,0,0,0.1); 
            }}
            .step-card {{ 
                background-color: #ffffff; 
                padding: 25px; 
                border-radius: 15px; 
                margin-bottom: 25px; 
                box-shadow: 0 4px 6px rgba(0,0,0,0.05); 
                border-top: 5px solid {PRIMARY_COLOR}; 
            }}
            .step-header {{ 
                display: flex; 
                align-items: center; 
                margin-bottom: 15px; 
            }}
            .step-number {{ 
                background-color: {PRIMARY_COLOR}; 
                color: white; 
                border-radius: 50%; 
                width: 30px; 
                height: 30px; 
                display: flex; 
                align-items: center; 
                justify-content: center; 
                font-weight: bold; 
                margin-right: 15px; 
            }}
            .step-title {{ 
                color: {PRIMARY_COLOR}; 
                font-size: 1.3em; 
                font-weight: bold; 
            }}
            .stat-card {{ 
                background-color: #f8f9fa; 
                padding: 15px; 
                border-radius: 10px; 
                text-align: center; 
                border: 1px solid #e9ecef; 
            }}
            .stat-val {{ 
                font-size: 1.8em; 
                font-weight: bold; 
                color: {PRIMARY_COLOR}; 
            }}
            .stat-label {{ 
                font-size: 0.9em; 
                color: #6c757d; 
            }}
            /* Styles améliorés pour les tâches quotidiennes */
            .task-card {{
                background-color: white;
                border-radius: 12px;
                padding: 20px;
                margin-bottom: 15px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.05);
                display: flex;
                align-items: center;
                border-left: 6px solid #e9ecef;
                transition: all 0.2s;
            }}
            .task-card-completed {{
                border-left-color: {SUCCESS_COLOR};
                background-color: #f8fff9;
            }}
            .task-icon {{
                font-size: 2em;
                margin-right: 20px;
                min-width: 50px;
                text-align: center;
            }}
            .task-content {{
                flex-grow: 1;
            }}
            .task-title {{
                font-weight: bold;
                font-size: 1.1em;
                margin-bottom: 4px;
                color: {PRIMARY_COLOR};
            }}
            .task-desc {{
                color: #6c757d;
                font-size: 0.9em;
            }}
            .task-badge {{
                padding: 4px 10px;
                border-radius: 20px;
                font-size: 0.75em;
                font-weight: bold;
                text-transform: uppercase;
                margin-top: 8px;
                display: inline-block;
            }}
            .badge-todo {{ background-color: #ffeeba; color: #856404; }}
            .badge-done {{ background-color: #d4edda; color: #155724; }}
            
            .footer {{ 
                text-align: center; 
                color: #adb5bd; 
                padding: 40px 0; 
                font-size: 0.85em; 
            }}
        </style>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def render_sidebar() -> None:
        """Affiche la barre latérale"""
        with st.sidebar:
            st.markdown(f"<h2 style='color:{PRIMARY_COLOR};'>Centre de Contrôle</h2>", unsafe_allow_html=True)
            
            # Navigation
            # st.markdown("---")
            if st.button("📊 Traitement des fichiers", use_container_width=True):
                st.session_state.page = "Traitement"
                st.rerun()
            if st.button("📅 Tâches Quotidiennes", use_container_width=True):
                st.session_state.page = "Taches"
                st.rerun()
            if st.button("📚 Bibliothèque des Index", use_container_width=True):
                st.session_state.page = "Index"
                st.rerun()
            
            # # Information stockage
            # st.markdown("---")
            # st.markdown("### 💾 Stockage")
            
            # try:
            #     index_folder = FileManager.get_index_folder()
            #     st.success(f"Dossier prêt : `{index_folder.name}`")
            # except Exception as e:
            #     st.error(f"Erreur dossier : {e}")
            
            # st.info("Les index sont sauvegardés localement dans votre dossier utilisateur.")
            
            st.markdown("---")
            if st.button("🔄 Réinitialiser la session", use_container_width=True):
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
            st.markdown(f"""
            <div class="step-card">
                <div class="step-header">
                    <div class="step-number">1</div>
                    <div class="step-title">Configuration</div>
                </div>
            """, unsafe_allow_html=True)
            
            file_type = st.selectbox(
                "Module métier", 
                options=list(COLUMNS_CONFIG.keys()),
                index=None,
                placeholder="Sélectionnez une catégorie..."
            )
            
            if file_type:
                st.markdown(f"*{COLUMNS_CONFIG[file_type]['description']}*")
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col_main2:
            st.markdown(f"""
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
                # Validation : le traitement n'est possible que si une catégorie est sélectionnée
                button_disabled = not file_type
                if button_disabled:
                    st.warning("Veuillez sélectionner une catégorie de traitement avant de lancer l'analyse.")
                
                if st.button("Lancer l'analyse", use_container_width=True, type="primary", disabled=button_disabled):
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
        st.markdown(f"""
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
                <div class="stat-val">{{total_in}}</div>
                <div class="stat-label">Lignes Entrantes</div>
            </div>
            """, unsafe_allow_html=True)
        
        with s2:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-val" style="color:{SUCCESS_COLOR};">{{total_out}}</div>
                <div class="stat-label">Lignes Uniques</div>
            </div>
            """, unsafe_allow_html=True)
        
        with s3:
            st.markdown(f"""
            <div class="stat-card">
                <div class="stat-val" style="color:{DANGER_COLOR};">{{total_in - total_out}}</div>
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
            
            if st.session_state.current_file_type:
                if st.button(f"Fusionner vers Index {st.session_state.current_file_type}", use_container_width=True, type="primary"):
                    success, fname = self.index_manager.merge_to_index(
                        st.session_state.processed_data, 
                        st.session_state.current_file_type
                    )
                    if success:
                        st.success(f"Données fusionnées avec succès dans `{fname}`")
                        if st.button("Voir la bibliothèque", use_container_width=True):
                            st.session_state.page = "Index"
                            st.rerun()
                    else:
                        st.error("Erreur lors de la fusion de l'index.")
        
        st.markdown('</div>', unsafe_allow_html=True)

    def render_daily_tasks_page(self) -> None:
        """Affiche la page des tâches quotidiennes avec un design amélioré"""
        st.markdown("<h1>Suivi des Tâches Quotidiennes</h1>", unsafe_allow_html=True)
        
        completed, total = self.task_manager.get_progress()
        progress_percent = (completed / total) * 100 if total > 0 else 0
        today_str = datetime.now().strftime("%A %d %B %Y")
        
        # En-tête de progression
        col_head1, col_head2 = st.columns([2, 1])
        with col_head1:
            st.markdown(f"### {today_str}")
            st.markdown(f"Complétion : **{completed} sur {total} tâches**")
        with col_head2:
            st.markdown(f"<h2 style='text-align:right; color:{SUCCESS_COLOR if completed==total else PRIMARY_COLOR};'>{progress_percent:.0f}%</h2>", unsafe_allow_html=True)
        
        st.progress(progress_percent / 100)
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Liste des tâches avec nouveau design
        for task_name, task_config in sorted(DAILY_TASKS_CONFIG.items(), key=lambda x: x[1]["priority"]):
            is_completed = self.task_manager.is_task_completed(task_name)
            
            status_class = "task-card-completed" if is_completed else ""
            badge_class = "badge-done" if is_completed else "badge-todo"
            badge_text = "Complété" if is_completed else "À faire"
            
            st.markdown(f"""
            <div class="task-card {status_class}">
                <div class="task-icon">{task_config['icon']}</div>
                <div class="task-content">
                    <div class="task-title">{task_name}</div>
                    <div class="task-desc">{task_config['description']}</div>
                    <div class="task-badge {badge_class}">{badge_text}</div>
                </div>
                <div style="font-size: 1.5em;">
                    {'✅' if is_completed else '⏳'}
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Section Info & Action
        col_info1, col_info2 = st.columns([1, 1])
        with col_info1:
            if completed == total:
                st.success("Félicitations ! Toutes les tâches sont terminées.")
            else:
                st.info(f"Il vous reste {total - completed} tâches pour clôturer la journée.")
        
        with col_info2:
            if st.button("?? Aller au Traitement", use_container_width=True, type="primary"):
                st.session_state.page = "Traitement"
                st.rerun()

    def render_index_page(self) -> None:
        """Affiche la page de la bibliothèque d'index"""
        st.markdown("<h1> Bibliothèque des Index</h1>", unsafe_allow_html=True)
        
        try:
            indexes = self.file_manager.list_local_indexes()
        except Exception as e:
            st.error(f"Erreur d'accès : {e}")
            indexes = []
        
        if not indexes:
            st.warning("Aucun index trouvé. Traitez des fichiers pour commencer à construire votre base de données.")
        else:
            for idx_name in sorted(indexes):
                with st.expander(f"📄 {idx_name}", expanded=False):
                    c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
                    with c1:
                        cat = idx_name.replace('index_', '').replace('.xlsx', '').replace('_', ' ')
                        st.markdown(f"**Catégorie** : {cat}")
                    
                    with c2:
                        content = self.file_manager.get_local_index(idx_name)
                        if content:
                            st.download_button("Télécharger", content, idx_name, key=f"dl_{idx_name}", use_container_width=True)
                    
                    with c3:
                        content_dl = self.file_manager.get_local_index(idx_name)
                        if content_dl:
                            st.download_button(
                                label="📥 Exporter",
                                data=content_dl,
                                file_name=idx_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"exp_{idx_name}",
                                use_container_width=True,
                            )
                    
                    with c4:
                        if st.button("Supprimer", key=f"del_{idx_name}"):
                            if self.file_manager.delete_index(idx_name):
                                st.rerun()

    def run(self) -> None:
        """Point d'entrée principal de l'application"""
        self.ui.setup_page_config()
        self.ui.apply_custom_styles()
        self.initialize_session()
        self.ui.render_sidebar()
        
        if st.session_state.page == "Traitement":
            self.render_processing_page()
        elif st.session_state.page == "Taches":
            self.render_daily_tasks_page()
        elif st.session_state.page == "Index":
            self.render_index_page()
        
        st.markdown('<div class="footer">DataHub Pro v3.1 | Stockage Local | © 2024</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    app = DataHubApp()
    app.run()
