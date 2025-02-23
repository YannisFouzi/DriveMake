import os
import io
import pandas as pd
from flask import Flask, request, jsonify
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2 import service_account
from openpyxl import Workbook
from functools import wraps
import time

app = Flask(__name__)

# === CONFIGURATION ===
# Utilisation des variables d'environnement pour plus de sécurité
CREDENTIALS_FILE = os.environ.get("CREDENTIALS_FILE", "client_secret.json")
FOLDER_SMC = os.environ.get("FOLDER_SMC", "1Qg_KdjEJirl0grOeDJt3dK9w3eq6fj9d")
FOLDER_TEMP = os.environ.get("FOLDER_TEMP", "1LWfFq9sD59raMuXddIgsFKm5cjdP3Gcy")

# === AUTHENTIFICATION À GOOGLE DRIVE ===
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

# Modification pour mieux gérer la clé privée
def clean_private_key(key):
    """Nettoie et formate correctement la clé privée"""
    key = key.replace('\\n', '\n').replace('\\\\n', '\n')
    key = key.strip().strip('"\'')
    if not key.startswith('-----BEGIN PRIVATE KEY-----'):
        key = '-----BEGIN PRIVATE KEY-----\n' + key
    if not key.endswith('-----END PRIVATE KEY-----'):
        key = key + '\n-----END PRIVATE KEY-----'
    return key

credentials_info = {
    "type": os.environ.get("GOOGLE_TYPE"),
    "project_id": os.environ.get("GOOGLE_PROJECT_ID"),
    "private_key_id": os.environ.get("GOOGLE_PRIVATE_KEY_ID"),
    "private_key": clean_private_key(os.environ.get("GOOGLE_PRIVATE_KEY", "")),
    "client_email": os.environ.get("GOOGLE_CLIENT_EMAIL"),
    "client_id": os.environ.get("GOOGLE_CLIENT_ID"),
    "auth_uri": os.environ.get("GOOGLE_AUTH_URI"),
    "token_uri": os.environ.get("GOOGLE_TOKEN_URI"),
    "auth_provider_x509_cert_url": os.environ.get("GOOGLE_AUTH_PROVIDER_X509_CERT_URL"),
    "client_x509_cert_url": os.environ.get("GOOGLE_CLIENT_X509_CERT_URL"),
    "universe_domain": os.environ.get("GOOGLE_UNIVERSE_DOMAIN")
}

# Ajout de logs pour debug
print("=== DEBUG INFO ===")
print("Private key format check:")
print("Starts with correct header:", credentials_info["private_key"].startswith('-----BEGIN PRIVATE KEY-----'))
print("Ends with correct footer:", credentials_info["private_key"].endswith('-----END PRIVATE KEY-----'))
print("Contains newlines:", '\n' in credentials_info["private_key"])
print("Key length:", len(credentials_info["private_key"]))
print("First 50 chars:", credentials_info["private_key"][:50])
print("Last 50 chars:", credentials_info["private_key"][-50:])
print("================")

try:
    creds = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
    drive_service = build("drive", "v3", credentials=creds)
except Exception as e:
    print("=== ERROR DETAILS ===")
    print(f"Error type: {type(e)}")
    print(f"Error message: {str(e)}")
    print("===================")
    raise

# === FONCTIONS UTILITAIRES ===

def list_files_in_folder(folder_id):
    """ Liste les fichiers dans un dossier Google Drive """
    query = f"'{folder_id}' in parents and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id, name, mimeType)").execute()
    files = results.get("files", [])
    return files

def get_latest_file(folder_id):
    """ Récupère le fichier le plus récent dans un dossier Google Drive """
    query = f"'{folder_id}' in parents and trashed=false"
    results = drive_service.files().list(q=query, orderBy="createdTime desc", fields="files(id, name, mimeType)").execute()
    files = results.get("files", [])
    return files[0] if files else None

def download_file(file_id, file_name, mime_type):
    """ Télécharge un fichier depuis Google Drive, en exportant s'il s'agit d'un Google Sheets """
    # Utiliser le répertoire /tmp pour les fichiers temporaires
    file_path = os.path.join('/tmp', file_name)
    
    if mime_type == "application/vnd.google-apps.spreadsheet":
        request = drive_service.files().export_media(fileId=file_id, mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        request = drive_service.files().get_media(fileId=file_id)
    
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    with open(file_path, "wb") as f:
        f.write(fh.read())
    print(f"Téléchargé : {file_path}")
    return file_path  # Retourner le chemin complet

def detect_email_column(df):
    """ Détecte la colonne contenant les emails """
    for col in df.columns:
        if df[col].astype(str).str.contains(r"@").sum() > 0:
            return col
    return None

def merge_data(old_file, new_file):
    """ Fusionne les données : Ajout de nouvelles lignes, mise à jour des existantes """
    df_old = pd.read_excel(old_file, engine='openpyxl', dtype=str)
    df_new = pd.read_excel(new_file, engine='openpyxl', dtype=str)

    email_col_old = detect_email_column(df_old)
    email_col_new = detect_email_column(df_new)

    if email_col_old != email_col_new:
        raise ValueError("La colonne email ne correspond pas entre les fichiers")

    # Pour tracker les modifications
    modified_cells = []
    new_rows = []
    
    # Vérifier si un email existe et mettre à jour les données
    for index, new_row in df_new.iterrows():
        match_index = df_old[df_old[email_col_old] == new_row[email_col_new]].index
        if not match_index.empty:
            # Pour les lignes existantes
            for col_idx, col in enumerate(df_new.columns):
                if pd.notna(new_row[col]) and df_old.iloc[match_index[0]][col] != new_row[col]:
                    df_old.loc[match_index, col] = new_row[col]
                    modified_cells.append((match_index[0], col_idx))
        else:
            # Pour les nouvelles lignes
            new_row_idx = len(df_old)
            df_old = pd.concat([df_old, pd.DataFrame([new_row])], ignore_index=True)
            new_rows.append(new_row_idx)
    
    # Sauvegarde comme avant
    data = [df_old.columns.tolist()] + df_old.values.tolist()
    wb = Workbook()
    ws = wb.active
    for row in data:
        ws.append(row)
    wb.save(old_file)
    
    return old_file, modified_cells, new_rows

def update_existing_file(file_path, file_id, modified_cells=None, new_rows=None):
    """Remplace le fichier existant sur Google Drive et applique le formatage"""
    # Mise à jour du contenu comme avant
    media = MediaFileUpload(file_path, 
                          mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                          resumable=True)
    
    file_metadata = {
        "name": os.path.basename(file_path),
        "mimeType": "application/vnd.google-apps.spreadsheet"  # Convertir en Google Sheets
    }
    
    # Mise à jour et conversion en Google Sheets
    updated_file = drive_service.files().update(
        fileId=file_id,
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    # Attendre un peu que la conversion soit terminée
    time.sleep(2)

    # Création du service Sheets
    sheets_service = build('sheets', 'v4', credentials=creds)
    
    # Préparation des requêtes de formatage
    requests = []
    
    # Pour les cellules modifiées (bleu)
    if modified_cells:
        print(f"Coloration des cellules modifiées : {modified_cells}")
        for row, col in modified_cells:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": 0,  # Première feuille
                        "startRowIndex": row,
                        "endRowIndex": row + 1,
                        "startColumnIndex": col,
                        "endColumnIndex": col + 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1}
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
    
    # Pour les nouvelles lignes (rouge)
    if new_rows:
        print(f"Coloration des nouvelles lignes : {new_rows}")
        for row in new_rows:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": 0,  # Première feuille
                        "startRowIndex": row,
                        "endRowIndex": row + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 999  # Toute la ligne
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 1, "green": 0.8, "blue": 0.8}
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
    
    # Application du formatage
    if requests:
        try:
            print("Application du formatage...")
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=file_id,
                body={"requests": requests}
            ).execute()
            print("Formatage appliqué avec succès")
        except Exception as e:
            print(f"Erreur lors du formatage : {str(e)}")
    
    print(f"Fichier mis à jour avec formatage : {os.path.basename(file_path)} (ID: {file_id})")

# Ajout après les configurations
API_KEY = os.environ.get("API_KEY")

def require_api_key(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        auth_header = request.headers.get('X-API-Key')
        print(f"Received API Key: {auth_header}")
        print(f"Expected API Key: {API_KEY}")
        if not auth_header or auth_header != API_KEY:
            return jsonify({
                "success": False,
                "message": "Accès non autorisé"
            }), 401
        return f(*args, **kwargs)
    return decorated_function

# Modification de la route
@app.route('/trigger-update', methods=['POST'])
@require_api_key
def trigger_update():
    try:
        print("🔍 Démarrage du processus de mise à jour...")
        
        # Vérification des fichiers
        old_file_info = get_latest_file(FOLDER_SMC)
        new_file_info = get_latest_file(FOLDER_TEMP)

        if not old_file_info or not new_file_info:
            return jsonify({
                "success": False,
                "message": "Fichiers non trouvés dans les dossiers spécifiés"
            }), 404

        # Utiliser des chemins complets avec /tmp
        old_file_path = os.path.join('/tmp', old_file_info["name"])
        new_file_path = os.path.join('/tmp', new_file_info["name"])

        # Téléchargement des fichiers
        old_file_path = download_file(old_file_info["id"], old_file_info["name"], old_file_info["mimeType"])
        new_file_path = download_file(new_file_info["id"], new_file_info["name"], new_file_info["mimeType"])

        # Fusion des données
        print("🔄 Fusion des fichiers...")
        updated_file, modified_cells, new_rows = merge_data(old_file_path, new_file_path)

        # Mise à jour sur Google Drive
        print("📤 Mise à jour du fichier sur Google Drive...")
        update_existing_file(updated_file, old_file_info["id"], modified_cells, new_rows)

        # Nettoyage des fichiers temporaires
        try:
            os.remove(old_file_path)
            os.remove(new_file_path)
        except Exception as e:
            print(f"Erreur lors du nettoyage des fichiers : {e}")

        return jsonify({
            "success": True,
            "message": "Mise à jour terminée avec succès"
        })

    except Exception as e:
        print(f"Erreur détaillée : {str(e)}")
        return jsonify({
            "success": False,
            "message": f"Erreur lors de la mise à jour: {str(e)}"
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
