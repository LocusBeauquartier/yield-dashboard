"""
download_sharepoint.py — Télécharge les xlsx depuis SharePoint via Microsoft Graph API
=======================================================================================
Appelé automatiquement par GitHub Actions avant generate.py.

Variables d'environnement requises (GitHub Secrets) :
    SHAREPOINT_TENANT_ID     → ID du tenant Azure AD (ex: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
    SHAREPOINT_CLIENT_ID     → App ID de votre App Registration Azure
    SHAREPOINT_CLIENT_SECRET → Secret de votre App Registration Azure
    SHAREPOINT_SITE_URL      → URL du site SharePoint (ex: "https://locusinvest.sharepoint.com/sites/Beauquartier")
    SHAREPOINT_FOLDER_PATH   → Chemin du dossier dans SharePoint (ex: "Documents/YieldDashboard/exports")

Comment créer l'App Registration Azure (une seule fois) :
    1. Azure Portal → Azure Active Directory → App Registrations → New registration
    2. Nom : "YieldDashboard", type compte : "Single tenant"
    3. API Permissions → Add → Microsoft Graph → Application → Sites.Read.All, Files.Read.All
    4. Grant admin consent
    5. Certificates & Secrets → New client secret → copier la valeur
    6. Copier l'Application (client) ID et le Directory (tenant) ID
"""

import os
import sys
import json
from pathlib import Path

try:
    import requests
    import msal
except ImportError:
    print("❌ Modules manquants — pip install requests msal")
    sys.exit(1)

# ── CONFIG ────────────────────────────────────────────────────────────────────

# Mapping nom local → nom du fichier sur SharePoint
# Adaptez les noms des fichiers SharePoint à votre convention
SHAREPOINT_FILE_MAP = {
    "export_J.xlsx":        "Export_J.xlsx",
    "export_J1.xlsx":       "Export_J1.xlsx",
    "fenetre_J1.xlsx":      "Fenetre_J1.xlsx",
    "fenetre_J3.xlsx":      "Fenetre_J3.xlsx",
    "fenetre_J7.xlsx":      "Fenetre_J7.xlsx",
    "fenetre_J14.xlsx":     "Fenetre_J14.xlsx",
    "fenetre_J21.xlsx":     "Fenetre_J21.xlsx",
    "fenetre_J45.xlsx":     "Fenetre_J45.xlsx",
    "budget.xlsx":          "Budget.xlsx",
    "reservations.xlsx":    "Reservations.xlsx",
}
# Pour les fichiers pickup, on télécharge tous les fichiers qui commencent par "pickup_" ou "export_"
PICKUP_PREFIX = "pickup_"

OUTPUT_DIR = Path("./data")

# ── AUTH MICROSOFT GRAPH ──────────────────────────────────────────────────────

def get_access_token():
    tenant_id     = os.environ["SHAREPOINT_TENANT_ID"]
    client_id     = os.environ["SHAREPOINT_CLIENT_ID"]
    client_secret = os.environ["SHAREPOINT_CLIENT_SECRET"]

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise ValueError(f"Authentification échouée : {result.get('error_description', result)}")
    return result["access_token"]


def get_site_id(token, site_url):
    """Récupère le site ID SharePoint depuis l'URL."""
    # Extraire hostname et path depuis l'URL
    # ex: https://locusinvest.sharepoint.com/sites/Beauquartier
    from urllib.parse import urlparse
    parsed = urlparse(site_url)
    hostname = parsed.netloc          # locusinvest.sharepoint.com
    site_path = parsed.path.lstrip("/")  # sites/Beauquartier

    url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    return resp.json()["id"]


def list_folder_files(token, site_id, folder_path):
    """Liste tous les fichiers d'un dossier SharePoint."""
    encoded_path = folder_path.replace(" ", "%20")
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{encoded_path}:/children"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    return resp.json().get("value", [])


def download_file(token, download_url, local_path):
    """Télécharge un fichier depuis SharePoint."""
    resp = requests.get(download_url, headers={"Authorization": f"Bearer {token}"}, stream=True)
    resp.raise_for_status()
    with open(local_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    # Créer le dossier de sortie
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    site_url    = os.environ.get("SHAREPOINT_SITE_URL", "")
    folder_path = os.environ.get("SHAREPOINT_FOLDER_PATH", "Documents/YieldDashboard")

    if not site_url:
        print("❌ SHAREPOINT_SITE_URL non défini")
        sys.exit(1)

    print(f"🔐 Authentification Microsoft Graph...")
    token = get_access_token()
    print("✓ Token obtenu")

    print(f"📁 Connexion au site : {site_url}")
    site_id = get_site_id(token, site_url)
    print(f"✓ Site ID : {site_id}")

    print(f"📂 Liste des fichiers dans : {folder_path}")
    files = list_folder_files(token, site_id, folder_path)
    print(f"✓ {len(files)} fichiers trouvés")

    # Index SharePoint par nom
    sp_index = {f["name"]: f for f in files if "file" in f}

    downloaded = 0
    errors = []

    # Télécharger les fichiers principaux
    for local_name, sp_name in SHAREPOINT_FILE_MAP.items():
        if sp_name in sp_index:
            local_path = OUTPUT_DIR / local_name
            try:
                download_url = sp_index[sp_name]["@microsoft.graph.downloadUrl"]
                download_file(token, download_url, local_path)
                size = local_path.stat().st_size // 1024
                print(f"  ✓ {sp_name:<35} → {local_name} ({size} Ko)")
                downloaded += 1
            except Exception as e:
                errors.append(f"{sp_name} : {e}")
                print(f"  ✗ {sp_name} — ERREUR : {e}")
        else:
            print(f"  ⚠ {sp_name:<35} — absent sur SharePoint")

    # Télécharger tous les fichiers pickup (prefixe "pickup_")
    pickup_files = [f for f in sp_index.values() if f["name"].lower().startswith(PICKUP_PREFIX)]
    for pf in pickup_files:
        local_path = OUTPUT_DIR / pf["name"]
        try:
            download_url = pf["@microsoft.graph.downloadUrl"]
            download_file(token, download_url, local_path)
            size = local_path.stat().st_size // 1024
            print(f"  ✓ {pf['name']:<35} → pickup ({size} Ko)")
            downloaded += 1
        except Exception as e:
            errors.append(f"{pf['name']} : {e}")

    print(f"\n{'='*50}")
    print(f"✅ {downloaded} fichier(s) téléchargé(s)")
    if errors:
        print(f"⚠  {len(errors)} erreur(s) :")
        for e in errors:
            print(f"   - {e}")

    # Écrire un résumé JSON pour le workflow
    summary = {
        "downloaded": downloaded,
        "errors": errors,
        "files": [f.name for f in OUTPUT_DIR.glob("*.xlsx")]
    }
    with open(OUTPUT_DIR / "_summary.json", "w") as f:
        json.dump(summary, f, indent=2)


if __name__ == "__main__":
    main()
