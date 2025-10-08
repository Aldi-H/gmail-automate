import os
import pandas as pd
from dotenv import load_dotenv
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# Load .env
load_dotenv()
SERVICE_ACCOUNT_FILE = os.getenv("CREDENTIAL_FILES")
FOLDER_ID = os.getenv("FOLDER_ID")

if not SERVICE_ACCOUNT_FILE or not os.path.exists(SERVICE_ACCOUNT_FILE):
    raise FileNotFoundError(f"Service account JSON not found: {SERVICE_ACCOUNT_FILE}")

# Create a temporary settings.yaml for PyDrive2
with open("settings.yaml", "w") as f:
    f.write(f"""client_config_backend: service
service_config:
  client_json_file_path: "{SERVICE_ACCOUNT_FILE}"
""")

# Authenticate
gauth = GoogleAuth(settings_file="settings.yaml")
gauth.ServiceAuth()
drive = GoogleDrive(gauth)

# Step 1: List Excel files in target folder
file_list = drive.ListFile({'q': f"'{FOLDER_ID}' in parents and trashed=false"}).GetList()

print("✅ Successfully authenticated via Service Account.")
print(f"Found {len(file_list)} files in folder {FOLDER_ID}:")
for f in file_list:
    print(f" - {f['title']} ({f['id']})")
