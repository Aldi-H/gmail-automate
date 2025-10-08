import os
import io
from datetime import datetime
import pandas as pd
from dotenv import load_dotenv
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# Load .env
load_dotenv()
SERVICE_ACCOUNT_FILE = os.getenv("CREDENTIAL_FILES")
FOLDER_ID = os.getenv("FOLDER_ID")

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
print("✅ Successfully authenticated via Service Account.\n")

# List Excel files in folder
file_list = drive.ListFile({'q': f"'{FOLDER_ID}' in parents and trashed=false"}).GetList()

excel_files = [f for f in file_list if f['title'].endswith(('.xlsx', '.xls'))]

print(f"🔍 Found {len(excel_files)} Excel files in folder {FOLDER_ID}:")
for f in excel_files:
    print(f" - {f['title']} ({f['id']})")

if not excel_files:
    raise ValueError("No Excel files found.")

# Merge all Excel Files
merge_df = pd.DataFrame()
sheet_name = "KPP"

for file in excel_files:
    print(f"📖 Reading: {file['title']}")
    try:
        file_stream = file.GetContentIOBuffer()  # ✅ correct method
        df = pd.read_excel(file_stream, sheet_name=sheet_name, skiprows=4)
        merge_df = pd.concat([merge_df, df], ignore_index=True)
    except Exception as e:
        print(f"⚠️ Skipping {file['title']} due to error: {e}")

# Generate new filename
now = datetime.now()
output_filename = f"DTH_{now.strftime('%B')}_{now.year}.xlsx"

# Save merged file locally
merge_df.to_excel(output_filename, index=False)
print(f"\n✅ Successfully merged {len(excel_files)} files.")
# print(f"📁 Saved merged file as: {output_filename}")

output_file = drive.CreateFile({'title': output_filename, 'parents': [{'id': FOLDER_ID}]})
output_file.SetContentFile(output_filename)
output_file.Upload()
print(f"☁️ Uploaded merged file to Drive folder: {FOLDER_ID}")

