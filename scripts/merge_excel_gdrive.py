import pandas as pd
import os
from datetime import datetime
from dotenv import load_dotenv # pyright: ignore[reportMissingImports]
from google.oauth2.credentials import Credentials # pyright: ignore[reportMissingImports]
from google.oauth2 import service_account # pyright: ignore[reportMissingImports]
from googleapiclient.discovery import build # pyright: ignore[reportMissingImports]
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload # pyright: ignore[reportMissingImports]
import io
import sys

# Load .env
load_dotenv()
SERVICE_ACCOUNT_FILE = os.getenv("CREDENTIAL_FILES")
FOLDER_ID = os.getenv("FOLDER_ID_NEW")
# print(f"Using folder ID: {FOLDER_ID}")
print("FOLDER_ID_NEW =", os.getenv("FOLDER_ID_NEW"))


def authenticate_gdrive():
  SCOPES = ['https://www.googleapis.com/auth/drive']
  creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
  )
  
  service = build('drive', 'v3', credentials=creds)

  return service

def list_excel_files(service, folder_id):
  query = f"'{folder_id}' in parents and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel') and trashed=false"
  
  results = service.files().list(
    q=query,
    fields="files(id, name, createdTime)",
    orderBy="createdTime desc"
  ).execute()
  
  files = results.get('files', [])

  return files

def main():
  try:
    service = authenticate_gdrive()

    excel_files = list_excel_files(service, FOLDER_ID)

    if not excel_files:
      sys.exit("No Excel files found.")
    
    print(f"✓ Ditemukan {len(excel_files)} file Excel:")
    for f in excel_files:
      print(f"  - {f['name']}")

  except Exception as e:
    # raise e
    print(f"\n✗ ERROR: {str(e)}")
    sys.exit(1)
    
if __name__ == "__main__":
  main()