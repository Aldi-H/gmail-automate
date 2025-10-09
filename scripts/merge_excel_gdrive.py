import pandas as pd
import os
from dotenv import load_dotenv # pyright: ignore[reportMissingImports]
from datetime import datetime
from google.oauth2.credentials import Credentials # pyright: ignore[reportMissingImports]
from google.oauth2 import service_account # pyright: ignore[reportMissingImports]
from googleapiclient.discovery import build # pyright: ignore[reportMissingImports]
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload # pyright: ignore[reportMissingImports]
import io
import sys

load_dotenv()
SERVICE_ACCOUNT_FILE =  os.getenv("CREDENTIAL_FILES")
FOLDER_ID = os.getenv("FOLDER_ID")

def authenticate_gdrive():
  SCOPES = ['https://www.googleapis.com/auth/drive']
  creds = service_account.Credentials.from_service_account_file(
  SERVICE_ACCOUNT_FILE, scopes=SCOPES)

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

def download_file(service, file_id, file_name):
  request = service.files().get_media(fileId=file_id)
  file_path = f"/tmp/{file_name}"

  fh = io.FileIO(file_path, 'wb')
  downloader = MediaIoBaseDownload(fh, request)

  done = False
  while done is False:
    status, done = downloader.next_chunk()
    print(f"Download {int(status.progress() * 100)}%")

  fh.close()
  return file_path

def merge_excel_files(file_paths):
  all_data = []

  for file_path in file_paths:
    try:
      df = pd.read_excel(file_path, sheet_name=0)
      all_data.append(df)
      print(f"✓ Berhasil membaca: {os.path.basename(file_path)}")
    except Exception as e:
      print(f"✗ Error membaca {file_path}: {str(e)}")

  if all_data:
    merged_df = pd.concat(all_data, ignore_index=True)
    print(f"\nTotal baris setelah penggabungan: {len(merged_df)}")

    return merged_df
  else:
    raise Exception("Tidak ada data yang berhasil dibaca")

def upload_to_gdrive(service, file_path, folder_id, file_name):
  file_metadata = {
    'name': file_name,
    'parents': [folder_id]
  }

  media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

  file = service.files().create(
    body=file_metadata,
    media_body=media,
    fields='id, name, webViewLink',
    supportsAllDrives=True
  ).execute()

  print(f"\n✓ File berhasil diupload: {file.get('name')}")
  print(f"Link: {file.get('webViewLink')}")

  return file

def main():
  try:
    print("=" * 60)
    print("SCRIPT PENGGABUNGAN FILE EXCEL")
    print("=" * 60)

    # 1. Autentikasi Google Drive
    print("\n[1] Melakukan autentikasi ke Google Drive...")
    service = authenticate_gdrive()
    print("✓ Autentikasi berhasil")

    # 2. Ambil daftar file Excel dari folder
    print(f"\n[2] Mengambil daftar file Excel dari folder...")
    excel_files = list_excel_files(service, FOLDER_ID)

    if not excel_files:
      print("✗ Tidak ada file Excel yang ditemukan di folder")
      sys.exit(1)

    print(f"✓ Ditemukan {len(excel_files)} file Excel:")
    for f in excel_files:
      print(f"  - {f['name']}")

    # 3. Download semua file
    print(f"\n[3] Mendownload file dari Google Drive...")
    downloaded_files = []
    for file in excel_files:
      file_path = download_file(service, file['id'], file['name'])
      downloaded_files.append(file_path)

    # 4. Gabungkan file Excel
    print(f"\n[4] Menggabungkan file Excel...")
    merged_df = merge_excel_files(downloaded_files)

    # 5. Simpan hasil penggabungan
    now = datetime.now()
    month = now.strftime("%B")
    year = now.strftime("%Y")
    output_filename = f"DTH_{month}_{year}.xlsx"
    output_path = f"/tmp/{output_filename}"

    print(f"\n[5] Menyimpan hasil penggabungan...")
    merged_df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"✓ File tersimpan: {output_path}")

    # 6. Upload hasil ke Google Drive
    print(f"\n[6] Mengupload hasil ke Google Drive...")
    uploaded_file = upload_to_gdrive(service, output_path, FOLDER_ID, output_filename)

    # 7. Cleanup temporary files
    print(f"\n[7] Membersihkan file temporary...")
    for file_path in downloaded_files:
      os.remove(file_path)
    os.remove(output_path)
    print("✓ Cleanup selesai")

    print("\n" + "=" * 60)
    print("PROSES SELESAI!")
    print("=" * 60)
    print(f"File ID: {uploaded_file.get('id')}")
        
  except Exception as e:
    print(f"\n✗ ERROR: {str(e)}")
    sys.exit(1)

if __name__ == "__main__":
  main()