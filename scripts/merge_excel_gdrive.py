# type: ignore
import pandas as pd
import os
import json
from dotenv import load_dotenv
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import sys
import re

load_dotenv()
SERVICE_ACCOUNT_FILE =  os.getenv("CREDENTIAL_FILES")
FOLDER_ID = os.getenv("FOLDER_ID")
LOCAL_SAVE_PATH = os.getenv("LOCAL_SAVE_PATH")

SAVE_TO_DOWNLOADS = True

def authenticate_gdrive():
  SCOPES = ["https://www.googleapis.com/auth/drive"]
  creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
  )
  service = build("drive", "v3", credentials=creds)

  return service


def list_excel_files(service, folder_id):
  query = f"'{folder_id}' in parents and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel') and trashed=false"

  results = (
    service.files()
      .list(
        q=query,
        fields="files(id, name, createdTime)",
        orderBy="createdTime desc",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
      )
      .execute()
  )

  files = results.get("files", [])

  return files


def download_file(service, file_id, file_name, temp_dir="/tmp"):
  request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
  file_path = os.path.join(temp_dir, file_name)

  fh = io.FileIO(file_path, "wb")
  downloader = MediaIoBaseDownload(fh, request)

  done = False
  while done is False:
    done = downloader.next_chunk()
    # print(f"Download {int(status.progress() * 100)}%")

  fh.close()

  return file_path


def merge_excel_files(file_paths):
  all_data = []

  for file_path in file_paths:
    try:
      df = pd.read_excel(
        file_path,
        sheet_name=0,
        skiprows=4,
        dtype={
          "KODE_AKUN_BELANJA": str,
          "KODE_AKUN_POTONGAN_PAJAK": str,
          "NPWP_BENDAHARA": str,
          "ID_BILLING": str
        }
      )
      all_data.append(df)
      # print(f"✓ Berhasil membaca: {os.path.basename(file_path)}")
    except Exception as e:
      # print(f"✗ Error membaca {file_path}: {str(e)}")
      raise e
    
  if not all_data:
    raise Exception("✗ No data to merge")

  merged_df = pd.concat(all_data, ignore_index=True)
  merged_df = format_dataframe(merged_df)
  # print(f"\nTotal rows after merging: {len(merged_df)}")

  return merged_df

def format_dataframe(df):
  currency_columns = ["NILAI_BELANJA_SP2D", "JUMLAH_PAJAK"]
  for col in currency_columns:
    if col in df.columns:
      df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('int64')
      
  if "KODE_AKUN_BELANJA" in df.columns:
    df["KODE_AKUN_BELANJA"] = df["KODE_AKUN_BELANJA"].apply(
      lambda x: str(x).replace(".", "").strip() if pd.notna(x) else ""
    )

  if "NPWP_BENDAHARA" in df.columns:
    df["NPWP_BENDAHARA"] = df["NPWP_BENDAHARA"].apply(
      lambda x: re.sub(r"[^\d]", "", str(x)) if pd.notna(x) else ""
    )

  if "ID_BILLING" in df.columns:
    df["ID_BILLING"] = df["ID_BILLING"].apply(
      lambda x: str(x).replace(".0", "") if pd.notna(x) else ""
    )

  if "KODE_AKUN_POTONGAN_PAJAK" in df.columns:
    df["KODE_AKUN_POTONGAN_PAJAK"] = df["KODE_AKUN_POTONGAN_PAJAK"].apply(
      lambda x: str(x).replace("-100", "").strip() if pd.notna(x) else ""
    )

  return df

def save_to_local(df, directory, filename):
  if not os.path.exists(directory):
    try:
      os.makedirs(directory, exist_ok=True)
    except Exception as e:
      # print(f"✗ Gagal membuat direktori {directory}: {str(e)}", file=sys.stderr)
      raise e
    
  if not os.access(directory, os.W_OK):
    raise PermissionError(f"✗ Didn't have write permission for {directory}")
  
  output_path = os.path.join(directory, filename)
  df.to_excel(output_path, index=False, engine="openpyxl")
  
  file_size = os.path.getsize(output_path)
  file_size_mb = file_size / (1024 * 1024)
  
  return {
    "path": output_path,
    "size_mb": file_size_mb,
    "total_rows": len(df),
    "total_columns": len(df.columns)
  }


def main():
  try:
    # print("=" * 60)
    # print("SCRIPT PENGGABUNGAN FILE EXCEL - SAVE TO LOCAL")
    # print("=" * 60)

    # 1. Autentikasi Google Drive
    # print("\n[1] Melakukan autentikasi ke Google Drive...")
    service = authenticate_gdrive()
    # print("✓ Autentikasi berhasil")

    # 2. Ambil daftar file Excel dari folder
    # print(f"\n[2] Mengambil daftar file Excel dari folder...")
    excel_files = list_excel_files(service, FOLDER_ID)

    if not excel_files:
      # print("✗ Tidak ada file Excel yang ditemukan di folder")
      sys.exit(1)

    # print(f"✓ Found {len(excel_files)} Excel files:")
    # for f in excel_files:
    #   print(f"  - {f['name']}")

    # 3. Download semua file
    # print(f"\n[3] Mendownload file dari Google Drive...")
    downloaded_files = []
    for file in excel_files:
      file_path = download_file(service, file["id"], file["name"])
      downloaded_files.append(file_path)

    # 4. Gabungkan file Excel
    # print(f"\n[4] Menggabungkan file Excel...")
    merged_df = merge_excel_files(downloaded_files)

    # 5. Simpan hasil penggabungan ke LOCAL
    now = datetime.now()
    month = now.strftime("%B")
    year = now.strftime("%Y")
    output_filename = f"DTH_{month}_{year}.xlsx"

    saved_locations = []
    
    if SAVE_TO_DOWNLOADS:
      if os.path.exists(LOCAL_SAVE_PATH):
        # print(f"\n[6] Saving to Downloads folder...", file=sys.stderr)

        try:
          downloads_info = save_to_local(merged_df, LOCAL_SAVE_PATH, output_filename)
          saved_locations.append({
            "location": "downloads",
            "path": downloads_info["path"],
            "size_mb": downloads_info["size_mb"]
          })
          # print(f"✓ Saved to: {downloads_info['path']}", file=sys.stderr)
        except Exception as e:
          # print(f"⚠ Error saving to Downloads: {str(e)}", file=sys.stderr)
          raise e
      else:
        # print(f"⚠ Downloads path not mounted: {LOCAL_SAVE_PATH}", file=sys.stderr)
        # print(f"  To enable, add volume mount in docker-compose.yml:", file=sys.stderr)
        # print(f"  - /mnt/c/Users/YOUR_USERNAME/Downloads:/downloads", file=sys.stderr)
        raise Exception(f"Downloads path not mounted: {LOCAL_SAVE_PATH}", file=sys.stderr)

    # 6. Cleanup temporary files
    # print(f"\n[6] Membersihkan file temporary...")
    for file_path in downloaded_files:
      if os.path.exists(file_path):
        os.remove(file_path)
    # print("✓ Cleanup selesai")

    # print("\n" + "=" * 60)
    # print("PROSES SELESAI!")
    # print("=" * 60)

    result = {
      "success": True,
      "filename": output_filename,
      "total_rows": len(merged_df),
      "total_columns": len(merged_df.columns),
      "files_merged": len(excel_files),
      "saved_locations": saved_locations,
      "created_at": datetime.now().isoformat()
    }

    print(json.dumps(result, indent=2))

    return 0

  except Exception as e:
    result = {
      "success": False,
      "error": str(e)
    }

    print(json.dumps(result, indent=2), file=sys.stderr)

    sys.exit(1)

if __name__ == "__main__":
  result = main()
  # print(f"\nOutput file: {result}")