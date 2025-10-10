#!/bin/bash

ENV_FILE=".env"

# Hapus HOST_DOWNLOADS_PATH lama jika ada
sed -i '/^HOST_DOWNLOADS_PATH=/d' $ENV_FILE

# Deteksi path Downloads
if [ -d "/mnt/c/Users" ]; then
    # WSL - ambil user pertama yang bukan system
    WIN_USER=$(ls /mnt/c/Users | grep -v "Public\|Default\|All Users\|desktop.ini" | head -n 1)
    DOWNLOADS_PATH="/mnt/c/Users/$WIN_USER/Downloads"
    echo "Detected Windows user: $WIN_USER"
else
    # Fallback
    DOWNLOADS_PATH="./downloads"
    echo "Using local downloads folder"
fi

# Tambahkan ke .env
echo "" >> $ENV_FILE
echo "# Downloads path (auto-generated)" >> $ENV_FILE
echo "HOST_DOWNLOADS_PATH=$DOWNLOADS_PATH" >> $ENV_FILE

echo "✓ Added HOST_DOWNLOADS_PATH=$DOWNLOADS_PATH to .env"