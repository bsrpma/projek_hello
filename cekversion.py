# script.py

import requests

versi_lokal = "1.0.0"
url = "https://raw.githubusercontent.com/bsrpma/projek_hello/main/version.txt"

def cek_versi(versi_lokal):
    try:
        r = requests.get(url, timeout=5)
        if r.status_code == 200:
            versi_online = r.text.strip()
            if versi_online != versi_lokal:
                print(f"⚠️ Versi baru tersedia: {versi_online}")
                print("Silakan download versi terbaru dari GitHub.")
            else:
                print("✅ Aplikasi sudah versi terbaru.")
        else:
            print("Gagal mengecek versi online.")
    except Exception as e:
        print("Error saat cek versi:", e)

# Jalankan fungsi cek versi
cek_versi(versi_lokal)

# Program utama
print("Hello, world!")
