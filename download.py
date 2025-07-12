import requests

# URL file Python yang mau didownload
url_file = "https://raw.githubusercontent.com/bsrpma/projek_hello/main/script.py"

# URL file version.txt
url_version = "https://raw.githubusercontent.com/bsrpma/projek_hello/main/version.txt"

try:
    # Ambil versi dari GitHub
    r_ver = requests.get(url_version, timeout=5)
    r_ver.raise_for_status()
    versi = r_ver.text.strip()

    # Nama file berdasarkan versi
    nama_file = f"script_{versi}.py"

    # Download file Python
    r_file = requests.get(url_file, timeout=10)
    r_file.raise_for_status()

    with open(nama_file, "w", encoding="utf-8") as f:
        f.write(r_file.text)

    print(f"✅ File berhasil di-download sebagai: {nama_file}")

except Exception as e:
    print(f"❌ Gagal download file: {e}")
