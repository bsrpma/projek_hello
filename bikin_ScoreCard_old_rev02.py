import os
import shutil
import win32com.client
import requests
import sys

class GitHelper:
    url_version = "https://raw.githubusercontent.com/bsrpma/projek_hello/main/version.txt"
    url_exe = "https://raw.githubusercontent.com/bsrpma/projek_hello/main/dist/bikin_ScoreCard.exe"
    url_txt = "https://raw.githubusercontent.com/bsrpma/projek_hello/main/depots.txt"

    def __init__(self, versi_lokal="1.0.1", download_dir="download"):
        self.versi_lokal = versi_lokal
        self.download_dir = download_dir
        os.makedirs(self.download_dir, exist_ok=True)

    def cek_versi(self):
        try:
            r = requests.get(self.url_version, timeout=5)
            r.raise_for_status()
            versi_online = r.text.strip()

            if versi_online != self.versi_lokal:
                print(f"⚠️ Versi baru tersedia: {versi_online} (lokal: {self.versi_lokal})")
                print("  [1] Download versi baru")
                print("  [2] Gunakan versi lokal")
                print("  [3] Keluar")
                pilihan = input("Masukkan pilihan (1/2/3): ").strip()

                if pilihan == "1":
                    self.download_file()
                    input("✅ Selesai. Tekan Enter untuk keluar...")
                    sys.exit()
                elif pilihan == "2":
                    print("Lanjut dengan versi lokal...\n")
                elif pilihan == "3":
                    sys.exit("Keluar dari program.")
                else:
                    sys.exit("Input tidak valid.")
            else:
                print("✅ Sudah versi terbaru.\n")

        except Exception as e:
            print(f"❌ Gagal cek versi: {e}")

    def download_file(self):
        # Download EXE
        try:
            exe_path = os.path.join(self.download_dir, "bikin_ScoreCard.exe")
            r = requests.get(self.url_exe, timeout=10)
            r.raise_for_status()
            with open(exe_path, "wb") as f:
                f.write(r.content)
            print(f"✅ Download EXE: {exe_path}")
        except Exception as e:
            print(f"❌ Gagal download EXE: {e}")

        # Download depots.txt
        try:
            txt_path = os.path.join(self.download_dir, "depots.txt")
            r = requests.get(self.url_txt, timeout=10)
            r.raise_for_status()
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(r.text)
            print(f"✅ Download TXT: {txt_path}")
        except Exception as e:
            print(f"❌ Gagal download TXT: {e}")

class ScoreCardApp:
    def copy_folder(self, source_folder, dest_folder):
        if os.path.exists(dest_folder):
            print(f"⚠️ Folder tujuan '{dest_folder}' sudah ada.")
        else:
            shutil.copytree(source_folder, dest_folder)
            print(f"✅ Copy: {source_folder} ➜ {dest_folder}")

    def rename_file_folder(self, folder_path, source_name, new_name):
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.endswith(('.xlsx', '.xlsm', '.xlsb')):
                    old_path = os.path.join(root, file)
                    name, ext = os.path.splitext(file)
                    if source_name in name:
                        new_file = name.replace(source_name, new_name) + ext
                        os.rename(old_path, os.path.join(root, new_file))
                        print(f"  ✅ Rename: {file} ➜ {new_file}")

    def ganti_link(self, folder_path, source_name, new_name):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        for file in os.listdir(folder_path):
            if file.startswith("~$") or not file.endswith(('.xlsb', '.xlsx', '.xlsm')):
                continue
            file_path = os.path.join(folder_path, file)
            print(f"\n🔗 Buka file: {file_path}")
            try:
                wb = excel.Workbooks.Open(file_path)
                links = wb.LinkSources()
                if links:
                    for link in links:
                        if source_name in link:
                            new_link = link.replace(source_name, new_name)
                            wb.ChangeLink(Name=link, NewName=new_link, Type=1)
                            print(f"    🔁 Ganti: {link} ➜ {new_link}")
                wb.Save()
                wb.Close(False)
            except Exception as e:
                print(f"    ⚠️ Error: {e}")

        excel.Quit()

    def baca_file_config(self, path):
        with open(path, "r") as f:
            lines = [line.strip() for line in f if line.strip()]

        if not lines[0].lower().startswith("copied:"):
            raise ValueError("Baris pertama harus 'copied: <nama_folder>'")

        source_folder = lines[0].split(":", 1)[1].strip()
        try:
            idx = lines.index("the_copy:")
        except ValueError:
            raise ValueError("Harus ada baris 'the_copy:'")

        return source_folder, lines[idx + 1:]

# =============================
# MAIN PROGRAM
# =============================
if __name__ == "__main__":
    # Cek versi
    git = GitHelper(versi_lokal="1.0.1")
    git.cek_versi()

    # Jalankan aplikasi utama
    app = ScoreCardApp()
    root = os.getcwd()
    config_path = "depots.txt"

    source_name, target_folders = app.baca_file_config(config_path)
    source_path = os.path.join(root, source_name)

    print(f"📁 Sumber: {source_name}")
    print(f"📁 Target: {target_folders}")

    for folder in target_folders:
        dest = os.path.join(root, folder)

        print(f"\n📌 Proses folder: {folder}")
        app.copy_folder(source_path, dest)
        app.rename_file_folder(dest, source_name, folder)
        app.ganti_link(dest, source_name, folder)

    print("\n🎉 Semua selesai!")
    input("Tekan Enter untuk keluar...")
