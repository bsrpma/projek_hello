import os
import shutil
import win32com.client

def copy_folder(source_folder, dest_folder):
    if os.path.exists(dest_folder):
        print(f"Folder tujuan '{dest_folder}' sudah ada. Silakan hapus dulu atau ganti nama.")
    else:
        shutil.copytree(source_folder, dest_folder)
        print(f"✅ Berhasil menyalin folder '{source_folder}' ke '{dest_folder}'")

def rename_file_folder(folder_path, source_name, new_name):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(('.xlsx', '.xlsm', '.xlsb')):
                old_file_path = os.path.join(root, file)
                file_name, file_ext = os.path.splitext(file)

                if source_name in file_name:
                    new_file_name = file_name.replace(source_name, new_name)
                    new_file_path = os.path.join(root, new_file_name + file_ext)

                    os.rename(old_file_path, new_file_path)
                    print(f"  Rename file: {file} => {new_file_name + file_ext}")

def ganti_link(folder_path, source_name, new_name):
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

            if links is not None:
                for link in links:
                    if source_name in link:
                        new_link = link.replace(source_name, new_name)
                        print(f"    Ganti link: {link}\n      ==> {new_link}")
                        wb.ChangeLink(Name=link, NewName=new_link, Type=1)

            wb.Save()
            wb.Close(False)
            print("    ✅ Link sudah diganti & file disimpan.")

        except Exception as e:
            print(f"    ⚠️ Error: {e}")

    excel.Quit()

def baca_file_config(file_path):
    """
    Membaca source folder dan daftar folder tujuan dari file txt
    Format file:
        copied: BBG
        coping:
        BJR
        BMA
        ...
    """
    with open(file_path, "r") as f:
        lines = [line.strip() for line in f if line.strip()]

    # Baris pertama = copied: BBG
    source_line = lines[0]
    if not source_line.lower().startswith("copied:"):
        raise ValueError("Format file salah. Baris pertama harus 'copied: <nama_folder>'")

    source_folder = source_line.split(":", 1)[1].strip()

    # Cari baris "coping:"
    try:
        coping_index = lines.index("coping:")
    except ValueError:
        raise ValueError("Format file salah. Harus ada baris 'coping:'")

    # Daftar folder dimulai setelah 'coping:'
    folders = lines[coping_index + 1:]

    return source_folder, folders


if __name__ == "__main__":
    root_folder = os.getcwd()  # ganti dari fixed ke lokasi sekarang
    config_file_path = "depots.txt"  # nama file txt

    source_folder_name, folders = baca_file_config(config_file_path)
    source_folder_path = os.path.join(root_folder, source_folder_name)

    print(f"📄 Source folder: {source_folder_name}")
    print(f"📄 Folder tujuan: {folders}")

    for folder in folders:
        dest_folder = os.path.join(root_folder, folder)

        # Step 1: Copy jika belum ada
        print(f"\n=== Copy folder untuk {folder} ===")
        copy_folder(source_folder_path, dest_folder)

        # Step 2: Rename file
        print(f"\n=== Rename file di folder {folder} ===")
        rename_file_folder(dest_folder, source_folder_name, folder)

        # Step 3: Ganti link
        print(f"\n=== Ganti link di folder {folder} ===")
        ganti_link(dest_folder, source_folder_name, folder)

    print("\n=== 🎉 Semua selesai! ===")
    input("Selamat, ScoreCard Anda sudah selesai! ^o^v Tekan Enter untuk keluar...")
