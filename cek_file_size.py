import os
import tkinter as tk
from tkinter import filedialog

# ==================== KONFIGURASI ====================
# Ukuran file target yang dicari
TARGET_SIZE_KB = 12.7  # Ukuran dalam KB (misalnya 12.7 KB)

# Jika ingin menentukan ukuran secara spesifik dalam Bytes, isi variabel di bawah ini (misal: 13005).
# Jika diset ke None, program akan otomatis menggunakan TARGET_SIZE_KB * 1024.
TARGET_SIZE_BYTES = None  

# Toleransi ukuran (dalam Bytes). 
# Set ke 0 jika ingin pencarian yang 100% persis/exact.
TOLERANCE_BYTES = 500  # Default 500 bytes (sekitar 0.5 KB) untuk mencakup variasi kecil

# Nama file output hasil pencatatan
OUTPUT_FILE = "file_gagal.txt"
# ======================================================

def pilih_folder():
    """Membuka dialog untuk memilih folder secara manual"""
    root = tk.Tk()
    root.withdraw()  # Sembunyikan window utama Tkinter
    root.attributes('-topmost', True)  # Tampilkan dialog di posisi teratas (topmost)
    
    folder_selected = filedialog.askdirectory(title="Pilih Folder untuk Diperiksa")
    return folder_selected

def main():
    # Pilih folder secara manual
    folder_path = pilih_folder()
    if not folder_path:
        print("[INFO] Pemilihan folder dibatalkan.")
        return
        
    print(f"Folder yang dipilih: {folder_path}")

    # Tentukan ukuran target dalam bytes
    if TARGET_SIZE_BYTES is not None:
        target_bytes = TARGET_SIZE_BYTES
        print(f"Mencari file berukuran: {target_bytes} Bytes (toleransi +/- {TOLERANCE_BYTES} Bytes)")
    else:
        target_bytes = int(TARGET_SIZE_KB * 1024)
        print(f"Mencari file berukuran: {TARGET_SIZE_KB} KB (~{target_bytes} Bytes) dengan toleransi +/- {TOLERANCE_BYTES} Bytes")

    min_size = target_bytes - TOLERANCE_BYTES
    max_size = target_bytes + TOLERANCE_BYTES

    matched_files = []

    # Baca seluruh file di dalam folder
    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)
        
        # Pastikan merupakan file (bukan folder)
        if os.path.isfile(filepath):
            file_size = os.path.getsize(filepath)
            
            # Periksa apakah ukuran file masuk dalam batas toleransi
            if min_size <= file_size <= max_size:
                # Dapatkan nama file tanpa ekstensi
                name_without_ext, _ = os.path.splitext(filename)
                matched_files.append((name_without_ext, file_size))

    # Tulis hasil pencatatan ke file TXT
    try:
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            if matched_files:
                for name, size in matched_files:
                    f.write(name + "\n")
                print(f"\n[SUKSES] Berhasil mencatat {len(matched_files)} nama file ke '{OUTPUT_FILE}'")
                for name, size in matched_files:
                    size_kb = size / 1024
                    print(f"  - {name} ({size_kb:.2f} KB / {size} Bytes)")
            else:
                print("\n[INFO] Tidak ditemukan file dengan ukuran tersebut.")
                # Kosongkan file jika sebelumnya ada isinya
                f.truncate(0)
    except Exception as e:
        print(f"[ERROR] Gagal menulis ke file output: {e}")

if __name__ == "__main__":
    main()
