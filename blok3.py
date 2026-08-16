import uiautomator2 as u2
import subprocess
import os
import time
import re

LDPLAYER_ADB = r"C:\LDPlayer\LDPlayer9\adb.exe"

# Jeda Waktu (Sleep Durations) - Setelan Asli yang Aman
SLEEP_SHORT = 0.2
SLEEP_MEDIUM = 0.5
SLEEP_LONG = 1.0

def hubungkan_adb():
    print("Mencoba menghubungkan ADB ke emulator...")
    for port in ["5554", "5555"]:
        if os.path.exists(LDPLAYER_ADB):
            subprocess.run([LDPLAYER_ADB, "connect", f"127.0.0.1:{port}"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        try:
            d = u2.connect(f"127.0.0.1:{port}")
            device_info = d.info
            print(f"[SUKSES] Terhubung ke emulator di port {port}!")
            return d
        except Exception:
            continue
    print("[GAGAL] Tidak dapat terhubung ke emulator.")
    return None

def baca_temp_alamat():
    # Default fallbacks jika file tidak ditemukan
    data = {
        "Provinsi": "[51] BALI",
        "Kabupaten": "[5103] KAB. BADUNG",
        "Kecamatan": "[510306] KUTA UTARA",
        "Desa/Kelurahan": "[5103061002] KEROBOKAN",
        "Alamat": "JL. GN. MANDALA"
    }
    
    file_path = "temp_alamat.txt"
    if os.path.exists(file_path):
        print(f"[EXCEL/TXT] Membaca data alamat dari {file_path}...")
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                for line in f:
                    if ":" in line:
                        key, val = line.split(":", 1)
                        key = key.strip()
                        val = val.strip()
                        if "Desa" in key or "Kelurahan" in key:
                            data["Desa/Kelurahan"] = val
                        elif "Kabupaten" in key or "Kota" in key:
                            data["Kabupaten"] = val
                        elif key in data:
                            data[key] = val
            print("[SUKSES] Berhasil membaca data alamat:")
            for k, v in data.items():
                print(f"  - {k}: '{v}'")
        except Exception as e:
            print(f"[WARNING] Gagal membaca {file_path}: {e}. Menggunakan default.")
    else:
        print(f"[WARNING] File {file_path} tidak ditemukan. Menggunakan nilai default.")
        
    return data

def ekstrak_nama_saja(teks):
    # Hilangkan bagian "[...]" beserta kurung sikunya jika ada
    clean = re.sub(r'\[.*?\]', '', teks)
    return clean.strip()

def klik_opsi_dropdown(d, teks_pilihan):
    nama_saja = ekstrak_nama_saja(teks_pilihan)
    print(f"[DROPDOWN] Mencari opsi untuk clean name: '{nama_saja}' (dari '{teks_pilihan}')...")
    
    for attempt in range(8):
        try:
            xpath_list = [
                f"//android.widget.ListView//*[contains(@text, '{teks_pilihan}')]",
                f"//android.widget.ListView//*[contains(@text, '{nama_saja}') and contains(@text, '[')]",
                f"//android.widget.ListView//*[contains(@text, '{nama_saja}')]",
                f"//android.app.Dialog//*[contains(@text, '{teks_pilihan}')]",
                f"//android.app.Dialog//*[contains(@text, '{nama_saja}') and contains(@text, '[')]",
                f"//android.app.Dialog//*[contains(@text, '{nama_saja}')]",
            ]
            
            for xp in xpath_list:
                xpath_el = d.xpath(xp)
                if xpath_el.exists:
                    matching_elements = xpath_el.all()
                    target_el = None
                    
                    for el in matching_elements:
                        txt = el.text or ""
                        if "EditText" in (el.info.get("className", "") or ""):
                            continue
                        if txt == teks_pilihan or txt.endswith(f"] {nama_saja}") or txt.endswith(f" {nama_saja}") or txt == nama_saja:
                            target_el = el
                            break
                    
                    if not target_el:
                        for el in matching_elements:
                            if "EditText" not in (el.info.get("className", "") or ""):
                                target_el = el
                                break
                    
                    if target_el:
                        print(f"[DROPDOWN] Menemukan opsi via XPath ('{xp}'): '{target_el.text}' -> Mengklik...")
                        target_el.click()
                        time.sleep(0.4)
                        return True
        except Exception as xpath_err:
            print(f"[DROPDOWN] [WARNING] Gagal mengevaluasi XPath: {xpath_err}")
            
        time.sleep(0.5)
            
    print(f"[DROPDOWN] [WARNING] Opsi '{teks_pilihan}' tidak ditemukan di layar. Menutup overlay.")
    d.click(540, 300)  # Klik area netral untuk menutup overlay
    time.sleep(0.5)
    return False

def main():
    d = hubungkan_adb()
    if not d:
        return

    alamat_data = baca_temp_alamat()
    
    provinsi = alamat_data["Provinsi"]
    kabupaten = alamat_data["Kabupaten"]
    kecamatan = alamat_data["Kecamatan"]
    desa = alamat_data["Desa/Kelurahan"]
    alamat_val = alamat_data["Alamat"]

    # ==========================================
    # PENGISIAN BLOK III (KETERANGAN TEMPAT)
    # ==========================================
    print("\n[BLOK III] Memulai pemrosesan BLOK III...")

    # 1. Cek text "BLOK III"
    print("\n[BLOK III] [STEP 1] Memeriksa teks 'BLOK III'...")
    blok3_header = d(textContains="BLOK III")
    if blok3_header.exists(timeout=20):
        print("[BLOK III] [SUKSES] Berada di halaman/bagian 'BLOK III'")
    else:
        print("[BLOK III] [WARNING] Header 'BLOK III' tidak ditemukan di layar saat ini.")

    # Swipe ke atas kemudian kembalikan lagi dan ketuk koordinat statis (996, 588)
    print("[BLOK III] Melakukan force refresh (swipe up & down) dan mengetuk koordinat statis (996, 588)...")
    d.swipe(540, 1200, 540, 600, duration=0.2)
    time.sleep(0.5)
    d.swipe(540, 600, 540, 1200, duration=0.2)
    time.sleep(0.5)
    d.click(996, 588)
    time.sleep(SLEEP_MEDIUM)

    # 2. di label "a. Provinsi" masukkan data Provinsi
    print("\n[BLOK III] [STEP 2] Mengisi 'a. Provinsi'...")
    label_a = d(textContains="Provinsi")
    input_a = None
    for attempt in range(40):
        if label_a.exists():
            input_a = label_a.down(className="android.widget.EditText")
            if input_a and input_a.exists():
                break
        time.sleep(0.5)

    if input_a and input_a.exists():
        provinsi_clean = ekstrak_nama_saja(provinsi)
        input_a.set_text(str(provinsi_clean))
        print(f"[BLOK III] Berhasil mengisi Provinsi: '{provinsi_clean}'")
        time.sleep(1.0)
        klik_opsi_dropdown(d, provinsi)
    else:
        raise Exception("Input text box untuk 'a. Provinsi' tidak ditemukan.")

    # 3. di label "b. Kabupaten/Kota" masukkan data Kabupaten
    print("\n[BLOK III] [STEP 3] Mengisi 'b. Kabupaten/Kota'...")
    label_b = d(textContains="Kabupaten/Kota")
    input_b = None
    for attempt in range(10):
        if label_b.exists():
            input_b = label_b.down(className="android.widget.EditText")
            if input_b and input_b.exists():
                break
        time.sleep(0.5)

    if input_b and input_b.exists():
        kabupaten_clean = ekstrak_nama_saja(kabupaten)
        input_b.set_text(str(kabupaten_clean))
        print(f"[BLOK III] Berhasil mengisi Kabupaten/Kota: '{kabupaten_clean}'")
        time.sleep(1.0)
        klik_opsi_dropdown(d, kabupaten)
    else:
        raise Exception("Input text box untuk 'b. Kabupaten/Kota' tidak ditemukan.")

    # 4. di label "c. Kecamatan" masukkan data Kecamatan
    print("\n[BLOK III] [STEP 4] Mengisi 'c. Kecamatan'...")
    print("[BLOK III] Melakukan swipe ke bawah agar 'c. Kecamatan' terlihat penuh...")
    d.swipe(540, 1200, 540, 600, duration=0.2)
    time.sleep(0.5)

    label_c = d(textContains="Kecamatan")
    input_c = None
    for attempt in range(10):
        if label_c.exists():
            input_c = label_c.down(className="android.widget.EditText")
            if input_c and input_c.exists():
                break
        time.sleep(0.5)

    if input_c and input_c.exists():
        kecamatan_clean = ekstrak_nama_saja(kecamatan)
        input_c.set_text(str(kecamatan_clean))
        print(f"[BLOK III] Berhasil mengisi Kecamatan: '{kecamatan_clean}'")
        time.sleep(1.0)
        klik_opsi_dropdown(d, kecamatan)
    else:
        raise Exception("Input text box untuk 'c. Kecamatan' tidak ditemukan.")

    # 5. di label "d. Desa/Kelurahan" masukkan data Desa/Kelurahan
    print("\n[BLOK III] [STEP 5] Mengisi 'd. Desa/Kelurahan'...")
    print("[BLOK III] Melakukan swipe ke bawah agar 'd. Desa/Kelurahan' terlihat penuh...")
    d.swipe(540, 1200, 540, 600, duration=0.2)
    time.sleep(0.5)

    label_d = d(textContains="Desa/Kelurahan")
    input_d = None
    for attempt in range(10):
        if label_d.exists():
            input_d = label_d.down(className="android.widget.EditText")
            if input_d and input_d.exists():
                break
        time.sleep(0.5)

    if input_d and input_d.exists():
        desa_clean = ekstrak_nama_saja(desa)
        input_d.set_text(str(desa_clean))
        print(f"[BLOK III] Berhasil mengisi Desa/Kelurahan: '{desa_clean}'")
        time.sleep(1.0)
        klik_opsi_dropdown(d, desa)
    else:
        raise Exception("Input text box untuk 'd. Desa/Kelurahan' tidak ditemukan.")

    # 6. di label "e. Alamat" masukkan data Alamat
    print("\n[BLOK III] [STEP 6] Mengisi 'e. Alamat'...")
    d.click(540, 300)
    time.sleep(0.5)

    print("[BLOK III] Melakukan swipe ke bawah agar 'e. Alamat' terlihat penuh...")
    d.swipe(540, 1200, 540, 600, duration=0.2)
    time.sleep(0.5)

    label_alamat = d(textContains="Alamat")
    input_alamat = None
    for attempt in range(10):
        if label_alamat.exists():
            input_alamat = label_alamat.down(className="android.widget.EditText")
            if input_alamat and input_alamat.exists():
                break
        time.sleep(0.5)

    if input_alamat and input_alamat.exists():
        input_alamat.set_text(str(alamat_val))
        print(f"[BLOK III] Berhasil mengisi Alamat: '{alamat_val}'")
        time.sleep(1)
    else:
        raise Exception("Input text box untuk 'e. Alamat' tidak ditemukan.")

    # 7. ketuk tombol kontrol Increment
    print("\n[BLOK III] [STEP 7] Mengetuk tombol kontrol 'Increment'...")
    
    # Cari tombol Increment dengan toleransi teks
    increment_btn = d(text="Increment")
    if not increment_btn.exists():
        increment_btn = d(textContains="Increment")

    if not increment_btn.exists():
        print("[BLOK III] [STEP 7] Melakukan scroll mencari tombol Increment...")
        try:
            d(scrollable=True).scroll.to(text="Increment")
        except Exception:
            try:
                d(scrollable=True).scroll.to(textContains="Increment")
            except Exception:
                pass

    # Fallback scroll manual jika scroll.to tidak menemukan/gagal
    if not increment_btn.exists():
        for _ in range(3):
            d.swipe(540, 1200, 540, 600, duration=0.2)
            time.sleep(0.5)
            if d(text="Increment").exists():
                increment_btn = d(text="Increment")
                break
            elif d(textContains="Increment").exists():
                increment_btn = d(textContains="Increment")
                break

    if increment_btn.exists():
        increment_btn.click()
        print("[BLOK III] Berhasil mengetuk tombol 'Increment'")
        time.sleep(1)
    else:
        raise Exception("Tombol 'Increment' tidak ditemukan.")

    print("\n==================================================")
    print("[SELESAI] Pengetesan BLOK III selesai dengan sukses!")
    print("==================================================")

if __name__ == "__main__":
    main()
