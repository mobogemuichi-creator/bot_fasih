import uiautomator2 as u2
import subprocess
import os
import time

LDPLAYER_ADB = r"C:\LDPlayer\LDPlayer9\adb.exe"

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

def main():
    d = hubungkan_adb()
    if not d:
        return

    print("\n--- MULAI PENGETESAN IKON BURGER ---")
    
    # Cetak resolusi layar saat ini
    w, h = d.window_size()
    print(f"Resolusi Layar Emulator: {w} x {h}")
    
    # 1. Mencari tombol burger menggunakan uiautomator2 selectors
    print("\n[INFO] Mencari ikon menu menggunakan uiautomator2 selectors...")
    
    menu_btn = d(description="Open navigation drawer")
    if menu_btn.exists():
        print(f"-> Ditemukan dengan description='Open navigation drawer'. Info: {menu_btn.info}")
    else:
        print("-> Tidak ditemukan dengan description='Open navigation drawer'.")
        
    menu_btn_up = d(description="Navigate up")
    if menu_btn_up.exists():
        print(f"-> Ditemukan dengan description='Navigate up'. Info: {menu_btn_up.info}")
    else:
        print("-> Tidak ditemukan dengan description='Navigate up'.")
        
    # Cari ImageButton di dalam toolbar
    image_buttons = []
    try:
        # Iterasi elemen secara langsung menggunakan uiautomator2
        for btn in d(className="android.widget.ImageButton"):
            image_buttons.append(btn)
    except Exception as e:
        print(f"[WARNING] Gagal iterasi ImageButton: {e}")
        
    print(f"-> Jumlah ImageButton yang ditemukan di layar: {len(image_buttons)}")
    for idx, btn in enumerate(image_buttons):
        try:
            print(f"   ImageButton [{idx}]: info = {btn.info}")
        except Exception:
            pass

    # 2. Percobaan Klik Ikon Burger
    print("\n[PERCOBAAN] Mengetuk ikon burger...")
    
    # Coba deskripsi default
    target_btn = d(description="Open navigation drawer")
    if not target_btn.exists():
        target_btn = d(description="Navigate up")
    if not target_btn.exists():
        # Pilih ImageButton yang posisinya paling kiri-atas (biasanya koordinat X < 150 dan Y < 150)
        for btn in image_buttons:
            try:
                bounds = btn.info.get('bounds')
                if bounds and bounds['left'] < 150 and bounds['top'] < 150:
                    target_btn = btn
                    print("-> Memilih ImageButton kiri-atas secara otomatis berdasarkan koordinat bounds.")
                    break
            except Exception:
                continue

    if target_btn.exists():
        print("-> Klik target_btn menggunakan click()...")
        target_btn.click()
    else:
        # Klik menggunakan koordinat dinamis berdasarkan posisi judul toolbar
        click_x = 70
        click_y = 100
        try:
            toolbar_title = d(textContains="GCPLN")
            if toolbar_title.exists():
                bounds = toolbar_title.info.get('bounds')
                click_x = bounds['left'] // 2
                click_y = bounds['top'] + (bounds['bottom'] - bounds['top']) // 2
                print(f"-> Berhasil menghitung koordinat dinamis: ({click_x}, {click_y})")
            else:
                print("-> Judul toolbar 'GCPLN' tidak ditemukan. Menggunakan fallback.")
        except Exception as e:
            print(f"-> Gagal menghitung koordinat dinamis: {e}. Menggunakan fallback.")
            
        print(f"-> Melakukan click pada koordinat: ({click_x}, {click_y})")
        d.click(click_x, click_y)
        
    time.sleep(1.0)
    
    # 3. Verifikasi apakah menu/drawer terbuka
    print("\n[VERIFIKASI] Memeriksa apakah menu sidebar berhasil terbuka...")
    # Menu drawer biasanya ditandai dengan munculnya opsi BLOK II/BLOK III atau judul
    is_opened = False
    for pattern in ["BLOK I", "BLOK II", "BLOK III", "GCPLN26.PRA"]:
        el = d(text=pattern)
        if not el.exists():
            el = d(textContains=pattern)
        if el.exists():
            print(f"-> Terdeteksi elemen '{pattern}': {el.info.get('text')}")
            is_opened = True
            
    if is_opened:
        print("[SUKSES] Menu sidebar berhasil terbuka!")
        
        # 4. Ketuk pilihan 'BLOK II' jika drawer terbuka
        print("\n[PERCOBAAN] Mengetuk pilihan 'BLOK II' pada menu sidebar...")
        blok2_menu_opt = d(text="BLOK II")
        if blok2_menu_opt.exists(timeout=5):
            blok2_menu_opt.click()
            print("-> Opsi 'BLOK II' berhasil diketuk!")
            time.sleep(2.0)
            
            # Verifikasi transisi halaman ke BLOK II
            if d(text="BLOK II").exists():
                print("[SUKSES] Berhasil masuk ke halaman 'BLOK II'!")
            else:
                print("[GAGAL] Layar tidak berpindah ke halaman 'BLOK II'.")
        else:
            print("[ERROR] Pilihan menu 'BLOK II' tidak ditemukan pada sidebar.")
    else:
        print("[GAGAL] Menu sidebar tidak terbuka.")

    print("\n==================================================")
    print("PENGETESAN BURGER ICON SELESAI!")
    print("==================================================")

if __name__ == "__main__":
    main()
