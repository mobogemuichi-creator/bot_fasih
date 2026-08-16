import uiautomator2 as u2
import os
import time
import subprocess

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

    print("\n[STEP 1] Mengisi catatan 'Catatan'...")
    label_catatan = d(textContains="Catatan")
    if not label_catatan.exists():
        print("Mencari label Catatan...")
        try:
            d(scrollable=True).scroll.to(textContains="Catatan")
        except Exception:
            pass
            
    if not label_catatan.exists():
        for _ in range(3):
            d.swipe(540, 1200, 540, 600, duration=0.2)
            time.sleep(0.5)
            if d(textContains="Catatan").exists():
                label_catatan = d(textContains="Catatan")
                break

    input_catatan = None
    if label_catatan.exists():
        for attempt in range(10):
            input_catatan = label_catatan.down(className="android.widget.EditText")
            if input_catatan and input_catatan.exists():
                break
            time.sleep(0.5)
            
    if input_catatan and input_catatan.exists():
        input_catatan.set_text("Sudah Diisi")
        print("Berhasil mengisi Catatan: 'Sudah Diisi'")
        time.sleep(1)
    else:
        # Alternatif jika down EditText gagal, cari EditText pertama
        print("[WARNING] EditText di bawah Catatan tidak ditemukan. Mencoba mengetik di EditText pertama di layar...")
        first_edit = d(className="android.widget.EditText")
        if first_edit.exists():
            first_edit.set_text("Sudah Diisi")
            print("Berhasil mengisi Catatan (via first EditText): 'Sudah Diisi'")
            time.sleep(1)
        else:
            print("[ERROR] Input text box Catatan tidak ditemukan.")

    print("\n[STEP 2] Mengetuk tombol 'Ambil Waktu'...")
    ambil_waktu_btn = d(text="Ambil Waktu")
    if not ambil_waktu_btn.exists():
        try:
            d(scrollable=True).scroll.to(text="Ambil Waktu")
        except Exception:
            pass
            
    if ambil_waktu_btn.exists(timeout=5):
        ambil_waktu_btn.click()
        print("Berhasil mengetuk 'Ambil Waktu'")
        time.sleep(1)
    else:
        print("[ERROR] Tombol 'Ambil Waktu' tidak ditemukan.")

    print("\n[STEP 3] Mengetuk konfirmasi 'Ya'...")
    ya_btn = d(text="Ya")
    if not ya_btn.exists(): ya_btn = d(text="YA")
    if not ya_btn.exists(): ya_btn = d(text="ya")
    if ya_btn.exists(timeout=5):
        ya_btn.click()
        print("Berhasil mengetuk 'Ya'")
        time.sleep(1)
    else:
        print("[WARNING] Konfirmasi dialog 'Ya' tidak muncul.")

    print("\n[STEP 4] Mengetuk tombol 'Kirim' pertama...")
    kirim_btn = d(text="Kirim")
    if not kirim_btn.exists(): kirim_btn = d(text="KIRIM")
    if not kirim_btn.exists(): kirim_btn = d(textContains="Kirim")
    
    if not kirim_btn.exists():
        try:
            d(scrollable=True).scroll.to(text="Kirim")
        except Exception:
            pass
            
    if kirim_btn.exists(timeout=5):
        kirim_btn.click()
        print("Berhasil mengetuk tombol 'Kirim' pertama")
        time.sleep(2)
    else:
        print("[ERROR] Tombol 'Kirim' pertama tidak ditemukan.")

    print("\n[STEP 5] Mengetuk tombol 'Kirim' kedua...")
    # Polling tunggu teks "GALAT 0 Perlu diperbaiki" muncul (maksimal 5 detik)
    for attempt in range(10):
        if d(textContains="GALAT 0 Perlu diperbaiki").exists() or d(textContains="GALAT 0").exists():
            print("[STEP 5] Menemukan teks 'GALAT 0 Perlu diperbaiki'!")
            break
        time.sleep(0.5)
        
    kirim_btn2 = d(text="Kirim", className="android.widget.Button")
    if not kirim_btn2.exists(): kirim_btn2 = d(text="KIRIM", className="android.widget.Button")
    if not kirim_btn2.exists(): kirim_btn2 = d(textContains="Kirim", className="android.widget.Button")
    
    if kirim_btn2.exists(timeout=5):
        kirim_btn2.click()
        print("Berhasil mengetuk tombol 'Kirim' kedua")
        time.sleep(2)
    else:
        print("[WARNING] Tombol 'Kirim' kedua tidak ditemukan/tidak muncul.")

    print("\n[STEP 6] Mengetuk tombol 'Konfirmasi'...")
    konfirmasi_btn = d(text="Konfirmasi", className="android.widget.Button")
    if not konfirmasi_btn.exists():
        konfirmasi_btn = d(textContains="Konfirmasi", className="android.widget.Button")
        
    if konfirmasi_btn.exists(timeout=5):
        konfirmasi_btn.click()
        print("Berhasil mengetuk tombol 'Konfirmasi'")
        time.sleep(1)
        # Tunggu loading/progress dialog selesai
        print("Menunggu loading selesai...")
        try:
            d(resourceId="id.go.bpsfasih:id/card_progress").wait_gone(timeout=20)
        except Exception:
            pass
        time.sleep(2)  # Jeda pengaman agar dialog konfirmasi submit muncul dengan stabil
    else:
        print("[ERROR] Tombol 'Konfirmasi' tidak ditemukan.")

    print("\n[STEP 7] Mengetuk tombol 'YA'...")
    ya_submit_btn = d(resourceId="id.go.bpsfasih:id/rButton_bottomDialog")
    if not ya_submit_btn.exists():
        ya_submit_btn = d(text="YA", className="android.widget.Button")
    if not ya_submit_btn.exists():
        ya_submit_btn = d(textContains="YA")

    if ya_submit_btn.exists(timeout=5):
        ya_submit_btn.click()
        print("Berhasil mengetuk tombol 'YA' Submit")
        time.sleep(2)
    else:
        print("[ERROR] Tombol 'YA' Submit tidak ditemukan.")

    print("\n[STEP 8] Menunggu loading submit selesai...")
    try:
        # Tunggu loading progress bar selesai jika muncul
        d(resourceId="id.go.bpsfasih:id/card_progress").wait_gone(timeout=30)
    except Exception:
        pass
        
    print("Menunggu halaman 'Daftar Assignment' termuat...")
    daftar_assignment_title = d(resourceId="id.go.bpsfasih:id/title_toolbar", text="Daftar Assignment")
    if daftar_assignment_title.wait(exists=True, timeout=30):
        print("[SUKSES] Berhasil kembali ke halaman 'Daftar Assignment'")
    else:
        print("[WARNING] Halaman 'Daftar Assignment' tidak terdeteksi setelah 30 detik.")

    print("\n==================================================")
    print("PROSES PENGETESAN BLOK IV SELESAI!")
    print("==================================================")

if __name__ == "__main__":
    main()
