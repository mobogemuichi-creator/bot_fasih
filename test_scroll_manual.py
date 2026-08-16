import os
import time
import subprocess
import uiautomator2 as u2

LDPLAYER_ADB = r"C:\LDPlayer\LDPlayer9\adb.exe"

def connect_emulator():
    """Menghubungkan ke emulator via ADB & uiautomator2"""
    print("Menghubungkan ke emulator...")
    d = None
    for port in ["5555", "5554"]:
        try:
            if os.path.exists(LDPLAYER_ADB):
                subprocess.run([LDPLAYER_ADB, "connect", f"127.0.0.1:{port}"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            temp_d = u2.connect(f"127.0.0.1:{port}")
            _ = temp_d.info
            d = temp_d
            print(f"[KONEKSI] Berhasil terhubung ke emulator di port {port}!")
            break
        except Exception:
            continue

    if not d:
        try:
            d = u2.connect()
            print("[KONEKSI] Berhasil terhubung menggunakan u2.connect() standar.")
        except Exception as e:
            print(f"[ERROR] Gagal terhubung ke emulator: {e}")
            return None
    return d

def main():
    d = connect_emulator()
    if not d:
        input("\nTekan ENTER untuk keluar...")
        return

    w, h = d.window_size()
    print(f"[INFO] Ukuran Layar Emulator: {w} x {h}")

    while True:
        print("\n" + "="*50)
        print("    PENGETESAN SCROLL MANUAL EMULATOR (u2)")
        print("="*50)
        print("1. Scroll Down - Swipe Lambat (d.swipe - duration)")
        print("2. Scroll Up   - Swipe Lambat (d.swipe - duration)")
        print("3. Scroll Forward (d(scrollable=True).scroll.forward - steps)")
        print("4. Scroll Backward (d(scrollable=True).scroll.backward - steps)")
        print("5. Scroll to Text (d(scrollable=True).scroll.to(text=...))")
        print("6. Custom Swipe (Masukan koordinat & durasi bebas)")
        print("0. Keluar")
        print("="*50)
        
        pilihan = input("Pilih menu (0-6): ").strip()

        if pilihan == "1":
            dur = input("Masukkan durasi dalam detik [Default: 0.6]: ").strip()
            durasi = float(dur) if dur else 0.6
            times = input("Berapa kali scroll? [Default: 1]: ").strip()
            jumlah = int(times) if times else 1

            print(f"\n[AKSI] Men-scroll KE BAWAH {jumlah}x dengan durasi {durasi}s...")
            for i in range(jumlah):
                d.swipe(w // 2, int(h * 0.8), w // 2, int(h * 0.2), duration=durasi)
                print(f"  - Scroll ke-{i+1} selesai")
                time.sleep(0.3)

        elif pilihan == "2":
            dur = input("Masukkan durasi dalam detik [Default: 0.6]: ").strip()
            durasi = float(dur) if dur else 0.6
            times = input("Berapa kali scroll? [Default: 1]: ").strip()
            jumlah = int(times) if times else 1

            print(f"\n[AKSI] Men-scroll KE ATAS {jumlah}x dengan durasi {durasi}s...")
            for i in range(jumlah):
                d.swipe(w // 2, int(h * 0.2), w // 2, int(h * 0.8), duration=durasi)
                print(f"  - Scroll ke-{i+1} selesai")
                time.sleep(0.3)

        elif pilihan == "3":
            st = input("Masukkan jumlah steps [Default: 15, makin besar makin lambat]: ").strip()
            steps = int(st) if st else 15
            try:
                print(f"\n[AKSI] Executing scroll.forward(steps={steps})...")
                res = d(scrollable=True).scroll.forward(steps=steps)
                print(f"[RESULT] Status scroll: {res}")
            except Exception as e:
                print(f"[ERROR] Gagal scroll.forward: {e}")

        elif pilihan == "4":
            st = input("Masukkan jumlah steps [Default: 15, makin besar makin lambat]: ").strip()
            steps = int(st) if st else 15
            try:
                print(f"\n[AKSI] Executing scroll.backward(steps={steps})...")
                res = d(scrollable=True).scroll.backward(steps=steps)
                print(f"[RESULT] Status scroll: {res}")
            except Exception as e:
                print(f"[ERROR] Gagal scroll.backward: {e}")

        elif pilihan == "5":
            target_text = input("Masukkan teks target (misal: 'Cek Nomor Meter'): ").strip()
            if target_text:
                try:
                    print(f"\n[AKSI] Men-scroll ke teks '{target_text}'...")
                    found = d(scrollable=True).scroll.to(text=target_text)
                    print(f"[RESULT] Ditemukan: {found}")
                except Exception as e:
                    print(f"[ERROR] Gagal scroll.to: {e}")

        elif pilihan == "6":
            try:
                fx = int(input(f"From X [0-{w}, Default {w//2}]: ") or (w//2))
                fy = int(input(f"From Y [0-{h}, Default {int(h*0.8)}]: ") or int(h*0.8))
                tx = int(input(f"To X   [0-{w}, Default {w//2}]: ") or (w//2))
                ty = int(input(f"To Y   [0-{h}, Default {int(h*0.2)}]: ") or int(h*0.2))
                durasi = float(input("Duration (detik) [Default 0.6]: ") or 0.6)

                print(f"\n[AKSI] Custom swipe: ({fx}, {fy}) -> ({tx}, {ty}) duration={durasi}s")
                d.swipe(fx, fy, tx, ty, duration=durasi)
                print("[RESULT] Swipe selesai.")
            except Exception as e:
                print(f"[ERROR] Input tidak valid: {e}")

        elif pilihan == "0":
            print("\nPengujian selesai. Sampai jumpa!")
            break
        else:
            print("Pilihan tidak valid.")

if __name__ == "__main__":
    main()
