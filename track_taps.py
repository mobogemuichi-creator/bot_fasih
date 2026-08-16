import subprocess
import os
import re
import sys

LDPLAYER_ADB = r"C:\LDPlayer\LDPlayer9\adb.exe"

def cari_adb():
    if os.path.exists(LDPLAYER_ADB):
        return LDPLAYER_ADB
    # Coba cari adb global
    return "adb"

def main():
    adb_path = cari_adb()
    
    # 1. Hubungkan ke emulator
    print("Menghubungkan ke emulator port 5555...")
    subprocess.run([adb_path, "connect", "127.0.0.1:5555"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    # Dapatkan resolusi layar
    try:
        res_output = subprocess.check_output([adb_path, "-s", "127.0.0.1:5555", "shell", "wm", "size"]).decode("utf-8")
        match = re.search(r"Physical size: (\d+)x(\d+)", res_output)
        if match:
            screen_w = int(match.group(1))
            screen_h = int(match.group(2))
            print(f"Resolusi Layar Emulator: {screen_w} x {screen_h}")
        else:
            print("Gagal mendeteksi resolusi layar, default 1080x1920.")
            screen_w, screen_h = 1080, 1920
    except Exception:
        screen_w, screen_h = 1080, 1920
        
    print("\n--- MULAI PELACAKAN SENTUHAN REAL-TIME ---")
    print("Silakan klik/ketuk di layar emulator LDPlayer.")
    print("Tekan CTRL+C di terminal ini untuk berhenti.")
    print("-------------------------------------------\n")

    # Jalankan adb shell getevent -l
    # -l untuk menampilkan nama event dalam bentuk teks (misal: ABS_MT_POSITION_X)
    cmd = [adb_path, "-s", "127.0.0.1:5555", "shell", "getevent", "-l"]
    
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    
    x_val = None
    y_val = None
    
    try:
        for line in proc.stdout:
            # Contoh baris: /dev/input/event1: EV_ABS       ABS_MT_POSITION_X    0000021c
            if "ABS_MT_POSITION_X" in line:
                parts = line.strip().split()
                if len(parts) >= 4:
                    hex_val = parts[-1]
                    x_val = int(hex_val, 16)
                    
            elif "ABS_MT_POSITION_Y" in line:
                parts = line.strip().split()
                if len(parts) >= 4:
                    hex_val = parts[-1]
                    y_val = int(hex_val, 16)
                    
            elif "SYN_REPORT" in line:
                # Cetak koordinat jika keduanya X dan Y terisi
                if x_val is not None and y_val is not None:
                    # Di beberapa emulator getevent koordinatnya bisa terbalik/perlu disesuaikan.
                    # Namun LDPlayer biasanya langsung 1:1.
                    print(f"[KLIK DETEKSI] Koordinat Layar -> X: {x_val}, Y: {y_val}")
                    # Reset setelah dilaporkan agar tidak duplikat
                    x_val = None
                    y_val = None
                    
    except KeyboardInterrupt:
        print("\nPelacakan dihentikan oleh pengguna.")
    finally:
        proc.terminate()

if __name__ == "__main__":
    main()
