# ==============================================================================
# KONFIGURASI TERPUSAT (BOT & SCRAPPER FASIH)
# ==============================================================================
import os

# ------------------------------------------------------------------------------
# 1. LDPLAYER & EMULATOR CONFIGURATION
# ------------------------------------------------------------------------------
LDPLAYER_DNCONSOLE = r"C:\LDPlayer\LDPlayer9\dnconsole.exe"
LDPLAYER_ADB = r"C:\LDPlayer\LDPlayer9\adb.exe"

# Index emulator LDPlayer untuk instance 1 dan instance 2
EMULATOR_INDEX_1 = "0"  # Digunakan oleh bot_emulator_idpel, bot_emulator_meter, bot_emulator_reject_list
EMULATOR_INDEX_2 = "1"  # Digunakan oleh bot_emulator_idpel2, bot_emulator_meter2

# Port Koneksi ADB / uiautomator2
EMULATOR_PORTS_1 = ["5555", "5555"] # Instance 1 (idpel, meter, reject_list, reject_input)
EMULATOR_PORTS_2 = ["5557", "5557"] # Instance 2 (idpel2, meter2)

# ------------------------------------------------------------------------------
# 2. FILE & DIRECTORY PATHS
# ------------------------------------------------------------------------------
# File Excel sumber data
EXCEL_FILE_1 = "data_tugas.xlsx"       # Instance 1: bot_emulator_idpel, bot_emulator_meter, ss_gmaps
EXCEL_FILE_2 = "data_tugas2.xlsx"      # Instance 2: bot_emulator_idpel2, bot_emulator_meter2, ss_gmaps2
EXCEL_FILE_REJECT = "data_reject.xlsx" # Reject Input: bot_emulator_reject_input

# Path Folder Foto (Lokal PC)
# FOTO_DIRECTORY = r"C:\Users\batan\OneDrive\Documents\XuanZhi9\Pictures"
FOTO_DIRECTORY = r"C:\Users\BaliAga\Documents\XuanZhi9\Pictures"
FOTO_DIR = FOTO_DIRECTORY  # Alias untuk ss_gmaps / ss_gmaps2

# File Log Gambar Hitam (ss_gmaps)
LOG_HITAM_FILE = "idpel_gambar_hitam.txt"

# ------------------------------------------------------------------------------
# 3. JEDA WAKTU (SLEEP DURATIONS IN SECONDS)
# ------------------------------------------------------------------------------
SLEEP_SHORT = 0.1         # Untuk ketik text, ENTER, scroll, radio button
SLEEP_MEDIUM = 0.3        # Untuk klik tombol normal, aksi loading cepat
SLEEP_LONG = 0.5          # Untuk transisi halaman, pemicu kamera/galeri, ambil GPS
SLEEP_LONG_REJECT = 0.6   # Jeda transisi halaman khusus bot_emulator_reject_input

# ------------------------------------------------------------------------------
# 4. GOOGLE MAPS STREETVIEW SCRAPER CONFIGURATION (ss_gmaps & ss_gmaps2)
# ------------------------------------------------------------------------------
# Toggle overwrite gambar: True = jika gambar sudah ada langsung ditimpa | False = skip
OVERWRITE_EXISTING = False  

# Nilai zoom kamera (Field of View dalam derajat). Semakin kecil semakin dekat. Default: 75
ZOOM_FOV = 75  

# Geser sudut kompas kamera (ke kanan/kiri dalam derajat) untuk menggeser watermark. Set 0 untuk sudut asli.
BEARING_OFFSET = 0  

# Sudut dongak kamera ke atas/bawah (93 adalah default dongak ke atas).
PITCH_VALUE = 93  

# ------------------------------------------------------------------------------
# 5. REJECT LIST CONFIGURATION (bot_emulator_reject_list)
# ------------------------------------------------------------------------------
# Formatter kolom output excel reject list: tuple(FIELD_NAME, HEADER_TITLE)
CUSTOM_COLUMNS = [
    ("NO_METER", "No. Meter"),
    ("IDPEL", "ID Pelanggan"),
]

# Filter status yang ingin dicentang di aplikasi FASIH
# Pilihan status filter yang tersedia di dialog aplikasi:
#   - "Open"
#   - "Pernah dibuka"
#   - "Submit"
#   - "Approve"
#   - "Reject"
#
# Masukkan status yang ingin DICENTANG / DIAKTIFKAN ke dalam list di bawah ini.
# Contoh multiple filter: FILTER_STATUS_TARGET = ["Open", "Pernah dibuka"] atau ["Submit", "Approve"]
FILTER_STATUS_TARGET = ["Reject"]
