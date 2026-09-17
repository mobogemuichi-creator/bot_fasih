"""
================================================================================
BOT EMULATOR REJECT LIST - FASIH SCRAPPER (OPTIMIZED & REFACTORED)
================================================================================
Deskripsi:
  Bot otomatisasi ekstraksi daftar status Reject pada aplikasi FASIH BPS
  menggunakan uiautomator2 dan ADB. Dilengkapi pemrosesan hierarki XML yang
  terpadu (1-pass), scroll proporsional adaptif terhadap resolusi layar,
  sistem I/O batch untuk reject.txt, serta sinkronisasi Excel berkecepatan tinggi.
================================================================================
"""

import os
import sys
import time
import re
import subprocess
import datetime
import xml.etree.ElementTree as ET
import openpyxl
import uiautomator2 as u2

from konfigurasi import (
    LDPLAYER_DNCONSOLE,
    LDPLAYER_ADB,
    EMULATOR_INDEX_1 as EMULATOR_INDEX,
    EMULATOR_PORTS_1 as EMULATOR_PORTS,
    SLEEP_SHORT,
    SLEEP_MEDIUM,
    SLEEP_LONG,
    CUSTOM_COLUMNS,
    FILTER_STATUS_TARGET,
)

try:
    from konfigurasi import EXCEL_FILE_REJECT as DEFAULT_EXCEL_FILE
except ImportError:
    DEFAULT_EXCEL_FILE = "data_reject.xlsx"


# ==============================================================================
# 1. EMULATOR DEVICE MANAGER
# ==============================================================================
class EmulatorDevice:
    """Mengelola koneksi ADB, uiautomator2, koordinat layar adaptif, dan gesture aman."""

    def __init__(self, ports=None, adb_path=None):
        self.ports = ports or EMULATOR_PORTS
        self.adb_path = adb_path or LDPLAYER_ADB
        self.d = None
        self.width = 1080
        self.height = 1920

    def connect(self) -> bool:
        """Menghubungkan ke emulator LDPlayer via ADB dan uiautomator2."""
        print("[KONEKSI] Menghubungkan ke emulator...")
        for port in self.ports:
            try:
                if os.path.exists(self.adb_path):
                    subprocess.run(
                        [self.adb_path, "connect", f"127.0.0.1:{port}"],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                temp_d = u2.connect(f"127.0.0.1:{port}")
                info = temp_d.info
                self.d = temp_d

                # Ambil dimensi layar riil untuk perhitungan koordinat dinamis
                self.width = info.get('displayWidth', 1080)
                self.height = info.get('displayHeight', 1920)

                print(f"[KONEKSI] Berhasil terhubung ke emulator di port {port}!")
                print(f"[DISPLAY] Resolusi layar: {self.width}x{self.height}")
                return True
            except Exception:
                continue

        print("[ERROR] Gagal terhubung ke emulator. Pastikan LDPlayer sudah aktif.")
        return False

    def swipe_aman(self, fx: int, fy: int, tx: int, ty: int, duration: float = 0.1):
        """
        Melakukan swipe dengan uiautomator2 dan fallback aman ke ADB shell input swipe.
        Menjamin seluruh koordinat berada dalam batas aman layar (clamped).
        """
        # Proteksi koordinat agar tidak out-of-bounds / negatif
        fx = max(5, min(self.width - 5, int(fx)))
        fy = max(5, min(self.height - 5, int(fy)))
        tx = max(5, min(self.width - 5, int(tx)))
        ty = max(5, min(self.height - 5, int(ty)))

        try:
            self.d.swipe(fx, fy, tx, ty, duration=duration)
        except Exception:
            duration_ms = int(duration * 1000)
            try:
                self.d.shell(f"input swipe {fx} {fy} {tx} {ty} {duration_ms}")
            except Exception as shell_err:
                print(f"[WARNING] Gagal swipe via ADB shell: {shell_err}")

    def swipe_down(self, ratio_start: float = 0.65, ratio_end: float = 0.25, duration: float = 0.1):
        """Swipe ke bawah (konten bergerak ke atas) berbasis persentase tinggi layar."""
        cx = self.width // 2
        sy = int(self.height * ratio_start)
        ey = int(self.height * ratio_end)
        self.swipe_aman(cx, sy, cx, ey, duration=duration)

    def swipe_up(self, ratio_start: float = 0.25, ratio_end: float = 0.70, duration: float = 0.08):
        """Swipe ke atas (konten bergerak ke bawah / kembali ke header) berbasis persentase tinggi layar."""
        cx = self.width // 2
        sy = int(self.height * ratio_start)
        ey = int(self.height * ratio_end)
        self.swipe_aman(cx, sy, cx, ey, duration=duration)

    def dump_hierarchy(self) -> str:
        """Mengambil dump hierarki XML tampilan saat ini dengan penanganan error."""
        try:
            return self.d.dump_hierarchy()
        except Exception as e:
            print(f"[WARNING] Gagal dump hierarki XML: {e}")
            return ""


# ==============================================================================
# 2. TABLE HIERARCHY & REGEX PARSER
# ==============================================================================
class TableHierarchyParser:
    """Parser untuk mengekstrak data baris tabel dan status pagination dari XML UI."""

    IGNORE_KEYWORDS = [
        'daftar assignment', 'search:', 'show', 'entries',
        'previous', 'next', 'filter by', 'showing', 'activate to sort',
        'menampilkan', 'sampai', 'dari'
    ]
    HEADER_LABELS = {'No. Meter', 'Nama', 'ID Pelanggan', 'No.', 'Status'}

    @classmethod
    def extract_pagination(cls, xml_content: str):
        """
        Mengekstrak informasi pagination dari teks 'Showing X to Y of Z entries'
        Returns: tuple (start_idx, end_idx, total_entries, is_last_page)
        """
        match = re.search(
            r'Showing\s+([\d\.,]+)\s+to\s+([\d\.,]+)\s+of\s+([\d\.,]+)\s+entries',
            xml_content,
            re.IGNORECASE
        )
        if not match:
            # Fallback bahasa Indonesia
            match = re.search(
                r'Menampilkan\s+([\d\.,]+)\s+(?:sampai|-)\s+([\d\.,]+)\s+dari\s+([\d\.,]+)',
                xml_content,
                re.IGNORECASE
            )

        if match:
            def clean_num(s):
                return int(re.sub(r'[^\d]', '', s) or 0)

            x = clean_num(match.group(1))
            y = clean_num(match.group(2))
            z = clean_num(match.group(3))
            is_last = (z == 0 or y >= z)
            return x, y, z, is_last

        return None, None, None, False

    @classmethod
    def extract_table_data(cls, xml_content: str):
        """
        Mengekstrak baris data tabel (no_meter, nama, idpel) dari hierarki XML.
        Returns: list of tuple (no_meter, nama, idpel)
        """
        if not xml_content:
            return []

        try:
            tree = ET.fromstring(xml_content)
        except Exception as e:
            print(f"[PARSER WARNING] Gagal membaca XML hierarki: {e}")
            return []

        # 1. Cari container GridView / Table jika hint terdefinisi
        grid_container = None
        for elem in tree.iter('node'):
            hint = elem.attrib.get('hint', '')
            if 'Showing' in hint and 'entries' in hint:
                grid_container = elem
                break

        search_root = grid_container if grid_container is not None else tree

        # 2. Rekursif mengumpulkan node teks non-header
        nodes = []

        def collect_nodes(elem):
            txt = elem.attrib.get('text', '').strip()
            bounds = elem.attrib.get('bounds', '')

            if txt and bounds:
                txt_lower = txt.lower()
                if not any(k in txt_lower for k in cls.IGNORE_KEYWORDS) and txt not in cls.HEADER_LABELS:
                    match = re.search(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', bounds)
                    if match:
                        x1, y1, x2, y2 = map(int, match.groups())
                        nodes.append({
                            'text': txt,
                            'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
                            'yc': (y1 + y2) // 2
                        })
                        return  # Cegah rekursi ke child jika parent sudah memuat teks

            for child in elem:
                collect_nodes(child)

        collect_nodes(search_root)

        # 3. Grouping baris berdasarkan Y-Center (toleransi 20px)
        rows = {}
        for n in nodes:
            found_row = False
            for yc in rows:
                if abs(n['yc'] - yc) <= 20:
                    rows[yc].append(n)
                    found_row = True
                    break
            if not found_row:
                rows[n['yc']] = [n]

        # 4. Parsing data per baris terurut X1 (kiri ke kanan)
        hasil = []
        for yc in sorted(rows.keys()):
            row_nodes = sorted(rows[yc], key=lambda item: item['x1'])
            texts = [rn['text'] for rn in row_nodes]

            no_meter = ""
            nama = ""
            idpel = ""

            for txt in texts:
                clean_digits = re.sub(r'[^\d]', '', txt)
                # ID Pelanggan: Tepat 12 digit angka
                if len(clean_digits) == 12:
                    idpel = clean_digits
                # No. Meter: 1-11 digit angka
                elif 1 <= len(clean_digits) <= 11 and (txt.startswith('+') or clean_digits == txt.replace('+', '').strip()):
                    no_meter = clean_digits
                else:
                    if txt and not any(k in txt.lower() for k in cls.IGNORE_KEYWORDS):
                        nama = txt

            # Baris valid jika IDPEL ditemukan (12 digit angka)
            if idpel:
                hasil.append((no_meter, nama, idpel))

        return hasil

    @classmethod
    def parse_screen(cls, xml_content: str):
        """
        Menggabungkan ekstraksi baris data tabel dan informasi pagination
        hanya dalam 1 kali pembacaan XML hierarki (sangat efisien).
        Returns: tuple (rows, pagination_tuple)
        """
        rows = cls.extract_table_data(xml_content)
        pagination = cls.extract_pagination(xml_content)
        return rows, pagination


# ==============================================================================
# 3. DATA MANAGER (TXT BATCH & EXCEL SYNC)
# ==============================================================================
class DataManager:
    """Mengelola penulisan data batch ke reject.txt dan sinkronisasi ke data_reject.xlsx."""

    @staticmethod
    def init_txt_file(file_path: str = "reject.txt", columns=CUSTOM_COLUMNS):
        """Mengosongkan dan menulis header baru pada reject.txt."""
        header_text = "\t".join([header for _, header in columns]) + "\n"
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(header_text)
        print(f"[FILE] Berhasil mengosongkan '{file_path}' dan menulis header: '{header_text.strip()}'")

    @staticmethod
    def save_batch_txt(file_path: str, data_list, processed_set: set, columns=CUSTOM_COLUMNS):
        """
        Menyimpan baris data baru ke file TXT dalam 1 kali penulisan I/O batch.
        Returns: jumlah baris baru yang ditambahkan
        """
        new_lines = []
        new_records = []

        for no_meter, nama, idpel in data_list:
            if idpel not in processed_set:
                processed_set.add(idpel)
                data_map = {
                    "NO_METER": no_meter,
                    "NAMA": nama,
                    "IDPEL": idpel
                }
                row_values = [str(data_map.get(field, "")) for field, _ in columns]
                new_lines.append("\t".join(row_values) + "\n")
                new_records.append(row_values)
                log_detail = " | ".join([f"{h}: {val}" for (_, h), val in zip(columns, row_values)])
                print(f"  [+ SIMPAN] {log_detail}")
            else:
                print(f"  [SKIP KEMBAR] IDPEL: {idpel} (sudah tersimpan sebelumnya)")

        if new_lines:
            with open(file_path, "a", encoding="utf-8") as f:
                f.writelines(new_lines)

        return len(new_records)

    @staticmethod
    def sync_to_excel(file_txt: str = "reject.txt", file_excel: str = DEFAULT_EXCEL_FILE, columns=CUSTOM_COLUMNS):
        """
        Menyalin seluruh data dari reject.txt ke file data_reject.xlsx
        pada baris paling bawah. Format sel ditetapkan string '@' agar leading zero aman.
        """
        print(f"\n[EXCEL] Menyalin data dari '{file_txt}' ke '{file_excel}'...")
        if not os.path.exists(file_txt):
            print(f"[WARNING] File '{file_txt}' tidak ditemukan. Batal menyalin ke Excel.")
            return False

        # 1. Buka atau buat workbook Excel
        if os.path.exists(file_excel):
            try:
                wb = openpyxl.load_workbook(file_excel)
                ws = wb.active
                print(f"[EXCEL] File '{file_excel}' ditemukan. Menambahkan data di baris terakhir...")
            except Exception as e:
                print(f"[WARNING] Gagal membaca '{file_excel}': {e}. Membuat file Excel baru.")
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Reject"
                ws.cell(row=1, column=1, value="IDPEL")
                ws.cell(row=1, column=2, value="NO_METER")
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Reject"
            ws.cell(row=1, column=1, value="IDPEL")
            ws.cell(row=1, column=2, value="NO_METER")
            print(f"[EXCEL] Membuat file Excel baru '{file_excel}'...")

        # 2. Cari baris terisi terakhir secara cepat
        last_filled_row = 1
        for r in range(ws.max_row, 0, -1):
            val_a = ws.cell(row=r, column=1).value
            val_b = ws.cell(row=r, column=2).value if ws.max_column >= 2 else None
            if (val_a is not None and str(val_a).strip() != "") or (val_b is not None and str(val_b).strip() != ""):
                last_filled_row = r
                break

        target_row = max(2, last_filled_row + 1)
        added_count = 0

        # Mapping posisi kolom dari CUSTOM_COLUMNS
        col_field_map = {field: idx for idx, (field, _) in enumerate(columns)}
        no_meter_idx = col_field_map.get("NO_METER", 0)
        idpel_idx = col_field_map.get("IDPEL", 1)

        try:
            with open(file_txt, "r", encoding="utf-8") as f:
                lines = f.readlines()

            for line in lines[1:]:  # Skip header
                line_str = line.strip()
                if not line_str:
                    continue

                parts = line.rstrip("\r\n").split("\t")
                if len(parts) < len(columns):
                    parts.extend([""] * (len(columns) - len(parts)))

                val_no_meter = parts[no_meter_idx].strip()
                val_idpel = parts[idpel_idx].strip()

                if val_idpel:
                    # Kolom A = IDPEL | Kolom B = NO_METER
                    c_idpel = ws.cell(row=target_row, column=1)
                    c_idpel.value = str(val_idpel)
                    c_idpel.number_format = '@'

                    c_meter = ws.cell(row=target_row, column=2)
                    c_meter.value = str(val_no_meter)
                    c_meter.number_format = '@'

                    target_row += 1
                    added_count += 1

            wb.save(file_excel)
            print(f"[SUKSES] Berhasil menyalin {added_count} baris data dari '{file_txt}' ke '{file_excel}'!")
            print(f"        Kolom A = IDPEL | Kolom B = NO_METER (dimulai dari baris {last_filled_row + 1})")
            return True

        except Exception as e:
            print(f"[ERROR] Gagal menyalin data ke Excel: {e}")
            return False


# ==============================================================================
# 4. FASIH REJECT BOT ORCHESTRATOR
# ==============================================================================
class FasihRejectBot:
    """Orkestrator alur kerja bot scraping data reject."""

    def __init__(self, device: EmulatorDevice):
        self.device = device
        self.d = device.d
        self.parser = TableHierarchyParser()
        self.data_manager = DataManager()
        self.output_txt = "reject.txt"
        self.output_excel = DEFAULT_EXCEL_FILE

    def cek_dan_tangani_kebablasan(self) -> bool:
        """
        Langkah 1: Cek jika layar tersasar ke halaman 'Periode' atau 'Daftar Wilayah'
        dan navigasikan kembali ke 'Daftar Assignment'.
        """
        print("\n[LANGKAH 1] Memeriksa apakah kebablasan ke halaman 'Periode' atau 'Daftar Wilayah'...")
        d = self.device.d
        try:
            is_periode = d(resourceId="id.go.bpsfasih:id/title_toolbar", text="Periode").exists()
            is_wilayah = d(resourceId="id.go.bpsfasih:id/title_toolbar", text="Daftar Wilayah").exists()

            if is_periode or is_wilayah:
                current_page = "Periode" if is_periode else "Daftar Wilayah"
                print(f"[NAV] Terdeteksi kebablasan di halaman '{current_page}'.")

                submit_btn = d(text="Submit")
                if submit_btn.exists(timeout=2):
                    print("[NAV] Mengetuk tombol 'Submit'...")
                    submit_btn.click()
                    time.sleep(SLEEP_LONG)

            # Cek jika berada di halaman SLS Wilayah (updateListingLayout)
            region_row = d(resourceId="id.go.bpsfasih:id/updateListingLayout")
            if region_row.exists():
                print("[NAV] Berada di halaman SLS Wilayah. Mengklik wilayah untuk masuk ke Daftar Assignment...")
                region_row.click()
                time.sleep(SLEEP_LONG)

            # Tunggu sampai halaman Daftar Assignment termuat
            title_el = d(resourceId="id.go.bpsfasih:id/title_toolbar", text="Daftar Assignment")
            if title_el.wait(exists=True, timeout=10.0):
                print("[NAV] Berhasil berada di halaman 'Daftar Assignment'.")
                return True
            else:
                print("[WARNING] Halaman 'Daftar Assignment' tidak terdeteksi secara pasti, melanjutkan...")
                return True
        except Exception as e:
            print(f"[ERROR] Gagal saat memeriksa kebablasan halaman: {e}")
            return False

    def ketuk_fab_dan_filter(self) -> bool:
        """Langkah 2: Mengetuk FAB dan memilih 'Filter By Status'."""
        print("\n[LANGKAH 2] Mengetuk FAB dan memilih 'Filter By Status'...")
        d = self.device.d
        try:
            fab = d(resourceId="id.go.bpsfasih:id/expendable_fab")
            if not fab.exists():
                fab = d(descriptionContains="FAB")

            if fab.wait(exists=True, timeout=5.0):
                print("[FAB] Mengetuk tombol FAB (expendable_fab)...")
                fab.click()
                time.sleep(SLEEP_SHORT)
            else:
                print("[ERROR] Tombol FAB tidak ditemukan di layar.")
                return False

            filter_btn = d(resourceId="id.go.bpsfasih:id/fab_filterAssignment")
            if not filter_btn.exists():
                filter_btn = d(text="Filter By Status")
            if not filter_btn.exists():
                filter_btn = d(textContains="Filter")

            if filter_btn.wait(exists=True, timeout=5.0):
                print("[FAB] Mengetuk opsi 'Filter By Status'...")
                filter_btn.click()
                time.sleep(SLEEP_MEDIUM)
                return True
            else:
                print("[ERROR] Opsi 'Filter By Status' tidak ditemukan setelah mengetuk FAB.")
                return False
        except Exception as e:
            print(f"[ERROR] Gagal saat mengetuk FAB / Filter By Status: {e}")
            return False

    def terapkan_filter_reject(self) -> bool:
        """
        Langkah 3: Mengatur filter status sesuai FILTER_STATUS_TARGET
        lalu mengetuk tombol 'TERAPKAN'.
        """
        print(f"\n[LANGKAH 3] Mengatur filter status ke {FILTER_STATUS_TARGET}...")
        d = self.device.d
        try:
            filter_title = d(text="Filter Assignment By Status")
            if not filter_title.exists():
                print("[INFO] Dialog filter belum terbuka. Membuka dialog filter...")
                if not self.ketuk_fab_dan_filter():
                    print("[ERROR] Gagal membuka dialog filter via FAB.")
                    return False

            all_checkboxes = [
                ("Open", "id.go.bpsfasih:id/open_cb_bottomSheetFilterAssignment"),
                ("Pernah dibuka", "id.go.bpsfasih:id/pernahDibuka_cb_bottomSheetFilterAssignment"),
                ("Submit", "id.go.bpsfasih:id/submit_cb_bottomSheetFilterAssignment"),
                ("Approve", "id.go.bpsfasih:id/approve_cb_bottomSheetFilterAssignment"),
                ("Reject", "id.go.bpsfasih:id/reject_cb_bottomSheetFilterAssignment"),
            ]

            def cari_cb(label, res_id):
                cb = d(resourceId=res_id)
                if cb.exists():
                    return cb
                cb = d(text=label)
                if cb.exists():
                    return cb
                cb = d(textContains=label)
                if cb.exists():
                    return cb
                return None

            for label, res_id in all_checkboxes:
                harus_centang = label in FILTER_STATUS_TARGET
                cb = cari_cb(label, res_id)

                if not cb:
                    print(f"[WARNING] Checkbox '{label}' tidak ditemukan di dialog filter.")
                    continue

                for attempt in range(1, 4):
                    cb = cari_cb(label, res_id)
                    if not cb:
                        break

                    info = cb.info if hasattr(cb, 'info') else {}
                    is_checked = info.get("checked", False)

                    if harus_centang:
                        if not is_checked:
                            print(f"[FILTER] Mencentang checkbox '{label}' (Percobaan {attempt}/3)...")
                            cb.click()
                            time.sleep(0.3)
                        else:
                            print(f"[FILTER] [OK] Checkbox '{label}' sudah tercentang (checked=True).")
                            break
                    else:
                        if is_checked:
                            print(f"[FILTER] Unchecking checkbox '{label}' (Percobaan {attempt}/3)...")
                            cb.click()
                            time.sleep(0.3)
                        else:
                            print(f"[FILTER] [OK] Checkbox '{label}' sudah bersih (checked=False).")
                            break

            btn_terapkan = d(resourceId="id.go.bpsfasih:id/rButton_bottomDialogFilterAssignment")
            if not btn_terapkan.exists():
                btn_terapkan = d(text="TERAPKAN")

            if btn_terapkan.wait(exists=True, timeout=3.0):
                print("[FILTER] Mengetuk tombol 'TERAPKAN'...")
                btn_terapkan.click()
                time.sleep(SLEEP_LONG)
                print(f"[SUKSES] Filter status {FILTER_STATUS_TARGET} berhasil diterapkan!")
                return True
            else:
                print("[ERROR] Tombol 'TERAPKAN' tidak ditemukan.")
                return False

        except Exception as e:
            print(f"[ERROR] Gagal saat mengatur filter status: {e}")
            return False

    def reset_posisi_entri_tabel(self) -> bool:
        """
        Langkah 4: Mereset posisi entri tabel ke Halaman 1 (Showing 1 to Y)
        dan melakukan scroll cepat kembali ke header paling atas.
        """
        print("\n[RESET] Memeriksa & mereset posisi entri tabel ke Halaman 1...")
        d = self.device.d
        try:
            # Periksa apakah sudah di Halaman 1
            xml = self.device.dump_hierarchy()
            x, y, z, _ = self.parser.extract_pagination(xml)
            if x == 1:
                print(f"[RESET] Sudah berada di Halaman 1 (Showing 1 to {y} of {z}).")
            else:
                # Scroll ke bawah untuk mencari tombol Previous jika diperlukan
                is_top = d(textContains="Filter By").exists() or d(textContains="Search").exists()
                if is_top:
                    print("[RESET] Berada di posisi atas, swiping down untuk memunculkan tombol Previous...")
                    for _ in range(3):
                        if d(text="Previous").exists() or d(description="Previous").exists():
                            break
                        self.device.swipe_down(ratio_start=0.65, ratio_end=0.25, duration=0.1)
                        time.sleep(0.2)

                # Loop klik tombol Previous sampai kembali ke Halaman 1
                for attempt in range(1, 15):
                    xml_curr = self.device.dump_hierarchy()
                    x_c, y_c, z_c, _ = self.parser.extract_pagination(xml_curr)
                    if x_c == 1:
                        print(f"[RESET] Berhasil kembali ke Halaman 1 (Showing 1 to {y_c} of {z_c}).")
                        break

                    prev_btn = d(text="Previous")
                    if not prev_btn.exists():
                        prev_btn = d(description="Previous")

                    if prev_btn.exists():
                        is_clickable = prev_btn.info.get("clickable", False)
                        if is_clickable:
                            print(f"[RESET] Mengetuk tombol 'Previous' (Percobaan {attempt})...")
                            prev_btn.click()
                            time.sleep(0.4)
                        else:
                            print("[RESET] Tombol 'Previous' sudah nonaktif (Halaman 1).")
                            break
                    else:
                        self.device.swipe_down(ratio_start=0.60, ratio_end=0.30, duration=0.1)
                        time.sleep(0.2)

            # Swipe up terukur (3-4 kali) untuk kembali ke atas tabel
            print("[SWIPE UP] Scroll cepat kembali ke bagian atas tabel...")
            for _ in range(3):
                self.device.swipe_up(ratio_start=0.20, ratio_end=0.75, duration=0.06)
                time.sleep(0.05)

            time.sleep(0.3)
            return True
        except Exception as e:
            print(f"[ERROR] Gagal saat mereset posisi entri tabel: {e}")
            return False

    def proses_ekstraksi_dan_swipe(self):
        """
        Langkah 5: Mengambil data tabel per halaman sambil swipe adaptif,
        menyimpan ke reject.txt dalam batch, navigasi antar halaman,
        dan menyalin hasil akhir ke data_reject.xlsx.
        """
        print("\n[LANGKAH 5] Memulai ekstraksi data tabel dan swipe dinamis...")
        d = self.device.d
        set_terproses = set()

        # Inisialisasi file output TXT dengan header
        self.data_manager.init_txt_file(self.output_txt, CUSTOM_COLUMNS)

        page_count = 1

        while True:
            print(f"\n==================================================")
            print(f"        MEMPROSES HALAMAN TABEL {page_count}        ")
            print(f"==================================================")

            target_page_label = str(page_count + 1)
            inner_scroll_count = 0
            max_inner_scrolls = 6  # Batas maksimal scroll down per halaman untuk mencegah infinite loop

            while inner_scroll_count < max_inner_scrolls:
                # 1. Ekstrak data dan pagination dalam 1-pass pembacaan XML
                xml_data = self.device.dump_hierarchy()
                rows, pagination = self.parser.parse_screen(xml_data)

                # Simpan baris baru ke buffer dan flush ke reject.txt
                self.data_manager.save_batch_txt(self.output_txt, rows, set_terproses, CUSTOM_COLUMNS)

                _, y_idx, z_idx, is_last = pagination
                if is_last:
                    print(f"[PAGE CHECK] Telah mencapai data terakhir ({y_idx}/{z_idx}). Ini adalah HALAMAN TERAKHIR.")
                    break

                # 2. Cek apakah tombol target ('2', '3', dst) atau 'Next' sudah terlihat di viewport
                target_visible = False
                for label in [target_page_label, "Next"]:
                    btn_el = d(text=label)
                    if not btn_el.exists():
                        btn_el = d(description=label)

                    if btn_el.exists():
                        try:
                            bounds = btn_el.info.get("bounds", {})
                            top = bounds.get("top", 0)
                            bottom = bounds.get("bottom", 0)
                            # Pastikan tombol benar-benar berada dalam rentang layar vertikal
                            if 0 < top < self.device.height and bottom <= self.device.height:
                                target_visible = True
                                break
                        except Exception:
                            target_visible = True
                            break

                if target_visible:
                    print(f"[PAGE {page_count}] Tombol halaman '{target_page_label}' / 'Next' terdeteksi di viewport.")
                    # Ekstrak sekali lagi sebelum beralih
                    final_xml = self.device.dump_hierarchy()
                    final_rows, _ = self.parser.parse_screen(final_xml)
                    self.data_manager.save_batch_txt(self.output_txt, final_rows, set_terproses, CUSTOM_COLUMNS)
                    break

                # 3. Swipe down adaptif persentase layar
                print(f"[SWIPE] Swipe dinamis ke bawah (Scroll {inner_scroll_count + 1}/{max_inner_scrolls})...")
                self.device.swipe_down(ratio_start=0.65, ratio_end=0.25, duration=0.1)
                time.sleep(0.25)
                inner_scroll_count += 1

            # Periksa kembali status halaman terakhir
            curr_xml = self.device.dump_hierarchy()
            _, y_chk, z_chk, is_last_chk = self.parser.extract_pagination(curr_xml)
            if is_last_chk:
                print(f"[INFO] Halaman terakhir terdeteksi ({y_chk}/{z_chk}). Ekstraksi Selesai!")
                break

            # 4. Verifikasi status clickable tombol target
            def is_button_clickable():
                for label in [target_page_label, "Next"]:
                    for selector in [d(description=label), d(text=label)]:
                        if selector.exists():
                            try:
                                info = selector.info
                                if info.get("clickable", False) and info.get("enabled", True):
                                    return True
                            except Exception:
                                pass
                return False

            clickable = is_button_clickable()

            # Jika belum clickable, coba swipe sedikit lagi (max 2 kali)
            if not clickable:
                print(f"[PAGE {page_count}] Tombol target belum clickable, mencoba swipe tambahan...")
                for retry in range(1, 3):
                    self.device.swipe_down(ratio_start=0.55, ratio_end=0.35, duration=0.1)
                    time.sleep(0.2)
                    retry_xml = self.device.dump_hierarchy()
                    r_rows, _ = self.parser.parse_screen(retry_xml)
                    self.data_manager.save_batch_txt(self.output_txt, r_rows, set_terproses, CUSTOM_COLUMNS)

                    if is_button_clickable():
                        clickable = True
                        print(f"[RETRY {retry}] [OK] Tombol target sekarang CLICKABLE!")
                        break

            # Jika tetap tidak clickable dan pagination menunjukkan ujung data
            if not clickable:
                print(f"[INFO] Tombol '{target_page_label}' / 'Next' nonaktif. Mencapai halaman terakhir.")
                break

            # 5. Ketuk tombol halaman target ('2', '3', dst) atau fallback 'Next'
            print(f"[PAGE {page_count}] Mengetuk tombol '{target_page_label}' (atau 'Next') untuk ke Halaman {page_count + 1}...")
            clicked = False
            for label in [target_page_label, "Next"]:
                if clicked:
                    break
                for selector in [d(text=label), d(description=label)]:
                    if selector.exists():
                        try:
                            selector.click()
                            clicked = True
                            break
                        except Exception:
                            pass
                if not clicked:
                    try:
                        xpath_el = d.xpath(f"//*[@text='{label}' or @content-desc='{label}']")
                        if xpath_el.exists:
                            xpath_el.click()
                            clicked = True
                    except Exception:
                        pass

            # Beri jeda transisi WebView memuat data halaman baru
            time.sleep(0.7)
            page_count += 1

            # 6. Scroll cepat kembali ke bagian atas tabel untuk halaman baru
            print(f"[SWIPE UP] Scroll cepat kembali ke atas halaman {page_count}...")
            for _ in range(3):
                self.device.swipe_up(ratio_start=0.20, ratio_end=0.75, duration=0.06)
                time.sleep(0.05)

            time.sleep(0.3)

        # 7. Sinkronisasi akhir seluruh data ke file Excel
        self.data_manager.sync_to_excel(self.output_txt, self.output_excel, CUSTOM_COLUMNS)


# ==============================================================================
# 5. BACKWARD-COMPATIBLE FUNCTION ALIASES & ENTRY POINT
# ==============================================================================
_global_device = EmulatorDevice()
d = None  # Global reference backward compatibility


def hubungkan_emulator():
    """Fungsi pembungkus backward compatibility untuk koneksi emulator."""
    global d
    success = _global_device.connect()
    if success:
        d = _global_device.d
    return success


def swipe_aman(fx, fy, tx, ty, duration=0.1):
    """Fungsi pembungkus backward compatibility untuk swipe aman."""
    _global_device.swipe_aman(fx, fy, tx, ty, duration=duration)


def cek_dan_tangani_kebablasan():
    """Fungsi pembungkus backward compatibility."""
    bot = FasihRejectBot(_global_device)
    return bot.cek_dan_tangani_kebablasan()


def ketuk_fab_dan_filter():
    """Fungsi pembungkus backward compatibility."""
    bot = FasihRejectBot(_global_device)
    return bot.ketuk_fab_dan_filter()


def terapkan_filter_reject():
    """Fungsi pembungkus backward compatibility."""
    bot = FasihRejectBot(_global_device)
    return bot.terapkan_filter_reject()


def ekstrak_data_tabel():
    """Fungsi pembungkus backward compatibility untuk ekstraksi baris tabel."""
    xml_data = _global_device.dump_hierarchy()
    return TableHierarchyParser.extract_table_data(xml_data)


def cek_apakah_halaman_terakhir():
    """Fungsi pembungkus backward compatibility untuk pengecekan halaman terakhir."""
    xml_data = _global_device.dump_hierarchy()
    _, _, _, is_last = TableHierarchyParser.extract_pagination(xml_data)
    return is_last


def salin_reject_ke_excel(file_txt="reject.txt", file_excel=DEFAULT_EXCEL_FILE):
    """Fungsi pembungkus backward compatibility untuk sinkronisasi Excel."""
    return DataManager.sync_to_excel(file_txt, file_excel, CUSTOM_COLUMNS)


def reset_posisi_entri_tabel():
    """Fungsi pembungkus backward compatibility untuk reset tabel."""
    bot = FasihRejectBot(_global_device)
    return bot.reset_posisi_entri_tabel()


def proses_ekstraksi_dan_swipe():
    """Fungsi pembungkus backward compatibility untuk proses scraping."""
    bot = FasihRejectBot(_global_device)
    bot.proses_ekstraksi_dan_swipe()


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

    print("==================================================")
    print("       BOT EMULATOR REJECT - FASIH SCRAPPER       ")
    print("==================================================")

    # 1. Hubungkan ke emulator
    if not hubungkan_emulator():
        return

    bot = FasihRejectBot(_global_device)

    # Cek apakah dialog filter sudah terbuka di layar saat ini
    d_inst = _global_device.d
    is_filter_open = (
        d_inst(text="Filter Assignment By Status").exists() or
        d_inst(resourceId="id.go.bpsfasih:id/lButton_bottomDialogFilterAssignment").exists() or
        d_inst(resourceId="id.go.bpsfasih:id/reject_cb_bottomSheetFilterAssignment").exists()
    )

    if is_filter_open:
        print("[INFO] Dialog 'Filter Assignment By Status' sudah terbuka di layar. Langsung ke Langkah 3...")
    else:
        # Cek jika kebablasan ke halaman 'Periode' atau 'Daftar Wilayah'
        if not bot.cek_dan_tangani_kebablasan():
            print("[HALT] Proses terhenti di Langkah 1.")
            return

        # Ketuk FAB & Filter By Status
        if not bot.ketuk_fab_dan_filter():
            print("[HALT] Proses terhenti di Langkah 2.")
            return

    # 2. Checklist filter status sesuai FILTER_STATUS_TARGET lalu ketuk 'TERAPKAN'
    if not bot.terapkan_filter_reject():
        print("[HALT] Proses terhenti di Langkah 3.")
        return

    # 3. Reset posisi entri tabel ke Halaman 1 & scroll ke paling atas
    bot.reset_posisi_entri_tabel()

    # 4. Ekstraksi data tabel per halaman, batch save TXT, navigasi Next, dan ekspor Excel
    bot.proses_ekstraksi_dan_swipe()

    print("\n==================================================")
    print("      PROSES BOT REJECT SELESAI DENGAN SUKSES     ")
    print("==================================================")


if __name__ == "__main__":
    main()
