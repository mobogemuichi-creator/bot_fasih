import re
import time
import xml.etree.ElementTree as ET
from konfigurasi import (
    SLEEP_SHORT,
    SLEEP_MEDIUM,
    SLEEP_LONG_REJECT as SLEEP_LONG,
)
from .base_page import BasePage


class KuesionerPage(BasePage):
    """
    Mengelola pengisian dan validasi seluruh blok kuesioner Fasih dengan Smart Wait:
    - BLOK I   : Cek ID Pelanggan, ambil data alamat, RadioButton '1. Berhasil didata', validasi awal & perbaikan galat koordinat/foto.
    - BLOK II  : 201. Nama penghuni, 202. NIK penghuni & Cek NIK, shortcut submit GALAT 0, 203. No Telp, RadioButton '1. Milik sendiri'.
    - BLOK III : 301. Alamat saat ini & kontrol increment (+).
    - BLOK IV  : 401. Catatan & eksekusi submit final.
    """

    def __init__(self, driver, galat_fixer, submit_service, excel_manager):
        super().__init__(driver)
        self.galat_fixer = galat_fixer
        self.submit_service = submit_service
        self.excel_manager = excel_manager

    def ketuk_sidebar_toggle(self):
        """Mengetuk ikon burger / sidebar toggle di pojok kiri atas (60, 160) dengan smart wait menu."""
        print("[SIDEBAR] Mengetuk tombol toggle sidebar...")
        try:
            self.d.click(60, 160)
        except Exception:
            self.ketuk("menu", exact=False)
        self.smart_wait(lambda: self.check_exists(self.d(textContains="BLOK")), timeout=0.8, poll_interval=0.04)

    def pilih_blok(self, nama_blok: str, max_attempts: int = 3) -> bool:
        """Universal switcher ke BLOK I, BLOK II, BLOK III, atau BLOK IV dengan penanganan dialog Pengaturan."""
        if not nama_blok:
            return False

        target = str(nama_blok).strip().upper()
        if target in ["1", "I"]:
            target = "BLOK I"
        elif target in ["2", "II"]:
            target = "BLOK II"
        elif target in ["3", "III"]:
            target = "BLOK III"
        elif target in ["4", "IV"]:
            target = "BLOK IV"
        elif not target.startswith("BLOK"):
            target = f"BLOK {target}"

        coords_map = {
            "BLOK I": (350, 450),
            "BLOK II": (350, 681),
            "BLOK III": (350, 912),
            "BLOK IV": (350, 1119),
        }

        for attempt in range(1, max_attempts + 1):
            print(f"[BLOK] Beralih ke '{target}' (Percobaan {attempt}/{max_attempts})...")
            if self.check_dan_tutup_pengaturan():
                time.sleep(0.2)

            clicked = False
            try:
                if self.check_exists(self.d(text=target)):
                    self.d(text=target).click()
                    clicked = True
                elif self.check_exists(self.d(textContains=target)):
                    self.d(textContains=target).click()
                    clicked = True
                elif self.check_exists(self.d(descriptionContains=target)):
                    self.d(descriptionContains=target).click()
                    clicked = True
            except Exception:
                pass

            if not clicked:
                try:
                    xp = f"//*[contains(@text, '{target}') or contains(@content-desc, '{target}')]"
                    if self.check_exists(self.d.xpath(xp)):
                        self.d.xpath(xp).click()
                        clicked = True
                except Exception:
                    pass

            if not clicked and target in coords_map:
                cx, cy = coords_map[target]
                self.d.click(cx, cy)
                clicked = True

            # Smart wait transisi ke blok target
            self.smart_wait_element(self.d(textContains=target), timeout=1.0, poll_interval=0.04)

            if self.check_dan_tutup_pengaturan():
                time.sleep(0.2)
                continue

            return True

        return False

    def ketuk_kirim_toolbar(self) -> bool:
        """Mengetuk tombol Kirim pada toolbar atas untuk memicu modal ringkasan validasi."""
        kirim_bounds = None
        try:
            btn_kirim = self.d(className="android.widget.Button", text="Kirim")
            if not btn_kirim.exists():
                btn_kirim = self.d(text="Kirim", clickable=True)
            if not btn_kirim.exists():
                btn_kirim = self.d(text="Kirim")

            if self.check_exists(btn_kirim):
                info = btn_kirim.info
                b = info.get("bounds") if isinstance(info, dict) else None
                if b and isinstance(b, dict):
                    cx = (b.get("left", 0) + b.get("right", 0)) // 2
                    cy = (b.get("top", 0) + b.get("bottom", 0)) // 2
                    kirim_bounds = (cx, cy)
        except Exception:
            pass

        if kirim_bounds:
            self.d.click(kirim_bounds[0], kirim_bounds[1])
            time.sleep(0.05)
            return True
        else:
            return self.ketuk("Kirim", sleep_after=0.05)

    # =========================================================================
    # BLOK I HELPERS
    # =========================================================================

    def normalisasi_nama_desa(self, val: str) -> str:
        """Normalisasi nama Desa/Kelurahan: PADANG SAMBIAN -> PADANGSAMBIAN, KELOD -> KLOD."""
        if not val:
            return val
        val = re.sub(r'PADANG\s+SAMBIAN', 'PADANGSAMBIAN', val, flags=re.IGNORECASE)
        val = re.sub(r'KELOD', 'KLOD', val, flags=re.IGNORECASE)
        return val

    def cek_idpel_tidak_ditemukan(self) -> tuple:
        """
        Memeriksa apakah muncul pesan / dialog 'tidak ditemukan' setelah Cek ID Pelanggan.
        Returns: tuple (is_not_found: bool, pesan: str)
        """
        not_found_keywords = [
            "id pelanggan tidak ditemukan",
            "ID pelanggan tidak ditemukan",
            "Id pelanggan tidak ditemukan",
            "ID Pelanggan tidak ditemukan",
            "Data ID Pelanggan tidak ditemukan",
            "Pelanggan tidak ditemukan",
            "tidak ditemukan",
            "TIDAK DITEMUKAN",
            "Tidak Ditemukan",
        ]
        try:
            for kw in not_found_keywords:
                el = self.d(textContains=kw)
                if self.check_exists(el):
                    txt = el.info.get('text', kw)
                    return True, txt

                el_desc = self.d(descriptionContains=kw)
                if self.check_exists(el_desc):
                    txt = el_desc.info.get('contentDescription', kw)
                    return True, txt

            xp = "//*[contains(translate(@text, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tidak ditemukan') or contains(translate(@content-desc, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tidak ditemukan')]"
            if self.check_exists(self.d.xpath(xp)):
                nodes = self.d.xpath(xp).all()
                if nodes:
                    txt = nodes[0].text if hasattr(nodes[0], 'text') and nodes[0].text else "tidak ditemukan"
                    return True, txt
        except Exception as e:
            print(f"[BLOK I] Exception scan idpel tidak ditemukan: {e}")

        return False, ""

    def cek_dan_ketuk_cek_id_pelanggan(self, row: int = 0, idpel: str = "", max_swipes: int = 10) -> str:
        """Scroll dinamis mencari tombol 'Cek ID Pelanggan' pada BLOK I, mengetuknya, dan memeriksa pesan 'tidak ditemukan'."""
        print("[BLOK I] Men-scroll secara dinamis ke tombol 'Cek ID Pelanggan'...")
        try:
            self.d(scrollable=True).scroll.to(text="Cek ID Pelanggan")
            time.sleep(0.1)
        except Exception:
            pass

        screen_w, screen_h = self.driver.screen_width, self.driver.screen_height
        btn_cek = None

        for swipe_cek in range(1, max_swipes + 1):
            btn_cek = self.d(text="Cek ID Pelanggan")
            if not btn_cek.exists():
                btn_cek = self.d(textContains="Cek ID Pelanggan")
            if not btn_cek.exists():
                btn_cek = self.d(descriptionContains="Cek ID Pelanggan")

            if btn_cek.exists():
                b = btn_cek.info.get("bounds", {})
                top = b.get("top", 0)
                bottom = b.get("bottom", 0)
                if 150 <= top and bottom <= (screen_h - 120):
                    print(f"[BLOK I] Tombol 'Cek ID Pelanggan' berada di posisi aman layar.")
                    break
                elif top > (screen_h - 120):
                    self.driver.swipe_aman(screen_w // 2, int(screen_h * 0.7), screen_w // 2, int(screen_h * 0.4), duration=0.15)
                    time.sleep(0.1)
                elif bottom < 150:
                    self.driver.swipe_aman(screen_w // 2, int(screen_h * 0.3), screen_w // 2, int(screen_h * 0.6), duration=0.15)
                    time.sleep(0.1)
            else:
                self.driver.swipe_aman(screen_w // 2, int(screen_h * 0.7), screen_w // 2, int(screen_h * 0.4), duration=0.15)
                time.sleep(0.1)

        sukses_ketuk_cek = False
        if btn_cek and btn_cek.exists():
            b = btn_cek.info.get("bounds", {})
            cx = (b.get("left", 0) + b.get("right", 0)) // 2
            cy = (b.get("top", 0) + b.get("bottom", 0)) // 2
            if 0 < cx < screen_w and 0 < cy < screen_h:
                print(f"[BLOK I] Mengetuk tombol 'Cek ID Pelanggan' pada titik tengah ({cx}, {cy})...")
                self.d.click(cx, cy)
                sukses_ketuk_cek = True
            else:
                btn_cek.click()
                sukses_ketuk_cek = True

        if not sukses_ketuk_cek:
            print("[BLOK I] Mengetuk tombol 'Cek ID Pelanggan' via ketuk()...")
            self.ketuk("Cek ID Pelanggan")

        time.sleep(0.2)
        progress_el = self.d(resourceId="id.go.bpsfasih:id/card_progress")
        if not progress_el.exists():
            progress_el = self.d(className="android.widget.ProgressBar")

        if not progress_el.exists():
            print("[BLOK I] [RETRY] Loading belum terdeteksi, mencoba mengetuk ulang 'Cek ID Pelanggan'...")
            self.ketuk("Cek ID Pelanggan")

        print("[BLOK I] [LOADING] Menunggu loading 'Cek ID Pelanggan' selesai...")
        self.tunggu_loading(timeout=30)
        time.sleep(0.05)

        # Pengecekan respon text 'tidak ditemukan' setelah Cek ID Pelanggan
        is_not_found, msg_not_found = self.cek_idpel_tidak_ditemukan()
        if is_not_found:
            print(f"\n[BLOK I] [TIDAK DITEMUKAN] Terdeteksi '{msg_not_found}' setelah Cek ID Pelanggan (IDPEL: {idpel}).")
            for btn_dismiss_name in ["OK", "Ok", "TUTUP", "Tutup", "Dismiss", "Batal"]:
                if self.check_exists(self.d(text=btn_dismiss_name)):
                    try:
                        self.d(text=btn_dismiss_name).click()
                        time.sleep(0.2)
                        break
                    except Exception:
                        pass
            if row > 0:
                print(f"[SKIP IDPEL] Menyimpan status Excel 'idpel tidak ditemukan' pada baris {row} & lanjut ke IDPEL berikutnya...")
                self.excel_manager.simpan_status(row, "idpel tidak ditemukan")
            self.kembali_ke_daftar_assignment()
            return "idpel_tidak_ditemukan"

        return "sukses"

    def ambil_data_alamat(self, file_output: str = "temp_alamat.txt", idpel: str = "") -> dict:
        """Mengambil data alamat dari screen BLOK I dan menyimpannya ke temp_alamat.txt."""
        print("\n[BLOK I] Men-scroll ke bawah secara dinamis hingga data alamat terlihat...")
        info_alamat = {
            "Provinsi": "",
            "Kabupaten": "",
            "Kecamatan": "",
            "Desa/Kelurahan": "",
            "Alamat": ""
        }
        max_swipes = 20
        for swipe_idx in range(1, max_swipes + 1):
            try:
                xml = self.d.dump_hierarchy()
                xml_lower = xml.lower()
            except Exception:
                xml = ""
                xml_lower = ""

            if ("terhubung ke server" in xml_lower or "tidak dapat terhubung" in xml_lower or
                (not xml and self.check_exists(self.d(textContains="TIDAK DAPAT TERHUBUNG")))):
                print(f"[BLOK I] [WARNING] Terdeteksi 'TIDAK DAPAT TERHUBUNG KE SERVER' pada swipe ke-{swipe_idx}!")
                info_alamat["server_error"] = True
                return info_alamat

            if "tidak match" in xml_lower or (not xml and self.check_exists(self.d(textContains="tidak match"))):
                print(f"[BLOK I] [TIDAK MATCH] Terdeteksi 'tidak match'! Mengetuk 'Cek ID Pelanggan'...")
                self.ketuk("Cek ID Pelanggan")
                self.tunggu_loading(timeout=30)
                time.sleep(0.05)

                is_nf, msg_nf = self.cek_idpel_tidak_ditemukan()
                if is_nf:
                    print(f"\n[BLOK I] [TIDAK DITEMUKAN] Terdeteksi '{msg_nf}' saat Cek ID Pelanggan di ambil_data_alamat (IDPEL: {idpel}).")
                    info_alamat["idpel_tidak_ditemukan"] = True
                    return info_alamat

            if "103." in xml or "nama pada id pelanggan" in xml_lower or (not xml and (self.d(textContains="103.").exists() or self.d(textContains="Nama pada ID Pelanggan").exists())):
                print(f"[BLOK I] Teks alamat ditemukan di layar (pemeriksaan ke-{swipe_idx}).")
                break

            try:
                self.driver.swipe_aman(540, 700, 540, 400, duration=0.1)
                time.sleep(0.04)
            except Exception:
                break

        def cek_label_provinsi():
            try:
                xml_check = self.d.dump_hierarchy().lower()
                return "provinsi" in xml_check
            except Exception:
                patterns = ["a. Provinsi", "a.  Provinsi", "a.Provinsi", "Provinsi", "provinsi"]
                for p in patterns:
                    if self.d(textContains=p).exists():
                        return True
                return False

        if not cek_label_provinsi():
            for swipe_up_idx in range(1, 11):
                if cek_label_provinsi():
                    break
                try:
                    self.driver.swipe_aman(540, 400, 540, 600, duration=0.1)
                    time.sleep(0.05)
                except Exception:
                    break

        mapping_keys = [
            ("Provinsi", ["a. Provinsi", "Provinsi"]),
            ("Kabupaten", ["b. Kabupaten/Kota", "Kabupaten/Kota", "Kabupaten"]),
            ("Kecamatan", ["c. Kecamatan", "Kecamatan"]),
            ("Desa/Kelurahan", ["d. Desa/Kelurahan", "Desa/Kelurahan", "Desa"]),
            ("Alamat", ["e. Alamat", "Alamat"])
        ]

        for key, patterns in mapping_keys:
            for p in patterns:
                el = self.d(textContains=p)
                if el.exists():
                    txt = el.info.get('text', '').strip()
                    if ":" in txt and len(txt) > len(p) + 2:
                        info_alamat[key] = f"{key}: {txt.split(':', 1)[1].strip()}"
                        break
                    else:
                        sibling = el.sibling(className="android.widget.EditText")
                        if not sibling.exists():
                            sibling = el.sibling(className="android.widget.TextView")
                        if not sibling.exists():
                            sibling = el.down(className="android.widget.EditText")
                        if not sibling.exists():
                            sibling = el.down(className="android.widget.TextView")

                        if sibling.exists():
                            val = sibling.info.get('text', '').strip()
                            if val and not any(val.startswith(x) for x in ["a. ", "b. ", "c. ", "d. ", "e. "]):
                                info_alamat[key] = f"{key}: {val}"
                                break

            if not info_alamat[key]:
                info_alamat[key] = f"{key}: (tidak ditemukan)"

        if info_alamat.get("Desa/Kelurahan"):
            info_alamat["Desa/Kelurahan"] = self.normalisasi_nama_desa(info_alamat["Desa/Kelurahan"])

        try:
            with open(file_output, "w", encoding="utf-8") as f:
                for k, v in info_alamat.items():
                    f.write(f"{v}\n")
        except Exception:
            pass

        return info_alamat

    def cek_radio_button_tercentang(self, option_text: str, exact: bool = False) -> bool:
        """Memeriksa apakah RadioButton bersaudara memiliki child (berstatus tercentang)."""
        if not option_text:
            return False
        target = str(option_text).strip()
        try:
            xml_data = self.d.dump_hierarchy()
            root = ET.fromstring(xml_data)
            for parent in root.iter():
                children = list(parent)
                if len(children) < 2:
                    continue
                has_target = any(
                    (target == (c.attrib.get("text") or "").strip()) if exact else (target in (c.attrib.get("text") or ""))
                    for c in children
                )
                if has_target:
                    for c in children:
                        if c.attrib.get("NAF") == "true" and len(list(c)) > 0:
                            return True
        except Exception:
            pass
        return False

    def pilih_radio_button_berhasil_didata(self, idpel: str, max_scan_attempts: int = 20) -> bool:
        """Scan dan scroll dinamis mencari RadioButton '1. Berhasil didata'."""
        print("[RADIO] Memulai scan & scroll dinamis untuk mencari 'Berhasil Didata'...")
        scroll_direction = "down"

        for scan_attempt in range(1, max_scan_attempts + 1):
            if self.cek_radio_button_tercentang("Berhasil didata", exact=False):
                print("[RADIO] [SUKSES] RadioButton 'Berhasil Didata' sudah tercentang!")
                return True

            target_bounds = None
            for pattern in ["Berhasil Didata", "Berhasil didata", "1. Berhasil didata"]:
                try:
                    elements = self.d.xpath(f"//*[contains(@text, '{pattern}') or contains(@content-desc, '{pattern}')]").all()
                    for el in elements:
                        attrib = el.attrib if hasattr(el, 'attrib') else {}
                        bounds_str = attrib.get('bounds', '')
                        pts = [int(x) for x in re.findall(r'\d+', bounds_str)]
                        if len(pts) == 4:
                            x1, y1, x2, y2 = pts
                            if (x2 - x1) >= 50:
                                target_bounds = {"left": x1, "top": y1, "right": x2, "bottom": y2}
                                break
                    if target_bounds:
                        break
                except Exception:
                    pass

            if target_bounds:
                b = target_bounds
                cx = (b["left"] + b["right"]) // 2
                cy = (b["top"] + b["bottom"]) // 2
                self.d.click(cx, cy)
                time.sleep(0.3)

                if self.cek_radio_button_tercentang("Berhasil didata", exact=False):
                    print("[VERIFIKASI] [SUKSES] Status RadioButton 'Berhasil Didata' kini TERCENTANG.")
                    return True
                else:
                    radio_x = max(20, b["left"] - 48)
                    self.d.click(radio_x, cy)
                    time.sleep(0.3)
                    if self.cek_radio_button_tercentang("Berhasil didata", exact=False):
                        print("[VERIFIKASI] [SUKSES] Status RadioButton 'Berhasil Didata' kini TERCENTANG via offset.")
                        return True

            has_105 = (
                self.d(textContains="105. Koordinat").exists() or
                self.d(textContains="Koordinat lokasi meteran").exists() or
                self.d(descriptionContains="105. Koordinat").exists()
            )

            if scroll_direction == "down" and has_105:
                scroll_direction = "up"
            elif scroll_direction == "up" and not has_105:
                scroll_direction = "down"

            if scroll_direction == "down":
                self.driver.loop_swipe_statis(delta_y=-300, loop=1)
            else:
                self.driver.loop_swipe_statis(delta_y=250, loop=1)

            time.sleep(0.05)

        return False

    def proses_blok_i_dan_validasi_awal(self, idpel: str, row: int, row_attempt: int) -> str:
        """
        Memproses BLOK I dan Validasi Awal dengan Smart Wait:
        1. Cek galat 'Nomor Meter' / 'ID pelanggan PLN' -> ketuk 'Cek ID Pelanggan' jika perlu.
        2. Ambil data alamat.
        3. Pastikan RadioButton '1. Berhasil didata' tercentang.
        4. Ketuk Kirim -> Cek GALAT 0 & KOSONG 0 (jika bersih langsung submit).
        5. Perbaiki Galat Error 105 (GPS) & 106 (Foto) jika ada.
        """
        # 1. Cek GALAT untuk Nomor Meter / ID pelanggan PLN
        print("[VALIDASI AWAL] Mengetuk tombol 'Kirim' untuk mengecek GALAT...")
        self.ketuk("Kirim", sleep_after=0.05)
        self.smart_wait_element(self.d(textContains="GALAT"), timeout=1.5, poll_interval=0.04)

        print("[VALIDASI AWAL] Mengetuk 'GALAT' pada modal validasi...")
        self.ketuk("GALAT", exact=False, sleep_after=0.05)

        try:
            xml_galat = self.d.dump_hierarchy().lower()
        except Exception:
            xml_galat = ""

        is_nomor_meter_galat = (
            "nomor meter" in xml_galat or
            "id pelanggan pln" in xml_galat or
            self.check_exists(self.d(textContains="Nomor Meter")) or
            self.check_exists(self.d(textContains="ID pelanggan PLN"))
        )

        if is_nomor_meter_galat:
            print(f"[GALAT CHECK] [TRUE] Terdeteksi galat 'Nomor Meter' / 'ID pelanggan PLN' untuk IDPEL {idpel}!")
            self.tutup_modal_dismiss_aman()
            res_cek_idpel = self.cek_dan_ketuk_cek_id_pelanggan(row=row, idpel=idpel)
            if res_cek_idpel == "idpel_tidak_ditemukan":
                return "skip"
        else:
            print("[GALAT CHECK] [FALSE] Tidak terdeteksi kata 'Nomor Meter' / 'ID pelanggan PLN'. Menutup modal...")
            self.tutup_modal_dismiss_aman()

        # 2. Ambil data alamat
        alamat_dict = self.ambil_data_alamat(file_output="temp_alamat.txt", idpel=idpel)
        time.sleep(0.05)

        if alamat_dict and alamat_dict.get("idpel_tidak_ditemukan"):
            print(f"[BLOK I] [SKIP] IDPEL {idpel} tidak ditemukan setelah Cek ID Pelanggan. Menyimpan status 'idpel tidak ditemukan' & lanjut ke IDPEL berikutnya...")
            for btn_dismiss_name in ["OK", "Ok", "TUTUP", "Tutup", "Dismiss", "Batal"]:
                if self.check_exists(self.d(text=btn_dismiss_name)):
                    try:
                        self.d(text=btn_dismiss_name).click()
                        time.sleep(0.2)
                        break
                    except Exception:
                        pass
            self.excel_manager.simpan_status(row, "idpel tidak ditemukan")
            self.kembali_ke_daftar_assignment()
            return "skip"

        if (alamat_dict and alamat_dict.get("server_error")) or \
           self.check_exists(self.d(textContains="TIDAK DAPAT TERHUBUNG KE SERVER")) or \
           self.check_exists(self.d(textContains="TIDAK DAPAT TERHUBUNG")):
            print(f"[BLOK I] [SKIP] Terdeteksi 'TIDAK DAPAT TERHUBUNG KE SERVER' untuk IDPEL {idpel}.")
            self.excel_manager.simpan_status(row, "Data Belum Cek IDPEL/NOMETER")
            self.kembali_ke_daftar_assignment()
            return "skip"

        # 3. RadioButton '1. Berhasil didata'
        sukses_radio = self.pilih_radio_button_berhasil_didata(idpel=idpel)
        if not sukses_radio:
            print(f"[RADIO CHECK] [SKIP] RadioButton 'Berhasil Didata' gagal tercentang untuk IDPEL {idpel}.")
            self.excel_manager.simpan_status(row, "Error : RadioButton 1. Berhasil didata tidak tercentang")
            self.kembali_ke_daftar_assignment()
            return "skip"

        # 4. Ketuk tombol 'Kirim' toolbar untuk membuka modal validasi secara reaktif
        max_kirim_attempts = 5
        for kirim_attempt in range(1, max_kirim_attempts + 1):
            self.ketuk_kirim_toolbar()
            # Smart wait hingga label Mulai Wawancara tertutup modal
            modal_opened = self.smart_wait(
                lambda: not (
                    self.check_exists(self.d(textContains="Mulai Wawancara")) or
                    self.check_exists(self.d(descriptionContains="Mulai Wawancara"))
                ),
                timeout=1.0,
                poll_interval=0.04,
            )
            if modal_opened:
                break
            time.sleep(0.1)

        # 5. Cek apakah langsung GALAT 0 & KOSONG 0 dari awal
        is_galat_0 = bool(
            self.check_exists(self.d(textContains="GALAT 0")) or
            self.check_exists(self.d(descriptionContains="GALAT 0")) or
            self.check_exists(self.d.xpath("//*[contains(@text, 'GALAT 0') or contains(@content-desc, 'GALAT 0')]"))
        )
        is_kosong_0 = bool(
            self.check_exists(self.d(textContains="KOSONG 0")) or
            self.check_exists(self.d(descriptionContains="KOSONG 0")) or
            self.check_exists(self.d.xpath("//*[contains(@text, 'KOSONG 0') or contains(@content-desc, 'KOSONG 0')]"))
        )

        if is_galat_0 and is_kosong_0:
            print(f"[KIRIM CHECK] [SUKSES BERSIH] Terdeteksi 'GALAT 0' dan 'KOSONG 0' dari awal!")
            if not self.submit_service.pause_proses_galat_0(idpel=idpel, row=row, keterangan="Cek Form Awal"):
                return "stop"
            res_submit = self.submit_service.eksekusi_submit_dan_selesai(row, idpel, row_attempt)
            return "selesai" if res_submit != "retry" else "retry"

        # 6. Ketuk GALAT dan perbaiki Error 105 / 106 jika ada
        self.ketuk("GALAT", exact=False, sleep_after=0.05)
        self.smart_wait(lambda: self.check_exists(self.d(textContains="105")) or self.check_exists(self.d(textContains="106")) or self.check_exists(self.d(text="Dismiss")), timeout=1.0, poll_interval=0.04)

        is_galat_koordinat_foto = bool(
            self.check_exists(self.d(textContains="Koordinat lokasi meteran")) or
            self.check_exists(self.d(descriptionContains="Koordinat lokasi meteran")) or
            self.check_exists(self.d.xpath("//*[contains(@text, 'Koordinat lokasi meteran') or contains(@content-desc, 'Koordinat lokasi meteran')]")) or
            self.check_exists(self.d(textContains="Foto rumah tampak depan")) or
            self.check_exists(self.d(textContains="Foto ruimah tampak depan")) or
            self.check_exists(self.d(descriptionContains="Foto rumah tampak depan")) or
            self.check_exists(self.d.xpath("//*[contains(@text, 'tampak depan') or contains(@content-desc, 'tampak depan')]"))
        )

        if is_galat_koordinat_foto:
            print(f"[GALAT CHECK] Terdeteksi galat koordinat/foto untuk IDPEL {idpel}. Memperbaiki otomatis...")
            fix_berhasil = self.galat_fixer.perbaiki_galat_koordinat_foto(skip_ketuk_galat=True)
            if fix_berhasil:
                print("[GALAT FIX] Perbaikan berhasil pada validasi awal. Melanjutkan ke BLOK II...")
                return "lanjut"
            else:
                print(f"[GALAT CHECK] [SKIP] Perbaikan otomatis gagal untuk IDPEL {idpel}.")
                self.excel_manager.simpan_status(row, "koordinat & foto tidak ada")
                self.kembali_ke_daftar_assignment()
                return "skip"
        else:
            self.tutup_modal_dismiss_aman()
            return "lanjut"

    # =========================================================================
    # BLOK II HELPERS & PROCESSING
    # =========================================================================

    def verifikasi_nama_penghuni(self, nama: str):
        """Memeriksa dan mengisi field '201. Nama penghuni'."""
        try:
            label = self.d(textContains="201. Nama penghuni")
            if label.exists():
                ed = label.down(className="android.widget.EditText")
                if ed.exists():
                    val = ed.info.get("text", "").strip()
                    if val and val != "Wajib diisi":
                        print(f"[SCAN NAMA] [SUKSES] Field '201. Nama penghuni' sudah terisi di UI ('{val}'). Melewati pengisian.")
                        return
        except Exception:
            pass

        if not nama:
            print("[SCAN NAMA] Kolom NAMA di Excel kosong, melewati pengisian.")
            return

        for attempt in range(1, 6):
            print(f"[INPUT NAMA] Memasukkan '201. Nama penghuni': '{nama}' (Percobaan {attempt}/5)...")
            self.input_textbox("201. Nama penghuni", nama, exact=False, sleep_after=0.05)
            time.sleep(0.2)

            try:
                label = self.d(textContains="201. Nama penghuni")
                if label.exists():
                    ed = label.down(className="android.widget.EditText")
                    if ed.exists():
                        val = ed.info.get("text", "").strip()
                        if val and val != "Wajib diisi":
                            print(f"[SCAN NAMA] [SUKSES] Field Nama terisi dengan sukses pada percobaan ke-{attempt}.")
                            break
            except Exception:
                pass

    def baca_nilai_field_nik(self) -> str:
        """Membaca isi nilai field NIK penghuni saat ini di layar."""
        try:
            label = self.d(textContains="202. NIK penghuni")
            if not label.exists():
                label = self.d(textContains="202.")
            if label.exists():
                ed = label.down(className="android.widget.EditText")
                if ed.exists():
                    return ed.info.get("text", "").strip()
        except Exception:
            pass
        return ""

    def cari_dan_scroll_ke_tombol_cek_nik(self, max_swipes: int = 10):
        """Scroll layar secara bertahap hingga tombol 'Cek NIK' terlihat jelas di viewport aman."""
        screen_w, screen_h = self.driver.screen_width, self.driver.screen_height
        for _ in range(max_swipes):
            btn = self.d(text="Cek NIK")
            if not btn.exists():
                btn = self.d(textContains="Cek NIK")
            if btn.exists():
                b = btn.info.get("bounds", {})
                top = b.get("top", 0)
                bottom = b.get("bottom", 0)
                if 150 <= top and bottom <= (screen_h - 120):
                    break
                elif top > (screen_h - 120):
                    self.driver.swipe_aman(screen_w // 2, int(screen_h * 0.7), screen_w // 2, int(screen_h * 0.4), duration=0.15)
                elif bottom < 150:
                    self.driver.swipe_aman(screen_w // 2, int(screen_h * 0.3), screen_w // 2, int(screen_h * 0.6), duration=0.15)
            else:
                self.driver.swipe_aman(screen_w // 2, int(screen_h * 0.7), screen_w // 2, int(screen_h * 0.4), duration=0.15)
            time.sleep(0.1)

    def cek_nik_tidak_ditemukan(self) -> tuple:
        """
        Memeriksa apakah ada pesan 'NIK tidak ditemukan' atau status 'TIDAK DITEMUKAN'.
        Returns: tuple (is_not_found: bool, pesan: str)
        """
        not_found_keywords = [
            "NIK tidak ditemukan",
            "nik tidak ditemukan",
            "NIK TIDAK DITEMUKAN",
            "Nik tidak ditemukan",
            "Data NIK tidak ditemukan",
            "TIDAK DITEMUKAN",
            "tidak ditemukan"
        ]
        try:
            for kw in not_found_keywords:
                el = self.d(textContains=kw)
                if self.check_exists(el):
                    txt = el.info.get('text', kw)
                    return True, txt

                el_desc = self.d(descriptionContains=kw)
                if self.check_exists(el_desc):
                    txt = el_desc.info.get('contentDescription', kw)
                    return True, txt

            xp = "//*[contains(@text, 'tidak ditemukan') or contains(@content-desc, 'tidak ditemukan') or contains(@text, 'TIDAK DITEMUKAN') or contains(@content-desc, 'TIDAK DITEMUKAN')]"
            if self.check_exists(self.d.xpath(xp)):
                nodes = self.d.xpath(xp).all()
                if nodes:
                    txt = nodes[0].text if hasattr(nodes[0], 'text') and nodes[0].text else "NIK tidak ditemukan"
                    return True, txt
        except Exception:
            pass

        return False, ""

    def cek_nik_tidak_valid(self) -> tuple:
        """
        Memeriksa apakah ada pesan error 'NIK tidak valid' di layar.
        Returns: tuple (is_invalid: bool, pesan_error: str)
        """
        invalid_keywords = ["NIK tidak valid", "periksa digit", "digit pertama", "digit ketiga", "digit ke-7"]
        try:
            for kw in invalid_keywords:
                if self.check_exists(self.d(textContains=kw)) or self.check_exists(self.d(descriptionContains=kw)):
                    txt = self.d(textContains=kw).info.get('text', kw) if self.check_exists(self.d(textContains=kw)) else kw
                    return True, txt

            xp = "//*[contains(@text, 'NIK tidak valid') or contains(@content-desc, 'NIK tidak valid') or contains(@text, 'periksa digit')]"
            if self.check_exists(self.d.xpath(xp)):
                nodes = self.d.xpath(xp).all()
                if nodes:
                    txt = nodes[0].text if hasattr(nodes[0], 'text') and nodes[0].text else "NIK tidak valid"
                    return True, txt
        except Exception:
            pass

        return False, ""

    def pilih_radio_button_milik_sendiri(self):
        """Memilih RadioButton '1. Milik sendiri' jika belum tercentang."""
        if self.cek_radio_button_tercentang("1. Milik sendiri"):
            print("[RADIO CHECK] RadioButton '1. Milik sendiri' sudah tercentang.")
            return

        print("[RADIO CHECK] RadioButton '1. Milik sendiri' BELUM tercentang. Memulai pengetukan...")
        for r_attempt in range(1, 6):
            target_bounds = None
            for pattern in ["1. Milik sendiri", "1.  Milik sendiri", "Milik sendiri"]:
                try:
                    elements = self.d.xpath(f"//*[contains(@text, '{pattern}') or contains(@content-desc, '{pattern}')]").all()
                    for el in elements:
                        attrib = el.attrib if hasattr(el, "attrib") else {}
                        bounds_str = attrib.get("bounds", "")
                        pts = [int(x) for x in re.findall(r"\d+", bounds_str)]
                        if len(pts) == 4:
                            x1, y1, x2, y2 = pts
                            if (x2 - x1) >= 50:
                                target_bounds = {"left": x1, "top": y1, "right": x2, "bottom": y2}
                                break
                    if target_bounds:
                        break
                except Exception:
                    pass

            if target_bounds:
                b = target_bounds
                cx = (b["left"] + b["right"]) // 2
                cy = (b["top"] + b["bottom"]) // 2
                self.d.click(cx, cy)
                time.sleep(0.3)

                if self.cek_radio_button_tercentang("1. Milik sendiri"):
                    print(f"[VERIFIKASI] [SUKSES] RadioButton '1. Milik sendiri' tercentang.")
                    return
                else:
                    radio_x = max(20, b["left"] - 48)
                    self.d.click(radio_x, cy)
                    time.sleep(0.3)
                    if self.cek_radio_button_tercentang("1. Milik sendiri"):
                        print(f"[VERIFIKASI] [SUKSES] RadioButton '1. Milik sendiri' tercentang via offset.")
                        return

            time.sleep(0.05)
            self.driver.loop_swipe_statis(delta_y=-100, loop=1)

    def proses_blok_ii(self, nama: str, nik: str, idpel: str, row: int, row_attempt: int) -> str:
        """
        Memproses BLOK II:
        1. Beralih ke Blok II & verifikasi halaman.
        2. Nama penghuni (201).
        3. NIK penghuni (202) & Cek NIK.
        4. Shortcut cek GALAT 0 langsung setelah Cek NIK (submit jika GALAT 0).
        5. No Telp & RadioButton '1. Milik sendiri'.
        """
        self.ketuk_sidebar_toggle()
        self.pilih_blok("II")

        is_blok_ii = bool(
            self.check_exists(self.d(textContains="BLOK II")) or
            self.check_exists(self.d(textContains="Blok II")) or
            self.check_exists(self.d(descriptionContains="BLOK II")) or
            self.check_exists(self.d(textContains="201. Nama penghuni"))
        )

        if not is_blok_ii:
            print(f"[BLOK II] [SKIP] Halaman tidak berada di 'Blok II' setelah pilih_blok('II') (IDPEL: {idpel}).")
            self.excel_manager.simpan_status(row, "Error : Gagal masuk Blok II")
            self.kembali_ke_daftar_assignment()
            return "skip"

        # 1. Isi & Verifikasi Nama
        self.verifikasi_nama_penghuni(nama)

        # 2. Input NIK & Cek NIK
        max_nik_attempts = 5
        nik_sukses = False

        for nik_attempt in range(1, max_nik_attempts + 1):
            nik_existing = self.baca_nilai_field_nik()
            perlu_isi = (not nik_existing) or (nik_existing == "9999999999999998") or (nik_existing != nik)

            if perlu_isi:
                print(f"[INPUT NIK] Memasukkan '202. NIK penghuni': '{nik}' (Percobaan {nik_attempt}/{max_nik_attempts})...")
                self.input_textbox("202. NIK penghuni", nik, exact=False, sleep_after=0.05)
                time.sleep(0.2)

            self.cari_dan_scroll_ke_tombol_cek_nik()
            print(f"[KLIK] Mengetuk 'Cek NIK' (Percobaan {nik_attempt}/{max_nik_attempts})...")
            self.ketuk("Cek NIK")
            self.tunggu_loading_cek_nik(timeout=30)

            # Cek NIK tidak ditemukan
            is_not_found, msg_not_found = self.cek_nik_tidak_ditemukan()
            if is_not_found:
                print(f"\n[SCAN NIK] [TIDAK DITEMUKAN] Terdeteksi '{msg_not_found}' untuk NIK '{nik}' (IDPEL: {idpel}).")
                for btn_dismiss_name in ["OK", "Ok", "TUTUP", "Tutup", "Dismiss", "Batal"]:
                    if self.check_exists(self.d(text=btn_dismiss_name)):
                        try:
                            self.d(text=btn_dismiss_name).click()
                            time.sleep(0.2)
                            break
                        except Exception:
                            pass
                self.excel_manager.simpan_status(row, "Error : NIK tidak ditemukan")
                self.kembali_ke_daftar_assignment()
                return "skip"

            # Cek NIK tidak valid
            is_invalid, msg_invalid = self.cek_nik_tidak_valid()
            if is_invalid:
                print(f"[SCAN NIK] [RETRY NIK] Terdeteksi '{msg_invalid}' (percobaan ke-{nik_attempt}/{max_nik_attempts}).")
                self.driver.loop_swipe_statis(delta_y=-400, loop=1)
                time.sleep(0.1)
                self.driver.loop_swipe_statis(delta_y=400, loop=1)
                time.sleep(0.05)
            else:
                print(f"[SCAN NIK] [SUKSES] NIK '{nik}' berhasil dicek.")
                nik_sukses = True
                break

        if not nik_sukses:
            print(f"[SCAN NIK] [GAGAL] NIK '{nik}' tetap invalid setelah {max_nik_attempts}x percobaan.")
            self.excel_manager.simpan_status(row, "Error : Nik tidak valid")
            self.kembali_ke_daftar_assignment()
            return "skip"

        # 3. PENGECEKAN GALAT 0 SETELAH CEK NIK (Shortcut Langsung Submit) dengan Smart Wait
        print(f"\n[CEK NIK -> VALIDASI] Memeriksa status 'GALAT 0' setelah Cek NIK untuk IDPEL {idpel}...")
        self.ketuk_kirim_toolbar()

        if self.smart_wait_element(self.d(text="YA"), timeout=1.0, poll_interval=0.04):
            self.ketuk("YA", sleep_after=0.05)

        is_galat_0 = False
        modal_terbuka = False

        # Smart wait deteksi modal dan status GALAT
        matched_idx, _ = self.smart_wait_any([
            lambda: (
                self.check_exists(self.d(textContains="GALAT 0 Perlu diperbaiki")) or
                self.check_exists(self.d(textContains="GALAT 0")) or
                self.check_exists(self.d(descriptionContains="GALAT 0")) or
                self.check_exists(self.d.xpath("//*[contains(@text, 'GALAT 0') or contains(@content-desc, 'GALAT 0')]"))
            ),
            lambda: (
                self.check_exists(self.d(textContains="GALAT")) or
                self.check_exists(self.d(text="Dismiss")) or
                self.check_exists(self.d(textContains="Perlu diperbaiki"))
            ),
        ], timeout=2.5, poll_interval=0.04)

        if matched_idx == 0:
            is_galat_0 = True
            modal_terbuka = True
        elif matched_idx == 1:
            modal_terbuka = True

        if is_galat_0:
            print(f"[CEK NIK -> SUBMIT] [SUKSES] Terdeteksi 'GALAT 0' setelah Cek NIK untuk IDPEL {idpel}! Langsung submit...")
            if not self.submit_service.pause_proses_galat_0(idpel=idpel, row=row, keterangan="Setelah Cek NIK (GALAT 0)"):
                return "stop"
            res_submit = self.submit_service.eksekusi_submit_dan_selesai(row, idpel, row_attempt)
            return "selesai" if res_submit != "retry" else "retry"
        else:
            print(f"[CEK NIK -> VALIDASI] Belum 'GALAT 0'. Menutup modal & melanjutkan pengisian form...")
            if modal_terbuka or self.check_exists(self.d(text="Dismiss")):
                self.tutup_modal_dismiss_aman()

        # 4. No Telp & RadioButton '1. Milik sendiri'
        self.driver.loop_swipe_statis(delta_y=-700, loop=3)
        self.input_textbox("203. Nomor telepon/HP penghuni", "-", exact=False, sleep_after=0.05)
        self.pilih_radio_button_milik_sendiri()

        return "lanjut"

    # =========================================================================
    # BLOK III & BLOK IV
    # =========================================================================

    def proses_blok_iii(self):
        """Memproses BLOK III: Alamat saat ini (301) dan tombol increment (+)."""
        self.ketuk_sidebar_toggle()
        self.pilih_blok("III")

        print("[BLOK III] Mengisi alamat saat ini (301)...")
        self.input_textbox("301. Alamat tempat tinggal", "-", exact=False, sleep_after=0.05)

        print("[BLOK III] Mengetuk tombol increment kontrol...")
        for inc_sel in [self.d(resourceId="id.go.bpsfasih:id/btn_increment"), self.d(text="+")]:
            if self.check_exists(inc_sel):
                try:
                    inc_sel.click()
                    break
                except Exception:
                    pass

    def proses_blok_iv(self, row: int, idpel: str, row_attempt: int) -> str:
        """
        Memproses BLOK IV: Catatan (401) dan Submisi Final.
        Returns: 'sukses', 'retry', atau 'stop'.
        """
        self.ketuk_sidebar_toggle()
        self.pilih_blok("IV")

        print("[BLOK IV] Mengisi catatan '-'...")
        self.input_textbox("Catatan", "-", exact=False, sleep_after=0.05)

        max_submit_retries = 3
        for submit_retry in range(1, max_submit_retries + 1):
            print(f"[SUBMIT BLOK IV] Mengetuk tombol 'Kirim' (Percobaan {submit_retry}/{max_submit_retries})...")
            self.ketuk("Kirim", sleep_after=0.05)

            if self.smart_wait_element(self.d(text="YA"), timeout=1.0, poll_interval=0.04):
                self.ketuk("YA", sleep_after=0.05)

            is_galat_0 = False
            # Smart wait hingga modal ringkasan validasi muncul
            matched_idx, _ = self.smart_wait_any([
                lambda: (
                    self.check_exists(self.d(textContains="GALAT 0 Perlu diperbaiki")) or
                    self.check_exists(self.d(textContains="GALAT 0")) or
                    self.check_exists(self.d(descriptionContains="GALAT 0")) or
                    self.check_exists(self.d.xpath("//*[contains(@text, 'GALAT 0') or contains(@content-desc, 'GALAT 0')]"))
                ),
                lambda: self.check_exists(self.d(textContains="GALAT")) or self.check_exists(self.d(text="Dismiss")),
            ], timeout=3.0, poll_interval=0.04)

            if matched_idx == 0:
                is_galat_0 = True

            if not is_galat_0:
                print(f"[SUBMIT] GALAT ≠ 0 terdeteksi untuk IDPEL {idpel}. Mencoba perbaikan otomatis...")
                fix_berhasil = self.galat_fixer.perbaiki_galat_koordinat_foto()
                if fix_berhasil:
                    self.ketuk("Kirim", sleep_after=0.05)
                    if self.smart_wait_element(self.d(text="YA"), timeout=1.0, poll_interval=0.04):
                        self.ketuk("YA", sleep_after=0.05)

            if not self.submit_service.pause_proses_galat_0(idpel=idpel, row=row, keterangan="Submit Blok IV"):
                return "stop"

            res_exec = self.submit_service.eksekusi_submit_dan_selesai(row, idpel, row_attempt)
            if res_exec == "retry_kirim" and submit_retry < max_submit_retries:
                continue
            return res_exec

        return "sukses"
