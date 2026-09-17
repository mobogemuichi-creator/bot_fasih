import time
from konfigurasi import (
    SLEEP_SHORT,
    SLEEP_MEDIUM,
    SLEEP_LONG_REJECT as SLEEP_LONG,
)
from .base_page import BasePage


class GalatFixer(BasePage):
    """Mengelola perbaikan otomatis error Galat 105 (Koordinat Lokasi) dan Galat 106 (Foto Rumah)."""

    def perbaiki_galat_koordinat_foto(self, skip_ketuk_galat: bool = False) -> bool:
        """
        Ketika GALAT ≠ 0, coba perbaiki error '105. Koordinat lokasi meteran'
        dan/atau '106. Foto rumah tampak depan' secara otomatis.
        Returns True jika berhasil diperbaiki, False jika error bukan 105/106 atau gagal.
        """
        print("\n[GALAT FIX] === MEMULAI PERBAIKAN ERROR GALAT ===")

        # 1. Ketuk teks 'GALAT' jika belum diketuk
        if not skip_ketuk_galat:
            print("[GALAT FIX] Mengetuk 'GALAT' pada modal ringkasan validasi...")
            galat_clicked = False
            try:
                galat_el = self.d(textContains="GALAT")
                if galat_el.exists(timeout=3):
                    galat_el.click()
                    galat_clicked = True
            except Exception:
                pass
            if not galat_clicked:
                try:
                    xp = self.d.xpath("//*[contains(@text, 'GALAT') or contains(@content-desc, 'GALAT')]")
                    if xp.exists:
                        xp.click()
                        galat_clicked = True
                except Exception:
                    pass

            if not galat_clicked:
                print("[GALAT FIX] Tidak bisa mengetuk 'GALAT'. Membatalkan perbaikan.")
                return False

            time.sleep(SLEEP_MEDIUM)
        else:
            print("[GALAT FIX] Skip ketuk 'GALAT' (sudah diketuk sebelumnya).")

        # 2. Cek apakah ada error 105 atau 106
        has_105 = (
            self.check_exists(self.d(textContains="105. Koordinat lokasi meteran"))
            or self.check_exists(self.d(textContains="105."))
            or self.check_exists(self.d(descriptionContains="105."))
        )
        has_106 = (
            self.check_exists(self.d(textContains="106. Foto rumah tampak depan"))
            or self.check_exists(self.d(textContains="106."))
            or self.check_exists(self.d(descriptionContains="106."))
        )

        print(f"[GALAT FIX] Deteksi error: 105 (Koordinat)={has_105}, 106 (Foto)={has_106}")
        if not has_105 and not has_106:
            print("[GALAT FIX] Error bukan 105/106. Tidak bisa diperbaiki otomatis.")
            return False

        # 3. Ketuk tombol 'Lihat'
        print("[GALAT FIX] Mengetuk tombol 'Lihat'...")
        lihat_clicked = self.ketuk("Lihat", sleep_after=SLEEP_SHORT)
        if not lihat_clicked:
            lihat_clicked = self.ketuk("lihat", sleep_after=SLEEP_SHORT)
        if not lihat_clicked:
            lihat_clicked = self.ketuk("LIHAT", sleep_after=SLEEP_SHORT)

        if not lihat_clicked:
            print("[GALAT FIX] Tombol 'Lihat' tidak ditemukan. Membatalkan.")
            return False

        time.sleep(SLEEP_MEDIUM)

        # 4. Tunggu halaman field dan scroll ke bawah dengan smart wait
        print("[GALAT FIX] Menunggu halaman field dimuat secara reaktif...")
        self.smart_wait(
            lambda: (
                self.check_exists(self.d(textContains="Koordinat lokasi meteran"))
                or self.check_exists(self.d(textContains="Foto rumah tampak depan"))
                or self.check_exists(self.d(text="Pilih", className="android.widget.Button"))
                or self.check_exists(self.d(text="Ambil Lokasi"))
            ),
            timeout=2.0,
            poll_interval=0.05,
        )

        print("[GALAT FIX] Men-scroll ke bawah untuk menemukan field 105/106...")
        for scroll_idx in range(1, 16):
            if (
                self.check_exists(self.d(textContains="Koordinat lokasi meteran"))
                or self.check_exists(self.d(textContains="Foto rumah tampak depan"))
                or self.check_exists(self.d(text="Pilih", className="android.widget.Button"))
                or self.check_exists(self.d(text="Ambil Lokasi"))
            ):
                print(f"[GALAT FIX] Field / tombol terkait ditemukan (scroll ke-{scroll_idx}).")
                break
            try:
                self.d.swipe(540, 1400, 540, 400, duration=0.1)
                time.sleep(0.3)
            except Exception:
                break
        time.sleep(SLEEP_SHORT)

        # 5. Perbaiki Error 106 jika ada
        if has_106:
            if not self._perbaiki_error_106():
                return False

        # 6. Perbaiki Error 105 jika ada
        if has_105:
            if not self._perbaiki_error_105():
                return False

        print("[GALAT FIX] === PERBAIKAN GALAT SELESAI ===\n")
        return True

    def _cek_status_foto_di_layar(self) -> str:
        for sel in [
            self.d(textContains="Sudah Terunggah"),
            self.d(textContains="sudah terunggah"),
            self.d(textContains="Dimuat dari server"),
            self.d(textContains="dimuat dari server"),
            self.d(textMatches=r"(?i).*(sudah terunggah|dimuat dari server).*"),
        ]:
            if self.check_exists(sel):
                return "sudah_terunggah"

        for sel in [
            self.d(textContains="Dimuat dari local"),
            self.d(textContains="Dimuat dari lokal"),
            self.d(textContains="dimuat dari local"),
            self.d(textContains="dimuat dari lokal"),
            self.d(text="Unggah Foto"),
            self.d(textContains="Unggah Foto"),
            self.d(textMatches=r"(?i).*dimuat dari lo[ck]al.*"),
        ]:
            if self.check_exists(sel):
                return "dimuat_local"

        return None

    def _scroll_ke_atas_setengah_layar(self):
        print("[GALAT FIX] Melakukan scroll ke atas setengah layar...")
        cx = self.driver.screen_width // 2
        sy = int(self.driver.screen_height * 0.3)
        ey = int(self.driver.screen_height * 0.8)
        self.driver.swipe_aman(cx, sy, cx, ey, duration=0.2)
        time.sleep(SLEEP_SHORT)

    def _eksekusi_unggah_foto(self) -> bool:
        print("[GALAT FIX] Mengetuk tombol 'Unggah Foto'...")
        unggah_btn = None
        for sel in [
            self.d(text="Unggah Foto"),
            self.d(textContains="Unggah Foto"),
            self.d(text="Unggah"),
            self.d(text="UNGGAH"),
            self.d(textContains="Unggah"),
        ]:
            if self.check_exists(sel):
                unggah_btn = sel
                break

        if unggah_btn:
            try:
                unggah_btn.click()
            except Exception as e:
                print(f"[WARNING] Gagal klik tombol unggah: {e}")
            time.sleep(1.0)

            # Konfirmasi Ya
            for ya in [self.d(text="Ya"), self.d(text="YA"), self.d(text="ya"), self.d(textContains="Ya")]:
                if self.check_exists(ya):
                    try:
                        ya.click()
                    except Exception:
                        pass
                    time.sleep(1.0)
                    break

        # Polling tunggu status 'Sudah Terunggah'
        print("[GALAT FIX] Memindai layar menunggu status 'Sudah Terunggah'...")
        for scan_idx in range(1, 16):
            time.sleep(1.0)
            status = self._cek_status_foto_di_layar()
            if status == "sudah_terunggah":
                print(f"[GALAT FIX] [SUKSES] Status 'Sudah Terunggah' terkonfirmasi (+{scan_idx}s)!")
                self._scroll_ke_atas_setengah_layar()
                return True
            if scan_idx == 6 and status == "dimuat_local":
                print("[GALAT FIX] Mencoba mengetuk 'Unggah Foto' ulang...")
                for sel in [self.d(text="Unggah Foto"), self.d(textContains="Unggah Foto"), self.d(text="Unggah")]:
                    if self.check_exists(sel):
                        try:
                            sel.click()
                            time.sleep(1.0)
                            ya = self.d(text="Ya")
                            if self.check_exists(ya):
                                ya.click()
                        except Exception:
                            pass
                        break

        if self._cek_status_foto_di_layar() == "sudah_terunggah":
            print("[GALAT FIX] [SUKSES] Status 'Sudah Terunggah' terkonfirmasi!")
            self._scroll_ke_atas_setengah_layar()
            return True
        return False

    def _perbaiki_error_106(self) -> bool:
        print("\n[GALAT FIX] === MEMPERBAIKI ERROR 106: FOTO RUMAH TAMPAK DEPAN ===")
        # Scroll statis 3x ke bawah
        for _ in range(3):
            self.driver.loop_swipe_statis(delta_y=-700, loop=1)
            time.sleep(0.3)
        time.sleep(SLEEP_SHORT)

        foto_berhasil = False
        status_awal = self._cek_status_foto_di_layar()
        if status_awal == "sudah_terunggah":
            print("[GALAT FIX] [SUKSES] Foto sudah terunggah / aktif di layar.")
            foto_berhasil = True
        elif status_awal == "dimuat_local":
            print("[GALAT FIX] Foto sudah ada di layar ('Dimuat dari local'). Langsung mengunggah...")
            if self._eksekusi_unggah_foto():
                foto_berhasil = True

        if not foto_berhasil:
            for attempt in range(1, 3):
                print(f"\n[GALAT FIX] === SELEKSI FOTO (Percobaan {attempt}/2) ===")
                time.sleep(0.5)

                pilih_foto_clicked = False
                for try_pilih in range(3):
                    try:
                        pilih_foto_btn = self.d(text="Pilih", className="android.widget.Button")
                        if not pilih_foto_btn.exists():
                            pilih_foto_btn = self.d(text="Pilih")
                        if pilih_foto_btn.exists(timeout=3):
                            pilih_foto_btn.click()
                            pilih_foto_clicked = True
                            time.sleep(SLEEP_MEDIUM)
                            break
                        time.sleep(0.5)
                    except Exception:
                        time.sleep(0.5)

                if not pilih_foto_clicked:
                    if self._cek_status_foto_di_layar() == "sudah_terunggah":
                        foto_berhasil = True
                        break
                    print("[WARNING] Tombol 'Pilih' tidak dapat diklik. Menekan BACK 2x...")
                    for _ in range(2):
                        self.d.press("back")
                        time.sleep(SLEEP_SHORT)
                    foto_berhasil = True
                    break

                # Ketuk GALERI
                galeri_btn = None
                for g_txt in ["GALERI", "Galeri", "galeri"]:
                    if self.d(text=g_txt).exists():
                        galeri_btn = self.d(text=g_txt)
                        break
                if galeri_btn and galeri_btn.exists(timeout=5):
                    galeri_btn.click()
                    time.sleep(SLEEP_MEDIUM)
                else:
                    self.d.press("back")
                    time.sleep(SLEEP_SHORT)
                    continue

                # Ketuk burger menu jika ada
                burger_btn = None
                for desc in ["Show roots", "Tampilkan laci", "Tampilkan root", "Show navigation drawer", "Open navigation drawer", "Laci navigasi", "Menu"]:
                    if self.d(descriptionContains=desc).exists():
                        burger_btn = self.d(descriptionContains=desc)
                        break
                if not burger_btn:
                    for res_id in ["android:id/home", "com.android.documentsui:id/toolbar"]:
                        if self.d(resourceId=res_id).exists():
                            burger_btn = self.d(resourceId=res_id)
                            break
                if burger_btn and burger_btn.exists(timeout=5):
                    burger_btn.click()
                    time.sleep(SLEEP_SHORT)

                # Ketuk Recent / Baru-baru ini / Terbaru
                print("[GALAT FIX] Mengetuk opsi 'Recent' atau 'Baru-baru ini'...")
                recent_btn = None
                recent_keywords = [
                    "Recent", "recent",
                    "Baru-baru ini", "baru-baru ini",
                    "Baru - baru ini", "baru - baru ini",
                    "Baru baru ini", "baru baru ini",
                    "Terbaru", "terbaru",
                ]
                for text_val in recent_keywords:
                    if self.d(text=text_val, resourceId="android:id/title").exists():
                        recent_btn = self.d(text=text_val, resourceId="android:id/title")
                        break
                    elif self.d(text=text_val, resourceId="com.android.documentsui:id/title").exists():
                        recent_btn = self.d(text=text_val, resourceId="com.android.documentsui:id/title")
                        break
                if not recent_btn:
                    for text_val in recent_keywords:
                        if self.d(text=text_val).exists():
                            recent_btn = self.d(text=text_val)
                            break
                if not recent_btn:
                    for text_val in ["Recent", "Baru-baru ini", "Baru baru ini", "Terbaru", "recent", "baru-baru ini", "terbaru"]:
                        if self.d(textContains=text_val).exists():
                            recent_btn = self.d(textContains=text_val)
                            break

                if recent_btn and recent_btn.exists(timeout=5):
                    print("[GALAT FIX] Opsi 'Recent'/'Baru-baru ini' ditemukan. Mengetuk...")
                    recent_btn.click()
                    time.sleep(1.5)

                # Pilih gambar pertama
                print("[GALAT FIX] Memilih gambar pertama di galeri...")
                time.sleep(1.5)
                first_file_btn = None
                try:
                    nameplates = self.d(resourceId="com.android.documentsui:id/nameplate")
                    if nameplates.exists() and len(nameplates) > 0:
                        first_file_btn = nameplates[0]
                except Exception:
                    pass

                if not first_file_btn:
                    folder_names = [
                        "pictures", "images", "gambar",
                        "recent", "terbaru",
                        "baru-baru ini", "baru - baru ini", "baru baru ini",
                        "downloads", "audio", "videos",
                    ]
                    for el in self.d(resourceId="android:id/title"):
                        txt = el.info.get("text", "").strip()
                        if txt and txt.lower() not in folder_names:
                            first_file_btn = el
                            break

                if first_file_btn and first_file_btn.exists():
                    bounds = first_file_btn.info.get("bounds")
                    click_x = (bounds["left"] + bounds["right"]) // 2
                    click_y = (bounds["top"] + bounds["bottom"]) // 2
                    self.d.click(click_x, click_y)
                else:
                    self.d.click(342, 998)

                time.sleep(1.5)
                if self._eksekusi_unggah_foto():
                    foto_berhasil = True
                    break

        return foto_berhasil

    def _perbaiki_error_105(self) -> bool:
        print("\n[GALAT FIX] === MEMPERBAIKI ERROR 105: KOORDINAT LOKASI METERAN ===")
        # Scroll ke atas untuk mencari 'Ambil Lokasi' jika belum terlihat
        ambil_lokasi_btn = self.d(text="Ambil Lokasi")
        if not ambil_lokasi_btn.exists():
            for swipe_up_idx in range(1, 8):
                if self.d(text="Ambil Lokasi").exists():
                    break
                self.driver.loop_swipe_statis(delta_y=700, loop=1)
                time.sleep(0.2)

        ambil_lokasi_btn = self.d(text="Ambil Lokasi")
        if ambil_lokasi_btn.wait(exists=True, timeout=5):
            print("[GALAT FIX] Mengetuk 'Ambil Lokasi'...")
            ambil_lokasi_btn.click()
            time.sleep(SLEEP_SHORT)
        else:
            print("[WARNING] 'Ambil Lokasi' tidak ditemukan. Gagal memperbaiki 105.")
            return False

        # Ketuk 'AMBIL LANGSUNG'
        opsi_lokasi = self.d(resourceId="id.go.bpsfasih:id/lButton_bottomDialog")
        if not opsi_lokasi.exists():
            opsi_lokasi = self.d(text="AMBIL LANGSUNG")
        if not opsi_lokasi.exists():
            opsi_lokasi = self.d(textContains="LANGSUNG")

        if opsi_lokasi.exists(timeout=5):
            opsi_lokasi.click()
            time.sleep(SLEEP_SHORT)
        else:
            print("[WARNING] 'AMBIL LANGSUNG' tidak ditemukan.")
            return False

        # Handle dialog keluar jika muncul
        dialog_keluar = self.d(textContains="Apakah Anda yakin akan keluar")
        if dialog_keluar.exists(timeout=2):
            for t_val in ["tidak", "Tidak", "TIDAK"]:
                if self.d(text=t_val).exists():
                    self.d(text=t_val).click()
                    time.sleep(SLEEP_SHORT)
                    break

        # Konfirmasi Ya lokasi
        for y_val in ["ya", "Ya", "YA"]:
            if self.d(text=y_val).exists(timeout=5):
                self.d(text=y_val).click()
                time.sleep(SLEEP_MEDIUM)
                break

        print("[GALAT FIX] Lokasi GPS berhasil diambil.")
        return True
