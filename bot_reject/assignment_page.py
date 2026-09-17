import time
from konfigurasi import (
    SLEEP_SHORT,
    SLEEP_MEDIUM,
    SLEEP_LONG_REJECT as SLEEP_LONG,
)
from .base_page import BasePage


class AssignmentPage(BasePage):
    """Mengelola pencarian IDPEL dan pembukaan kuesioner dari halaman Daftar Assignment."""

    def cari_dan_ketuk_search_box(self, idpel: str) -> bool:
        """Mencari text box search, memasukkan IDPEL, dan menekan tombol Enter."""
        print(f"\n[SEARCH] Mencari text box search di layar untuk IDPEL '{idpel}'...")
        search_input = None

        # 1. Target spesifik via hint="Search:" dan className EditText
        try:
            if self.d(className="android.widget.EditText", hint="Search:").exists():
                search_input = self.d(className="android.widget.EditText", hint="Search:")
        except Exception:
            pass

        # 2. Target via XPath hint="Search:"
        if not search_input:
            try:
                xp_el = self.d.xpath("//android.widget.EditText[@hint='Search:']")
                if xp_el.exists:
                    search_input = xp_el
            except Exception:
                pass

        # 3. Target via Sibling dari TextView "Search:"
        if not search_input:
            try:
                if self.d(text="Search:").exists():
                    search_input = self.d(text="Search:").sibling(className="android.widget.EditText")
                elif self.d(textContains="Search:").exists():
                    search_input = self.d(textContains="Search:").sibling(className="android.widget.EditText")
            except Exception:
                pass

        # 4. Target via EditText pertama di layar
        if not search_input:
            try:
                if self.d(className="android.widget.EditText").exists():
                    search_input = self.d(className="android.widget.EditText")
            except Exception:
                pass

        # 5. Fallback koordinat bounds (618, 826)
        if not self.check_exists(search_input):
            print("[WARNING] Text box search tidak terdeteksi via selector. Mengklik koordinat fallback (618, 826)...")
            self.d.click(618, 826)
            time.sleep(SLEEP_SHORT)
            if self.d(className="android.widget.EditText").exists():
                search_input = self.d(className="android.widget.EditText")

        if self.check_exists(search_input):
            print("[SEARCH] Mengetuk text box search...")
            try:
                search_input.click()
            except Exception:
                self.d.click(618, 826)
            time.sleep(SLEEP_SHORT)

            print(f"[INPUT] Memasukkan IDPEL: '{idpel}'...")
            try:
                if hasattr(search_input, "set_text"):
                    search_input.set_text(str(idpel))
                else:
                    self.d.send_keys(str(idpel))
            except Exception:
                self.d.send_keys(str(idpel))
            time.sleep(SLEEP_SHORT)

            print("[ENTER] Menekan tombol Enter...")
            self.d.press("enter")
            time.sleep(SLEEP_MEDIUM)
            return True
        else:
            print("[SEARCH] Mencoba klik koordinat (618, 826) & send_keys langsung...")
            self.d.click(618, 826)
            time.sleep(SLEEP_SHORT)
            self.d.send_keys(str(idpel))
            time.sleep(SLEEP_SHORT)
            self.d.press("enter")
            time.sleep(SLEEP_MEDIUM)
            return True

    def cek_no_matching_records(self) -> bool:
        """Mengecek apakah hasil pencarian menampilkan 'No matching records found'."""
        try:
            return bool(
                self.check_exists(self.d(textContains="No matching records"))
                or self.check_exists(self.d(descriptionContains="No matching records"))
                or self.check_exists(self.d.xpath("//*[contains(@text, 'No matching records') or contains(@content-desc, 'No matching records')]"))
            )
        except Exception:
            return False

    def cek_dan_ketuk_plus_minus(self, nometer: str = "", idpel: str = "", sleep_after: float = SLEEP_SHORT, bounds_str: str = "[48,1248][435,1482]") -> bool:
        """Mengecek dan mengetuk simbol plus (+) atau minus (-) pada item baris penugasan."""
        val_str = str(nometer).strip() if (nometer and str(nometer).strip()) else str(idpel).strip()
        target_plus_text = f"+ {val_str}" if val_str else "+"
        target_minus_text = f"- {val_str}" if val_str else "-"

        # 1. Cek apakah simbol minus (-) sudah terbuka
        is_minus_detected = False
        try:
            if (
                self.check_exists(self.d(textContains=target_minus_text))
                or self.check_exists(self.d(descriptionContains=target_minus_text))
                or self.check_exists(self.d.xpath(f"//*[contains(@text, '{target_minus_text}') or contains(@content-desc, '{target_minus_text}')]"))
            ):
                is_minus_detected = True
            elif val_str and (self.check_exists(self.d(textContains=f"- {val_str}")) or self.check_exists(self.d(descriptionContains=f"- {val_str}"))):
                is_minus_detected = True
        except Exception:
            pass

        if is_minus_detected:
            print(f"[PLUS/MINUS] Baris detail sudah terbuka (terdeteksi simbol minus '{target_minus_text}'). Melewati ketuk.")
            return True

        # 2. Ketuk simbol plus (+)
        print(f"[PLUS/MINUS] Mengetuk simbol plus '{target_plus_text}'...")
        sukses_plus = False
        try:
            if self.d(textContains=target_plus_text).exists():
                self.d(textContains=target_plus_text).click()
                sukses_plus = True
            elif self.d(text=target_plus_text).exists():
                self.d(text=target_plus_text).click()
                sukses_plus = True
            elif self.d.xpath(f"//*[contains(@text, '{target_plus_text}') or contains(@content-desc, '{target_plus_text}')]").exists:
                self.d.xpath(f"//*[contains(@text, '{target_plus_text}') or contains(@content-desc, '{target_plus_text}')]").click()
                sukses_plus = True
            elif self.d.xpath("//*[starts-with(@text, '+') or starts-with(@content-desc, '+')]").exists:
                self.d.xpath("//*[starts-with(@text, '+') or starts-with(@content-desc, '+')]").click()
                sukses_plus = True
            elif bounds_str and self.d.xpath(f"//*[@bounds='{bounds_str}']").exists:
                self.d.xpath(f"//*[@bounds='{bounds_str}']").click()
                sukses_plus = True
        except Exception:
            pass

        # 3. Fallback koordinat (241, 1365)
        if not sukses_plus:
            print("[PLUS/MINUS] Fallback mengetuk koordinat (241, 1365)...")
            try:
                self.d.click(241, 1365)
                sukses_plus = True
            except Exception:
                sukses_plus = self.ketuk(val_str if val_str else "+", sleep_after=sleep_after)

        if sleep_after > 0:
            time.sleep(sleep_after)

        return sukses_plus

    def cek_status_penugasan(self) -> str:
        """
        Mengecek status penugasan setelah baris dibuka.
        Returns: 'SUBMITTED', 'REJECTED', atau 'UNKNOWN'
        """
        try:
            if (
                self.check_exists(self.d(textContains="SUBMITTED"))
                or self.check_exists(self.d(descriptionContains="SUBMITTED"))
                or self.check_exists(self.d(textContains="SUBMIT (PENDING)"))
                or self.check_exists(self.d(descriptionContains="SUBMIT (PENDING)"))
                or self.check_exists(self.d.xpath("//*[contains(@text, 'SUBMITTED') or contains(@content-desc, 'SUBMITTED') or contains(@text, 'SUBMIT (PENDING)')]"))
            ):
                return "SUBMITTED"

            if (
                self.check_exists(self.d(textContains="REJECTED"))
                or self.check_exists(self.d(descriptionContains="REJECTED"))
                or self.check_exists(self.d(textContains="Responden menolak"))
                or self.check_exists(self.d(descriptionContains="Responden menolak"))
                or self.check_exists(self.d.xpath("//*[contains(@text, 'REJECTED') or contains(@content-desc, 'REJECTED') or contains(@text, 'Responden menolak')]"))
            ):
                return "REJECTED"
        except Exception:
            pass

        return "UNKNOWN"

    def tangani_dialog_download(self) -> bool:
        """Mengecek dan menangani dialog 'Perhatian' (Download Sekarang) jika muncul secara reaktif."""
        # Polling singkat (maks 0.4s) untuk mendeteksi apakah dialog muncul
        dialog_detected = self.smart_wait(
            lambda: (
                self.check_exists(self.d(text="Perhatian"))
                or self.check_exists(self.d(textContains="Perhatian"))
                or self.check_exists(self.d(resourceId="id.go.bpsfasih:id/judul_bottomDialog"))
            ),
            timeout=0.4,
            poll_interval=0.05,
        )

        if not dialog_detected:
            return False

        try:
            print("[DIALOG] Terdeteksi dialog 'Perhatian'. Mengetuk 'DOWNLOAD SEKARANG'...")
            if self.check_exists(self.d(resourceId="id.go.bpsfasih:id/rButton_bottomDialog")):
                self.d(resourceId="id.go.bpsfasih:id/rButton_bottomDialog").click()
            else:
                self.ketuk("DOWNLOAD SEKARANG", sleep_after=SLEEP_SHORT)

            print("[LOADING] Menunggu proses download kuesioner selesai...")
            self.tunggu_loading(timeout=30)
            time.sleep(0.1)

            # Ulangi ketuk 'Aksi' -> 'BUKA' -> 'YA' setelah download selesai
            print("[RETRY] Mengulangi ketuk 'Aksi' -> 'BUKA' -> 'YA'...")
            self.ketuk("Aksi", sleep_after=SLEEP_SHORT)
            self.smart_wait_element(self.d(text="BUKA"), timeout=1.5, poll_interval=0.05)
            self.ketuk("BUKA", sleep_after=SLEEP_SHORT)
            self.smart_wait_element(self.d(text="YA"), timeout=1.5, poll_interval=0.05)
            self.ketuk("YA", sleep_after=0.1)
            return True
        except Exception as e:
            print(f"[WARNING] Error handling dialog download: {e}")
        return False

    def buka_detail_kuesioner(self, nometer: str, idpel: str, detail_attempt: int = 1) -> tuple:
        """
        Membuka kuesioner dari Daftar Assignment dengan smart wait.
        Returns tuple: (sukses: bool, status_awal: str, is_still_da: bool)
        """
        # 1. Expand row (+ / -)
        self.cek_dan_ketuk_plus_minus(nometer=nometer, idpel=idpel, sleep_after=SLEEP_SHORT)

        # 2. Swipe ke bawah untuk melihat status & tombol Aksi
        self.driver.loop_swipe_statis(delta_y=-700, loop=1)
        time.sleep(SLEEP_SHORT)

        # 3. Cek status
        status = self.cek_status_penugasan()
        if status == "SUBMITTED":
            print(f"[STATUS] Terdeteksi status 'SUBMITTED' pada IDPEL {idpel}. Melewati kuesioner.")
            return True, "SUBMITTED", False

        # 4. Ketuk Aksi -> BUKA -> YA secara reaktif
        self.driver.loop_swipe_statis(delta_y=-700, loop=1)
        time.sleep(SLEEP_SHORT)

        self.ketuk("Aksi", sleep_after=SLEEP_SHORT)
        self.smart_wait_element(self.d(text="BUKA"), timeout=1.5, poll_interval=0.05)

        self.ketuk("BUKA", sleep_after=SLEEP_SHORT)
        self.smart_wait_element(self.d(text="YA"), timeout=1.5, poll_interval=0.05)

        self.ketuk("YA", sleep_after=0.1)

        # 5. Handle dialog download jika ada
        self.tangani_dialog_download()

        # 6. Scan apakah form kuesioner sudah terbuka secara cepat
        print("[SCAN] Menunggu form kuesioner terbuka (timeout 30s)...")
        is_opened = False
        t_end = time.time() + 30
        while time.time() < t_end:
            if (
                self.check_exists(self.d(textContains="Mulai Wawancara"))
                or self.check_exists(self.d(descriptionContains="Mulai Wawancara"))
                or self.check_exists(self.d(text="Kirim"))
                or self.check_exists(self.d.xpath("//*[contains(@text, 'Mulai Wawancara') or contains(@content-desc, 'Mulai Wawancara') or @text='Kirim']"))
            ):
                is_opened = True
                print("[SCAN] [SUKSES] Form kuesioner berhasil terbuka di layar!")
                break
            time.sleep(0.1)

        if not is_opened:
            is_still_da = (
                self.check_exists(self.d(textContains="Daftar Assignment"))
                or self.check_exists(self.d(descriptionContains="Daftar Assignment"))
                or self.check_exists(self.d.xpath("//*[contains(@text, 'Daftar Assignment') or contains(@content-desc, 'Daftar Assignment')]"))
            )
            return False, status, is_still_da

        return True, status, False
