import time
from konfigurasi import (
    SLEEP_SHORT,
    SLEEP_MEDIUM,
    SLEEP_LONG_REJECT as SLEEP_LONG,
)


class BasePage:
    """Kelas dasar yang menyediakan utilitas interaksi UI uiautomator2 dan recovery layar."""

    def __init__(self, driver):
        self.driver = driver

    @property
    def d(self):
        return self.driver.d

    def smart_wait(self, condition_fn, timeout: float = 2.0, poll_interval: float = 0.05):
        """
        Menunggu kondisi callable/lambda bernilai truthy.
        Segera return hasil kondisi begitu terpenuhi tanpa menunggu timeout habis.
        """
        t_end = time.time() + timeout
        while time.time() < t_end:
            try:
                res = condition_fn()
                if res:
                    return res
            except Exception:
                pass
            time.sleep(poll_interval)
        return False

    def smart_wait_element(self, el_or_selector, timeout: float = 2.0, poll_interval: float = 0.05) -> bool:
        """Menunggu elemen muncul di layar (Selector / XPath / callable)."""
        def _check():
            if callable(el_or_selector):
                return bool(el_or_selector())
            return self.check_exists(el_or_selector)
        return bool(self.smart_wait(_check, timeout=timeout, poll_interval=poll_interval))

    def smart_wait_gone(self, el_or_selector, timeout: float = 2.0, poll_interval: float = 0.05) -> bool:
        """Menunggu elemen menghilang dari layar."""
        def _check_gone():
            if callable(el_or_selector):
                return not bool(el_or_selector())
            return not self.check_exists(el_or_selector)
        return bool(self.smart_wait(_check_gone, timeout=timeout, poll_interval=poll_interval))

    def smart_wait_any(self, conditions_list, timeout: float = 2.0, poll_interval: float = 0.05):
        """
        Menunggu salah satu dari list kondisi terpenuhi.
        Returns tuple: (matched_index, result) atau (None, None)
        """
        t_end = time.time() + timeout
        while time.time() < t_end:
            for idx, cond in enumerate(conditions_list):
                try:
                    val = cond() if callable(cond) else self.check_exists(cond)
                    if val:
                        return idx, val
                except Exception:
                    pass
            time.sleep(poll_interval)
        return None, None

    def check_exists(self, el) -> bool:
        """Helper universal memeriksa keberadaan elemen Selector maupun XPath."""
        return self.driver.check_exists(el)

    def ketuk(self, target_text: str, exact: bool = False, sleep_after: float = SLEEP_SHORT) -> bool:
        """
        Fungsi universal untuk mencari dan mengetuk elemen teks/description/XPath.
        Memprioritaskan elemen berjenis android.widget.Button atau yang clickable=True.
        """
        if not self.d:
            return False

        print(f"[KLIK] Mencari dan mengetuk '{target_text}'...")
        target_btn = None

        # 1. Prioritas 1: Target spesifik android.widget.Button
        try:
            btn_el = self.d(className="android.widget.Button", text=target_text)
            if not btn_el.exists():
                btn_el = self.d(className="android.widget.Button", textContains=target_text)
            if btn_el.exists():
                target_btn = btn_el
                print(f"[KLIK] Ditemukan via android.widget.Button: '{target_text}'")
        except Exception:
            pass

        # 2. Prioritas 2: Target elemen dengan clickable=True
        if not target_btn:
            try:
                click_el = self.d(text=target_text, clickable=True)
                if not click_el.exists():
                    click_el = self.d(textContains=target_text, clickable=True)
                if click_el.exists():
                    target_btn = click_el
                    print(f"[KLIK] Ditemukan via clickable elemen: '{target_text}'")
            except Exception:
                pass

        # 3. Prioritas 3: Target teks biasa
        if not target_btn:
            try:
                txt_el = self.d(text=target_text) if exact else self.d(textContains=target_text)
                if not txt_el.exists():
                    txt_el = self.d(description=target_text) if exact else self.d(descriptionContains=target_text)
                if txt_el.exists():
                    target_btn = txt_el
                    print(f"[KLIK] Ditemukan via teks/description: '{target_text}'")
            except Exception:
                pass

        # 4. Prioritas 4: Fallback XPath
        if not target_btn:
            try:
                if exact:
                    xp = f"//*[@text='{target_text}' or @content-desc='{target_text}']"
                else:
                    xp = f"//*[contains(@text, '{target_text}') or contains(@content-desc, '{target_text}')]"
                if self.check_exists(self.d.xpath(xp)):
                    target_btn = self.d.xpath(xp)
                    print(f"[KLIK] Ditemukan via XPath: {xp}")
            except Exception:
                pass

        # Eksekusi klik
        if target_btn:
            try:
                target_btn.click()
                print(f"[KLIK] Berhasil mengetuk '{target_text}'")
                if sleep_after > 0:
                    time.sleep(sleep_after)
                return True
            except Exception as e:
                print(f"[WARNING] Gagal mengeksekusi klik pada '{target_text}': {e}")

        print(f"[WARNING] Elemen '{target_text}' tidak ditemukan untuk diketuk.")
        return False

    def input_textbox(self, label_text: str, value: str, bounds_fallback: tuple = None, exact: bool = False, sleep_after: float = SLEEP_MEDIUM) -> bool:
        """Mengisi field EditText berdasarkan label teks (TextView) dengan verifikasi set_text."""
        if value is None:
            print(f"[INPUT] Nilai untuk '{label_text}' kosong, melewati pengisian.")
            return False

        val_str = str(value).strip()
        print(f"[INPUT] Mencari field '{label_text}' untuk menginput: '{val_str}'...")

        target_input = None

        # 1. Target via .down() / .sibling() dari label
        try:
            if exact:
                label = self.d(text=label_text)
                if not self.check_exists(label):
                    label = self.d(description=label_text)
            else:
                label = self.d(textContains=label_text)
                if not self.check_exists(label):
                    label = self.d(descriptionContains=label_text)

            if self.check_exists(label):
                temp_input = label.down(className="android.widget.EditText")
                if self.check_exists(temp_input):
                    target_input = temp_input
                else:
                    temp_sibling = label.sibling(className="android.widget.EditText")
                    if self.check_exists(temp_sibling):
                        target_input = temp_sibling
        except Exception:
            pass

        # 2. Target via XPath (following EditText)
        if not self.check_exists(target_input):
            try:
                if exact:
                    xp = f"//*[@text='{label_text}' or @content-desc='{label_text}']/following::android.widget.EditText[1]"
                else:
                    xp = f"//*[contains(@text, '{label_text}') or contains(@content-desc, '{label_text}')]/following::android.widget.EditText[1]"

                if self.check_exists(self.d.xpath(xp)):
                    target_input = self.d.xpath(xp)
                    print(f"[INPUT] Ditemukan via XPath: {xp}")
            except Exception:
                pass

        # 3. Eksekusi Set Text
        success = False
        if self.check_exists(target_input):
            try:
                target_input.click()
                time.sleep(SLEEP_SHORT)
                try:
                    if hasattr(target_input, "clear_text") and callable(target_input.clear_text):
                        target_input.clear_text()
                except Exception:
                    pass

                if hasattr(target_input, "set_text") and callable(target_input.set_text):
                    target_input.set_text(val_str)
                else:
                    self.d.send_keys(val_str)
                print(f"[INPUT] Berhasil mengisi '{label_text}' dengan: '{val_str}'")
                success = True
            except Exception as e:
                print(f"[WARNING] Gagal set_text pada '{label_text}': {e}")

        # 4. Fallback koordinat jika bounds_fallback diberikan atau default khusus '202'
        if not success:
            if bounds_fallback:
                cx, cy = bounds_fallback
                print(f"[INPUT] Mengetuk fallback koordinat ({cx}, {cy}) untuk '{label_text}'...")
                try:
                    self.d.click(cx, cy)
                    time.sleep(SLEEP_SHORT)
                    self.d.send_keys(val_str)
                    print(f"[INPUT] Berhasil mengisi '{label_text}' via koordinat ({cx}, {cy}): '{val_str}'")
                    success = True
                except Exception as err:
                    print(f"[ERROR] Gagal mengisi '{label_text}' via koordinat: {err}")
            elif "202" in label_text or "NIK" in label_text:
                print(f"[INPUT] Mengetuk fallback koordinat (540, 1156) untuk '{label_text}'...")
                try:
                    self.d.click(540, 1156)
                    time.sleep(SLEEP_SHORT)
                    self.d.send_keys(val_str)
                    print(f"[INPUT] Berhasil mengisi '{label_text}' via koordinat (540, 1156): '{val_str}'")
                    success = True
                except Exception as err:
                    print(f"[ERROR] Gagal mengisi '{label_text}' via koordinat: {err}")

        if sleep_after > 0:
            time.sleep(sleep_after)
        return success

    def tunggu_loading(self, target_text: str = None, timeout: int = 30, sleep_before: float = 0.1) -> bool:
        """Menunggu animasi progress dialog hilang atau target text muncul di layar."""
        if sleep_before > 0:
            time.sleep(sleep_before)

        progress_ids = ["id.go.bpsfasih:id/card_progress", "id.go.bpsfasih:id/progressBar"]
        t_end = time.time() + timeout
        while time.time() < t_end:
            is_loading = False
            for pid in progress_ids:
                if self.check_exists(self.d(resourceId=pid)):
                    is_loading = True
                    break

            if not is_loading and self.check_exists(self.d(className="android.widget.ProgressBar")):
                is_loading = True

            if target_text:
                if (
                    self.check_exists(self.d(text=target_text))
                    or self.check_exists(self.d(textContains=target_text))
                    or self.check_exists(self.d(description=target_text))
                    or self.check_exists(self.d(descriptionContains=target_text))
                ):
                    return True

            if not is_loading and not target_text:
                return True
            if not is_loading:
                time.sleep(0.05)
                return True
            time.sleep(0.05)

        return False

    def tunggu_loading_cek_nik(self, timeout: int = 30, sleep_before: float = 0.1) -> bool:
        """Menunggu proses loading 'Cek NIK' selesai dengan polling progress bar reaktif."""
        if sleep_before > 0:
            time.sleep(sleep_before)

        progress_selectors = [
            self.d(resourceId="id.go.bpsfasih:id/card_progress"),
            self.d(className="android.widget.ProgressBar"),
            self.d(textContains="Memuat"),
            self.d(textContains="Loading"),
        ]

        t_end = time.time() + timeout
        while time.time() < t_end:
            is_loading = False
            for sel in progress_selectors:
                if self.check_exists(sel):
                    is_loading = True
                    break

            if not is_loading:
                time.sleep(0.05)
                return True
            time.sleep(0.05)

        return False

    def tutup_modal_dismiss_aman(self):
        """Menutup dialog/modal validasi (Dismiss) secara cepat dengan smart wait."""
        for attempt in range(1, 3):
            try:
                if self.check_exists(self.d(text="Dismiss")) or self.check_exists(self.d(textContains="Dismiss")):
                    print(f"[DISMISS] Mengetuk 'Dismiss' (percobaan {attempt})...")
                    self.ketuk("Dismiss", sleep_after=0.03)
                    self.smart_wait_gone(self.d(text="Dismiss"), timeout=0.8, poll_interval=0.04)
                else:
                    break
            except Exception as err:
                print(f"[DISMISS] Modal sudah tertutup ({err}).")
                break

    def check_dan_tutup_pengaturan(self) -> bool:
        """Mengecek apakah dialog 'Pengaturan' muncul secara tidak sengaja, lalu mengetuk 'Batal'."""
        try:
            is_pengaturan = (
                self.check_exists(self.d(text="Pengaturan"))
                or self.check_exists(self.d(textContains="Pengaturan"))
                or self.check_exists(self.d(description="Pengaturan"))
                or self.check_exists(self.d(descriptionContains="Pengaturan"))
            )
            if is_pengaturan:
                print("[PENGATURAN] Terdeteksi dialog 'Pengaturan'. Mencoba mengetuk 'Batal'...")
                clicked = False
                for btn_text in ["Batal", "BATAL", "batal", "Cancel", "CANCEL"]:
                    btn = self.d(text=btn_text)
                    if btn.exists():
                        btn.click()
                        clicked = True
                        break

                if not clicked:
                    self.d.click(646, 1677)

                time.sleep(SLEEP_SHORT)
                self.driver.loop_swipe_statis(delta_y=700, loop=1)
                time.sleep(SLEEP_SHORT)
                return True
        except Exception as e:
            print(f"[WARNING] Error pengecekan Pengaturan: {e}")
        return False

    def kembali_ke_daftar_assignment(self, max_retry: int = 5) -> bool:
        """Prosedur recovery menekan tombol BACK dan menutup modal hingga kembali ke halaman 'Daftar Assignment'."""
        print("[RECOVERY] Memulai prosedur kembali ke halaman 'Daftar Assignment'...")
        for i in range(1, max_retry + 1):
            try:
                # Tutup dialog jika ada
                for btn_name in ["Dismiss", "Batal", "Tutup", "OK"]:
                    if self.check_exists(self.d(text=btn_name)):
                        print(f"[RECOVERY] Menutup dialog via tombol '{btn_name}'...")
                        self.d(text=btn_name).click()
                        time.sleep(SLEEP_SHORT)

                # Cek apakah sudah di Daftar Assignment
                if (
                    self.check_exists(self.d(textContains="Daftar Assignment"))
                    or self.check_exists(self.d(descriptionContains="Daftar Assignment"))
                    or self.check_exists(self.d(text="Search"))
                ):
                    print(f"[RECOVERY] Berhasil kembali ke 'Daftar Assignment' pada percobaan ke-{i}!")
                    time.sleep(SLEEP_SHORT)
                    return True

                print(f"[RECOVERY] Menekan tombol BACK (Percobaan {i}/{max_retry})...")
                self.d.press("back")
                time.sleep(SLEEP_LONG)

                if (
                    self.check_exists(self.d(textContains="Daftar Assignment"))
                    or self.check_exists(self.d(descriptionContains="Daftar Assignment"))
                    or self.check_exists(self.d(text="Search"))
                ):
                    print(f"[RECOVERY] Berhasil kembali ke 'Daftar Assignment' pada percobaan ke-{i}!")
                    time.sleep(SLEEP_SHORT)
                    return True
            except Exception as e:
                print(f"[RECOVERY] Error pada percobaan ke-{i}: {e}")

        print("[RECOVERY] Gagal kembali ke 'Daftar Assignment' setelah batas maksimal percobaan.")
        return False
