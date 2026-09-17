import time
from konfigurasi import (
    SLEEP_SHORT,
    SLEEP_MEDIUM,
    SLEEP_LONG_REJECT as SLEEP_LONG,
)
try:
    from konfigurasi import PAUSE_ON_GALAT_0
except ImportError:
    PAUSE_ON_GALAT_0 = False

from .base_page import BasePage


class SubmitService(BasePage):
    """Mengelola alur pengiriman (submit) data kuesioner ke server BPS Fasih."""

    def __init__(self, driver, excel_manager, pause_on_galat_0: bool = PAUSE_ON_GALAT_0):
        super().__init__(driver)
        self.excel_manager = excel_manager
        self.pause_on_galat_0 = pause_on_galat_0

    def pause_proses_galat_0(self, idpel: str = "", row: int = 0, keterangan: str = "") -> bool:
        """
        Fungsi Pause interaktif saat status 'GALAT 0' terdeteksi.
        Diaktifkan atau dinonaktifkan via konfigurasi PAUSE_ON_GALAT_0 atau parameter class.
        """
        if not self.pause_on_galat_0:
            return True

        info_idpel = f"IDPEL: {idpel}" if idpel else ""
        info_row = f"Baris: {row}" if row else ""
        header_info = " | ".join(filter(None, [info_idpel, info_row]))
        info_ket = f" [{keterangan}]" if keterangan else ""

        print("\n" + "=" * 60)
        print(f"  [PAUSE - GALAT 0 TERDETEKSI]{info_ket}")
        if header_info:
            print(f"  {header_info}")
        print("=" * 60)
        print("  Pilihan:")
        print("  [1 / Enter / Y] : Lanjutkan proses submit data")
        print("  [2 / S / Stop]  : Stop seluruh proses bot")
        print("=" * 60)

        try:
            jawaban = input(">> Pilihan Anda [Default: Lanjut (Enter)]: ").strip().lower()
            if jawaban in ["2", "s", "stop", "t", "tidak", "exit", "q", "keluar", "no", "n"]:
                print("[STOP] Pengguna memilih untuk MENGHENTIKAN SELURUH PROSES bot.")
                return False
            print("[LANJUT] Pengguna memilih MELANJUTKAN proses submit...")
            return True
        except (KeyboardInterrupt, EOFError):
            print("\n[STOP] Interupsi terdeteksi. Menghentikan seluruh proses bot.")
            return False

    def ketuk_ok_submit_diproses(self, max_attempts: int = 10):
        """
        Mengetuk tombol 'OK' pada modal 'Submit diproses'.
        Jika layar kembali ke 'BLOK IV', mengembalikan 'kembali_blok_iv'.
        """
        print("\n[SUBMIT DIPROSES] Memeriksa modal 'Submit diproses'...")
        ok_bounds_x, ok_bounds_y = 528, 1691

        for attempt in range(1, max_attempts + 1):
            is_blok_iv = (
                self.check_exists(self.d(textContains="BLOK IV"))
                or self.check_exists(self.d(textContains="Blok IV"))
                or self.check_exists(self.d(descriptionContains="BLOK IV"))
                or self.check_exists(self.d(descriptionContains="Blok IV"))
                or self.check_exists(self.d(textContains="401. Catatan"))
            )
            is_submit_modal = (
                self.check_exists(self.d(textContains="Submit diproses"))
                or self.check_exists(self.d(resourceId="id.go.bpsfasih:id/tv_submit_progress_title"))
                or self.check_exists(self.d(resourceId="id.go.bpsfasih:id/btn_submit_progress_close"))
                or self.check_exists(self.d(text="OK"))
            )

            if is_blok_iv and not self.check_exists(self.d(textContains="Submit diproses")):
                print("[SUBMIT DIPROSES] [RETRY TRIGGER] Terdeteksi layar kembali ke 'BLOK IV' saat pengetukan OK!")
                return "kembali_blok_iv"

            if not is_submit_modal:
                print("[SUBMIT DIPROSES] Modal 'Submit diproses' sudah tertutup.")
                return True

            print(f"[SUBMIT DIPROSES] Terdeteksi modal 'Submit diproses'. Percobaan ketuk 'OK' ke-{attempt}/{max_attempts}...")
            clicked = False
            try:
                btn_close = self.d(resourceId="id.go.bpsfasih:id/btn_submit_progress_close")
                if btn_close.exists():
                    btn_close.click()
                    clicked = True
                elif self.d(text="OK").exists():
                    self.d(text="OK").click()
                    clicked = True
                else:
                    clicked = self.ketuk("OK", sleep_after=SLEEP_SHORT)
            except Exception as e:
                print(f"[WARNING] Klik tombol OK via elemen gagal: {e}")

            if not clicked:
                print(f"[SUBMIT DIPROSES] Mengetuk koordinat statis bounds OK ({ok_bounds_x}, {ok_bounds_y})...")
                self.d.click(ok_bounds_x, ok_bounds_y)

            # Smart wait hingga modal submit diproses menghilang (maks 1.5s)
            self.smart_wait_gone(self.d(textContains="Submit diproses"), timeout=1.5, poll_interval=0.04)

            is_blok_iv_after = (
                self.check_exists(self.d(textContains="BLOK IV"))
                or self.check_exists(self.d(textContains="Blok IV"))
                or self.check_exists(self.d(descriptionContains="BLOK IV"))
                or self.check_exists(self.d(descriptionContains="Blok IV"))
                or self.check_exists(self.d(textContains="401. Catatan"))
            )
            if is_blok_iv_after and not self.check_exists(self.d(textContains="Submit diproses")):
                print("[SUBMIT DIPROSES] [RETRY TRIGGER] Terdeteksi layar kembali ke 'BLOK IV' setelah pengetukan OK!")
                return "kembali_blok_iv"

            still_exists = (
                self.check_exists(self.d(textContains="Submit diproses"))
                or self.check_exists(self.d(resourceId="id.go.bpsfasih:id/tv_submit_progress_title"))
            )
            if not still_exists:
                print(f"[SUBMIT DIPROSES] [SUKSES] Modal 'Submit diproses' berhasil ditutup pada percobaan ke-{attempt}.")
                return True

        # Fallback statis setelah max_attempts
        print(f"[SUBMIT DIPROSES] Fallback statis pengetukan OK ({ok_bounds_x}, {ok_bounds_y})...")
        self.d.click(ok_bounds_x, ok_bounds_y)
        self.smart_wait_gone(self.d(textContains="Submit diproses"), timeout=1.0, poll_interval=0.04)
        return True

    def eksekusi_submit_dan_selesai(self, row: int, idpel: str, row_attempt: int) -> str:
        """
        Rantai submisi penuh:
        1. Mengetuk 'Kirim' kedua di dalam modal
        2. Mengetuk 'Konfirmasi' -> 'Konfirmasi' -> 'YA'
        3. Mengetuk 'OK' pada progress 'Submit diproses'
        4. Deteksi 'Halaman Upload' -> Back
        5. Tunggu 'Daftar Assignment' dan simpan status 'SUKSES DIUPDATE' ke Excel.
        """
        print("[SUBMIT] [SUKSES] Mengetuk tombol 'Kirim' kedua di dalam modal...")
        self.ketuk("Kirim", sleep_after=0.03)
        self.smart_wait_element(self.d(text="Konfirmasi"), timeout=1.0, poll_interval=0.04)

        self.ketuk("Konfirmasi", sleep_after=0.03)
        self.smart_wait_element(self.d(text="Konfirmasi"), timeout=1.0, poll_interval=0.04)

        self.ketuk("Konfirmasi", sleep_after=0.03)
        self.smart_wait_element(self.d(text="YA"), timeout=1.0, poll_interval=0.04)

        max_ya_attempts = 5
        submit_muncul = False
        for ya_attempt in range(1, max_ya_attempts + 1):
            print(f"[SUBMIT] Mengetuk tombol 'YA' (Percobaan {ya_attempt}/{max_ya_attempts})...")
            self.ketuk("YA", sleep_after=0.03)

            # Smart wait hingga modal Submit diproses muncul
            is_submit_modal = self.smart_wait(
                lambda: (
                    self.check_exists(self.d(textContains="Submit diproses"))
                    or self.check_exists(self.d(resourceId="id.go.bpsfasih:id/tv_submit_progress_title"))
                    or self.check_exists(self.d(resourceId="id.go.bpsfasih:id/btn_submit_progress_close"))
                    or self.check_exists(self.d(text="OK"))
                ),
                timeout=1.5,
                poll_interval=0.04,
            )

            if is_submit_modal:
                print(f"[SUBMIT] [SUKSES] Modal 'Submit diproses' terdeteksi setelah ketuk 'YA' pada percobaan ke-{ya_attempt}.")
                submit_muncul = True
                break
            else:
                print(f"[SUBMIT] [RETRY] Modal 'Submit diproses' belum muncul (percobaan ke-{ya_attempt}/{max_ya_attempts}). Mengulangi...")
                time.sleep(0.1)

        if not submit_muncul:
            print(f"[WARNING] Modal 'Submit diproses' tetap tidak terdeteksi setelah {max_ya_attempts}x mengetuk 'YA'.")

        # Ketuk OK modal progress
        res_ok = self.ketuk_ok_submit_diproses(max_attempts=10)
        if res_ok == "kembali_blok_iv":
            print("[SUBMIT] Terdeteksi layar kembali ke 'BLOK IV' saat pengetukan OK. Mengembalikan status 'retry_kirim'...")
            return "retry_kirim"

        # Deteksi 'Halaman Upload' secara reaktif (maks 1.5s)
        print("[EMULATOR] Memeriksa apakah teks 'Halaman Upload' muncul di layar...")
        is_halaman_upload = self.smart_wait(
            lambda: (
                self.check_exists(self.d(textContains="Halaman Upload"))
                or self.check_exists(self.d(descriptionContains="Halaman Upload"))
                or self.check_exists(self.d.xpath("//*[contains(@text, 'Halaman Upload') or contains(@content-desc, 'Halaman Upload')]"))
            ),
            timeout=1.5,
            poll_interval=0.05,
        )

        if is_halaman_upload:
            print("[EMULATOR] Teks 'Halaman Upload' terdeteksi! Menekan tombol Back...")
            self.d.press("back")
            time.sleep(0.1)

        # Tunggu kembali ke halaman 'Daftar Assignment'
        sukses_da = self.tunggu_loading("Daftar Assignment", timeout=15)
        if not sukses_da:
            print(f"[WARNING] Teks 'Daftar Assignment' tidak terdeteksi setelah 15s untuk baris {row} (IDPEL: {idpel}).")
            self.kembali_ke_daftar_assignment()
            if row_attempt < 3:
                print(f"[RETRY BARIS] Mengulangi pemrosesan baris {row} ({idpel}) dari awal (percobaan ke-{row_attempt + 1})...")
                return "retry"
            else:
                print(f"[RETRY GAGAL] Sudah mencoba 3x untuk baris {row} ({idpel}). Melanjutkan ke baris berikutnya.")
                return "next"
        else:
            self.excel_manager.simpan_status(row, "SUKSES DIUPDATE")
            return "sukses"
