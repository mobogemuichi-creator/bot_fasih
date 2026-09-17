import os
import time
import subprocess
import uiautomator2 as u2
from konfigurasi import (
    LDPLAYER_ADB,
    EMULATOR_PORTS_1 as EMULATOR_PORTS,
    SLEEP_SHORT,
)


class FasihDriver:
    """Mengelola koneksi UIAutomator2 ke emulator Android dan menyediakan fungsi gesture layar aman."""

    def __init__(self, ports=None, adb_path=None):
        self.ports = ports or EMULATOR_PORTS
        self.adb_path = adb_path or LDPLAYER_ADB
        self.d = None
        self.screen_width = 1080
        self.screen_height = 1920

    @property
    def device(self):
        """Property alias untuk self.d"""
        return self.d

    def hubungkan_emulator(self, ports=None, adb_path=None) -> bool:
        """Menghubungkan ke emulator via uiautomator2 secara cepat dan mendeteksi dimensi layar."""
        if ports is not None:
            self.ports = ports
        if adb_path is not None:
            self.adb_path = adb_path

        print("[KONEKSI] Mencoba menghubungkan ke emulator LDPlayer...")

        for port in self.ports:
            try:
                if os.path.exists(self.adb_path):
                    subprocess.run(
                        [self.adb_path, "connect", f"127.0.0.1:{port}"],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        timeout=2,
                    )
                temp_d = u2.connect(f"127.0.0.1:{port}")
                info = temp_d.info
                self.d = temp_d
                self.screen_width = info.get("displayWidth", 1080)
                self.screen_height = info.get("displayHeight", 1920)
                print(f"[KONEKSI] Berhasil terhubung ke emulator di port {port}!")
                print(f"[EMULATOR] Layar: {self.screen_width}x{self.screen_height} ({info.get('productName', 'Android')})")
                return True
            except Exception:
                continue

        # Fallback default connection jika adb sudah auto-detect
        try:
            temp_d = u2.connect()
            info = temp_d.info
            self.d = temp_d
            self.screen_width = info.get("displayWidth", 1080)
            self.screen_height = info.get("displayHeight", 1920)
            print("[KONEKSI] Berhasil terhubung ke emulator via default connection!")
            return True
        except Exception:
            pass

        print("[ERROR] Gagal terhubung ke emulator. Pastikan LDPlayer sudah berjalan.")
        return False

    @staticmethod
    def check_exists(el) -> bool:
        """Helper universal untuk memeriksa keberadaan elemen Selector maupun XPath uiautomator2."""
        if el is None:
            return False
        try:
            if hasattr(el, "exists"):
                return bool(el.exists)
            elif hasattr(el, "wait"):
                return bool(el.wait(exists=True, timeout=0.1))
        except Exception:
            pass
        return False

    def swipe_aman(self, fx: int, fy: int, tx: int, ty: int, duration: float = 0.1):
        """Melakukan swipe dengan koordinat aman di dalam batas layar."""
        if not self.d:
            return
        fx = max(10, min(self.screen_width - 10, fx))
        tx = max(10, min(self.screen_width - 10, tx))
        fy = max(10, min(self.screen_height - 10, fy))
        ty = max(10, min(self.screen_height - 10, ty))
        try:
            self.d.swipe(fx, fy, tx, ty, duration=duration)
        except Exception as e:
            print(f"[WARNING] Gagal melakukan swipe aman ({fx},{fy} -> {tx},{ty}): {e}")

    def loop_swipe_statis(self, delta_y: int = 700, loop: int = 3, tengah_layar: tuple = None, duration: float = 0.05):
        """
        Melakukan swipe berulang (statis) sejumlah 'loop' kali.
        delta_y > 0 -> swipe ke bawah (scroll up / melihat konten atas)
        delta_y < 0 -> swipe ke atas (scroll down / melihat konten bawah)
        """
        if not self.d:
            return
        cx = tengah_layar[0] if tengah_layar else self.screen_width // 2
        cy = tengah_layar[1] if tengah_layar else self.screen_height // 2

        for _ in range(loop):
            sy = cy - (delta_y // 2)
            ey = cy + (delta_y // 2)
            self.swipe_aman(cx, sy, cx, ey, duration=duration)
            time.sleep(0.05)

    def loop_swipe_dinamis(self, tengah_layar: tuple = None, delta_y: int = -700, target_text: str = None, max_retry: int = 15, duration: float = 0.05) -> bool:
        """
        Melakukan swipe dinamis sampai elemen dengan target_text terlihat di layar.
        Returns True jika ditemukan, False jika melewati batas max_retry.
        """
        if not self.d:
            return False
        cx = tengah_layar[0] if tengah_layar else self.screen_width // 2
        cy = tengah_layar[1] if tengah_layar else self.screen_height // 2

        if target_text and (self.d(textContains=target_text).exists() or self.d(descriptionContains=target_text).exists()):
            print(f"[SWIPE] Text target '{target_text}' sudah terlihat di layar sebelum swipe.")
            return True

        for attempt in range(1, max_retry + 1):
            sy = cy - (delta_y // 2)
            ey = cy + (delta_y // 2)
            self.swipe_aman(cx, sy, cx, ey, duration=duration)
            time.sleep(0.05)

            if target_text:
                if self.d(textContains=target_text).exists() or self.d(descriptionContains=target_text).exists():
                    print(f"[SWIPE] Text target '{target_text}' terdeteksi di layar pada percobaan ke-{attempt}!")
                    return True

        if target_text:
            print(f"[WARNING] Target text '{target_text}' tidak terdeteksi setelah {max_retry}x swipe.")
            return False
        return True
