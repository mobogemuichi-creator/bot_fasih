"""
================================================================================
BOT EMULATOR REJECT OOP - UJI COBA OPTIMASI KECEPATAN (SMART WAIT)
================================================================================
Runner uji coba khusus optimasi kecepatan bot emulator reject FASIH dengan
arsitektur Object-Oriented Programming (OOP) dan Smart Wait dinamis.

Fitur Utama:
1. Smart Wait (Polling reaktif 40ms - 50ms menggantikan time.sleep statis).
2. Pemangkasan delay transisi halaman dan dialog.
3. Stopwatch / Benchmark per baris IDPEL dengan pelaporan statistik kecepatan.
4. Preservasi 100% logika validasi: Cek NIK, Shortcut GALAT 0, Error 105/106.
================================================================================
"""

import sys
import time
from datetime import datetime

from konfigurasi import (
    LDPLAYER_ADB,
    EXCEL_FILE_REJECT as EXCEL_FILE,
    EMULATOR_PORTS_1 as EMULATOR_PORTS,
    SLEEP_SHORT,
    SLEEP_MEDIUM,
)

try:
    from konfigurasi import PAUSE_ON_GALAT_0
except ImportError:
    PAUSE_ON_GALAT_0 = False

from bot_reject import (
    FasihDriver,
    ExcelManager,
    BasePage,
    AssignmentPage,
    GalatFixer,
    SubmitService,
    KuesionerPage,
)


def proses_update_reject_nik_fast(
    driver: FasihDriver,
    excel_mgr: ExcelManager,
    assign_page: AssignmentPage,
    kuesioner_page: KuesionerPage,
    submit_service: SubmitService,
):
    """
    Orchestrator cepat dengan benchmarking waktu per data.
    """
    data_reject = excel_mgr.baca_data_reject()
    if not data_reject:
        print("[HALT] Tidak ada data reject yang dapat diproses.")
        return

    total = len(data_reject)
    durasi_list = []

    print(f"\n==================================================")
    print(f"   MEMULAI UJI COBA OPTIMASI CEPAT ({total} DATA)   ")
    print(f"==================================================")

    for idx, item in enumerate(data_reject, start=1):
        idpel = item["idpel"]
        nometer = item.get("nometer", "")
        nik = item["nik"]
        nama = item["nama"]
        row = item["row"]

        # Stopwatch per baris
        t_start_row = time.time()

        # 1. Validasi NAMA: Hanya alfabet dan spasi
        if nama and not all(c.isalpha() or c.isspace() for c in nama):
            print(f"[SKIP] Baris {row} | IDPEL {idpel} dilewati karena NAMA '{nama}' non-alfabet.")
            excel_mgr.simpan_status(row, "Error : nama tidak murni alphabeth")
            continue

        # 2. Loop Percobaan Baris (Maksimal 3x)
        sukses_baris = False
        for row_attempt in range(1, 4):
            print(f"\n--------------------------------------------------")
            print(f"[{idx}/{total}] [TURBO] Baris {row} (Attempt {row_attempt}/3) | IDPEL: {idpel} | NIK: {nik}")
            print(f"--------------------------------------------------")

            # 2.1 Swipe ke atas untuk memastikan posisi paling atas
            driver.loop_swipe_statis(delta_y=-700, loop=1)
            driver.loop_swipe_statis(delta_y=800, loop=1)
            time.sleep(0.05)

            # 2.2 Tunggu elemen Search dengan smart wait
            if not assign_page.tunggu_loading("Search", timeout=15):
                print("[HALT] Halaman 'Search' tidak ditemukan setelah 15 detik. Mengakhiri proses.")
                return

            # 2.3 Input IDPEL ke search box
            sukses_search = assign_page.cari_dan_ketuk_search_box(idpel)
            if not sukses_search:
                print(f"[SKIP] Gagal di search box IDPEL {idpel}. Mengulangi...")
                continue

            # 2.4 Cek 'No matching records found'
            if assign_page.cek_no_matching_records():
                print(f"[NOT FOUND] Terdeteksi 'No matching records' untuk IDPEL {idpel}.")
                excel_mgr.simpan_status(row, "DATA TIDAK DITEMUKAN")
                sukses_baris = True
                break

            # 2.5 Buka Detail Kuesioner secara reaktif
            sukses_buka, status_awal, is_still_da = assign_page.buka_detail_kuesioner(nometer, idpel)
            if status_awal == "SUBMITTED":
                print(f"[STATUS] Terdeteksi status 'SUBMITTED' dari awal pada IDPEL {idpel}.")
                excel_mgr.simpan_status(row, "SUKSES DARI AWAL SUDAH SUBMIT")
                sukses_baris = True
                break

            if not sukses_buka:
                if is_still_da:
                    print(f"[RETRY] Masih di Daftar Assignment untuk IDPEL {idpel}. Mengulangi...")
                    continue
                else:
                    print(f"[TIMEOUT] Form kuesioner gagal terbuka untuk IDPEL {idpel}.")
                    excel_mgr.simpan_status(row, "Error: Timeout Mulai Wawancara")
                    assign_page.kembali_ke_daftar_assignment()
                    sukses_baris = True
                    break

            # 2.6 BLOK I & Validasi Awal (Smart Wait & auto-fix 105/106)
            res_v1 = kuesioner_page.proses_blok_i_dan_validasi_awal(idpel, row, row_attempt)
            if res_v1 == "selesai":
                sukses_baris = True
                break
            elif res_v1 == "stop":
                print(f"[HALT] Bot dihentikan manual oleh pengguna pada baris {row} (IDPEL: {idpel}).")
                return
            elif res_v1 == "retry":
                continue
            elif res_v1 == "skip":
                sukses_baris = True
                break

            # 2.7 BLOK II (Nama, NIK, Cek NIK, Shortcut GALAT 0, No Telp, Radio)
            res_b2 = kuesioner_page.proses_blok_ii(nama, nik, idpel, row, row_attempt)
            if res_b2 == "selesai":
                sukses_baris = True
                break
            elif res_b2 == "stop":
                print(f"[HALT] Bot dihentikan manual oleh pengguna pada baris {row} (IDPEL: {idpel}).")
                return
            elif res_b2 == "retry":
                continue
            elif res_b2 == "skip":
                sukses_baris = True
                break

            # 2.8 BLOK III (Alamat saat ini & Increment)
            kuesioner_page.proses_blok_iii()

            # 2.9 BLOK IV (Catatan & Submisi Final)
            res_b4 = kuesioner_page.proses_blok_iv(row, idpel, row_attempt)
            if res_b4 == "stop":
                print(f"[HALT] Bot dihentikan manual oleh pengguna pada baris {row} (IDPEL: {idpel}).")
                return
            elif res_b4 == "retry":
                continue
            else:
                sukses_baris = True
                break

        # Hitung waktu pemrosesan per baris
        durasi_row = time.time() - t_start_row
        durasi_list.append(durasi_row)
        print(f"[BENCHMARK] Baris {row} selesai dalam {durasi_row:.2f} detik (Rata-rata: {sum(durasi_list)/len(durasi_list):.2f}s/data).")

    # Ringkasan Statistik Performa
    if durasi_list:
        avg_speed = sum(durasi_list) / len(durasi_list)
        total_time = sum(durasi_list)
        print("\n==================================================")
        print("          RINGKASAN BENCHMARK KECEPATAN           ")
        print("==================================================")
        print(f" Total Data Diproses : {len(durasi_list)}")
        print(f" Total Waktu Eksekusi: {total_time:.2f} detik")
        print(f" Kecepatan Rata-rata : {avg_speed:.2f} detik / data")
        print(f" Waktu Tercepat      : {min(durasi_list):.2f} detik")
        print(f" Waktu Terlambat     : {max(durasi_list):.2f} detik")
        print("==================================================")


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

    print("==================================================")
    print("      BOT EMULATOR REJECT - FAST OOP RUNNER       ")
    print("      (OPTIMASI KECEPATAN & SMART WAIT POLAR)     ")
    print("==================================================")

    # Inisialisasi Driver
    driver = FasihDriver()
    if not driver.hubungkan_emulator(EMULATOR_PORTS, LDPLAYER_ADB):
        return

    # Inisialisasi Page Objects & Services
    excel_mgr = ExcelManager(file_path=EXCEL_FILE)
    assign_page = AssignmentPage(driver)
    galat_fixer = GalatFixer(driver)
    submit_service = SubmitService(driver, excel_mgr, pause_on_galat_0=PAUSE_ON_GALAT_0)
    kuesioner_page = KuesionerPage(driver, galat_fixer, submit_service, excel_mgr)

    # Eksekusi Runner Cepat
    proses_update_reject_nik_fast(driver, excel_mgr, assign_page, kuesioner_page, submit_service)

    print("\n==================================================")
    print("  UJI COBA OPTIMASI BOT REJECT SELESAI DENGAN SUKSES")
    print("==================================================")


if __name__ == "__main__":
    main()
