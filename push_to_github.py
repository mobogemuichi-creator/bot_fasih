"""
Script otomatisasi Commit & Push ke GitHub.
Cara Penggunaan di Antigravity:
  1. Buka file ini (push_to_github.py)
  2. Tekan tombol RUN (▶) di sudut kanan atas IDE Antigravity!
"""
import subprocess
import datetime
import sys

def main():
    print("=" * 60)
    print("       SINKRONISASI & PUSH KE GITHUB       ")
    print("=" * 60)

    # Pesan commit dengan timestamp otomatis
    timestamp = datetime.datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    commit_msg = f"Update perbaikan kode ({timestamp})"

    # Jika user menyertakan pesan custom lewat command line
    if len(sys.argv) > 1:
        commit_msg = " ".join(sys.argv[1:])

    # 1. Git Add
    print(f"\n[1/3] Menambahkan perubahan file (git add .)...")
    res_add = subprocess.run(["git", "add", "."], capture_output=True, text=True)
    if res_add.returncode != 0:
        print(f"[ERROR] Gagal git add:\n{res_add.stderr}")
        return

    # 2. Git Commit
    print(f"[2/3] Membuat commit ('{commit_msg}')...")
    res_commit = subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True, text=True)
    output_commit = res_commit.stdout.strip()
    if output_commit:
        print(output_commit)
    else:
        print("Tidak ada perubahan baru yang perlu di-commit.")

    # 3. Git Push
    print(f"\n[3/3] Mengunggah perubahan ke GitHub (git push)...")
    res_push = subprocess.run(["git", "push"], capture_output=True, text=True)
    if res_push.returncode == 0:
        print("\n" + "=" * 60)
        print(" [SUKSES] Seluruh perubahan berhasil di-push ke GitHub!")
        print(" URL: https://github.com/mobogemuichi-creator/bot_fasih")
        print("=" * 60)
    else:
        print(f"\n[ERROR] Gagal push ke GitHub:\n{res_push.stderr}")

if __name__ == "__main__":
    main()
