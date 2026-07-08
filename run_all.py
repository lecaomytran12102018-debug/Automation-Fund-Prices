# -*- coding: utf-8 -*-
import subprocess
import sys
from pathlib import Path

# Danh sách các script cần chạy, ĐÚNG THỨ TỰ mong muốn.
# Có thể comment (thêm #"" trước) dòng nào đó để tạm bỏ qua quỹ đó."
SCRIPTS = [
    "Daiichi.py",
    "Manulife.py",
    "Chubb.py",
    "FWD.py",
    "Hanwha.py",
    "Generali.py",
    "AIA.py",
    "Sunlife.py",
    "PRUlink.py",
]

THIS_DIR = Path(__file__).parent



def run_one(script_name: str) -> bool:
    """Chạy 1 script con, trả về True nếu chạy xong không lỗi (exit code 0)."""
    script_path = THIS_DIR / script_name
    print("\n" + "=" * 70)
    print(f"[*] Đang chạy: {script_name}")
    print("=" * 70)

    if not script_path.exists():
        print(f"[x] Không tìm thấy file '{script_name}' trong thư mục {THIS_DIR}")
        return False

    result = subprocess.run([sys.executable, str(script_path)], cwd=THIS_DIR)
    ok = result.returncode == 0
    if not ok:
        print(f"[x] '{script_name}' kết thúc với lỗi (exit code {result.returncode}).")
    return ok


def main():
    results = {}
    for script_name in SCRIPTS:
        results[script_name] = run_one(script_name)

    print("\n" + "=" * 70)
    print("TỔNG KẾT")
    print("=" * 70)
    for script_name, ok in results.items():
        status = "✅ OK" if ok else "❌ LỖI"
        print(f"  {status:10s} {script_name}")

    n_fail = sum(1 for ok in results.values() if not ok)
    if n_fail:
        print(f"\n[!] Có {n_fail} script bị lỗi, xem log phía trên để biết chi tiết.")
        sys.exit(1)
    else:
        print("\n[+] Tất cả script đã chạy xong, không có lỗi nào.")


if __name__ == "__main__":
    main()