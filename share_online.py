"""
Chia sẻ Web ra Internet tự động qua Cloudflare Tunnel.
Tự động lấy URL công khai, sao chép vào bộ nhớ tạm (Clipboard) và hiển thị to rõ ràng.
"""

import subprocess
import re
import sys
import os
import time

def main():
    print("=" * 78)
    print("      DANG KHOI TAO DUONG LINK ONLINE CONG KHAI CHO WEB CUA BAN...")
    print("=" * 78)
    print()

    cloudflared_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cloudflared.exe")
    if not os.path.exists(cloudflared_path):
        cloudflared_path = "cloudflared"

    cmd = [cloudflared_path, "tunnel", "--url", "http://localhost:8501"]

    try:
        proc = subprocess.Popen(
            cmd,
            stderr=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
    except Exception as e:
        print(f"[LOI] Khong the chay cloudflared: {e}")
        input("\nNhan Enter de thoat...")
        return

    public_url = None
    start_time = time.time()

    print("[INFO] Dang ket noi voi mang luoi Cloudflare toan cau (cho khoang 3-5s)...")

    # Đọc log từ stderr để bóc tách URL
    for line in iter(proc.stderr.readline, ''):
        if not line:
            break
        
        # Tìm URL dạng https://xxx.trycloudflare.com
        if "trycloudflare.com" in line:
            m = re.search(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com", line)
            if m:
                public_url = m.group(0)
                break
        
        if time.time() - start_time > 30:
            break

    if public_url:
        # Tự động sao chép URL vào Clipboard trên Windows
        try:
            clip_proc = subprocess.Popen(["clip"], stdin=subprocess.PIPE, shell=True)
            clip_proc.communicate(input=public_url.encode("utf-8"))
            copied_text = " (DA TU DONG COPY VAO CLIPBOARD, BAN CHI CAN CTRL+V DE GUI!)"
        except Exception:
            copied_text = ""

        print()
        print("=" * 78)
        print(" [THANH CONG] DUONG LINK ONLINE CUA BAN DA SAN SANG:")
        print(f" >> {public_url} <<")
        print(f"{copied_text}")
        print("=" * 78)
        print()
        print("* Bat ky ai (dung 4G, dien thoai, may tinh khac) deu co the truy cap vao link nay.")
        print("* Luu y: Giu cua so nay mo de duy tri duong link online. Khi dong cua so nay, link se tat.")
        print()
        print("------------------------------------------------------------------------------")
        print("Nhat ky ket noi (Logs):")

        # Tiếp tục in các log kết nối
        for line in iter(proc.stderr.readline, ''):
            if not line:
                break
            # Chỉ in các thông tin kết nối hữu ích
            if "Connection" in line or "Registered tunnel connection" in line:
                print(">> " + line.strip())
    else:
        print("[CANH BAO] Chua lay duoc link tu dong. Vui long kiem tra xem ung dung Web (run.bat) da bat chua.")
        input("\nNhan Enter de thoat...")

    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()

if __name__ == "__main__":
    main()
