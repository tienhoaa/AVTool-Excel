import os, sys, json, subprocess, zipfile

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
os.chdir(workspace)

print("=" * 65)
print("     HE THONG PHAT HANH PHIEN BAN MOI CHO AVTOOL PRO")
print("=" * 65)
print()

# 1. Doc phien ban hien tai
v_file = os.path.join(workspace, "version.json")
cur_ver = "1.0.2"
if os.path.exists(v_file):
    try:
        with open(v_file, "r", encoding="utf-8") as f:
            data = json.load(f)
            cur_ver = data.get("version", "1.0.2")
    except Exception:
        pass

print(f"[*] Phien ban hien tai tren GitHub: v{cur_ver}")
print()

new_ver = input(f"Nhap so phien ban moi (VD: 1.0.3, Enter de bo qua): ").strip()
if not new_ver:
    # Auto increment last digit
    parts = cur_ver.split(".")
    try:
        parts[-1] = str(int(parts[-1]) + 1)
        new_ver = ".".join(parts)
    except Exception:
        new_ver = "1.0.3"
    print(f"--> Tu dong chon phien ban ke tiep: v{new_ver}")

changelog = input("Nhap noi dung tinh nang moi (VD: Nang cap DValue, sua loi in an...): ").strip()
if not changelog:
    changelog = "- Cap nhat toi uu he thong va sua loi on dinh"

print()
print(f"[+] Dang phat hanh phien ban v{new_ver}...")
print(f"[+] Noi dung: {changelog}")
print("-" * 65)

# 2. Cap nhat version.json
v_data = {
    "version": new_ver,
    "latest_version": new_ver,
    "release_date": subprocess.getoutput("powershell Get-Date -Format 'yyyy-MM-dd'").strip(),
    "changelog": changelog,
    "download_url": "https://raw.githubusercontent.com/tienhoaa/AVTool-Excel/main/TN_FontViet_Pro_2026.xlam"
}

with open(v_file, "w", encoding="utf-8") as f:
    json.dump(v_data, f, ensure_ascii=False, indent=2)
print("[1/4] Da cap nhat file version.json.")

# 3. Chay build_master.py
print("[2/4] Dang bien dich file TN_FontViet_Pro_2026.xlam...")
res = subprocess.run([sys.executable, "build_master.py"], capture_output=True, text=True)
if res.returncode != 0:
    print("LOI BIEN DICH:")
    print(res.stderr)
    input("\nNhan Enter de thoat...")
    sys.exit(1)
print("      --> Bien dich thanh cong!")

# 4. Dong goi zip ra Desktop
print("[3/4] Dang dong goi file AVTool_Pro_2026.zip ra Desktop...")
desktop = os.path.expanduser("~/Desktop")
zip_desktop = os.path.join(desktop, "AVTool_Pro_2026.zip")
zip_local = os.path.join(workspace, "AVTool_Pro_2026.zip")

files_to_pack = [
    "TN_FontViet_Pro_2026.xlam",
    "Cai_Dat_AVTool.bat",
    "Cap_Nhat_Online.bat",
    "Go_Cai_Dat_AVTool.bat",
    "Huong_Dan_Su_Dung.txt"
]

for zpath in [zip_desktop, zip_local]:
    with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for f in files_to_pack:
            if os.path.exists(f):
                zf.write(f, f)
print("      --> Da dong goi ZIP thanh cong!")

# 5. Git Commit & Push len GitHub
print("[4/4] Dang day ban moi len GitHub tienhoaa/AVTool-Excel...")
git_exe = r"C:\Program Files\Git\cmd\git.exe"
subprocess.run([git_exe, "add", "."], capture_output=True)
subprocess.run([git_exe, "commit", "-m", f"Release v{new_ver}: {changelog}"], capture_output=True)
push_res = subprocess.run([git_exe, "push", "origin", "main"], capture_output=True, text=True)

if push_res.returncode == 0:
    print("      --> Da day len GitHub thanh cong 100%!")
    print()
    print("=" * 65)
    print(f" [OK] DA PHAT HANH THANH CONG PHIEN BAN v{new_ver} LEN GITHUB!")
    print(" Tu bay gio, moi nguoi dung chi can bam nut [Cap Nhat Ban Moi]")
    print(" tren thanh Ribbon Excel la se nhan duoc ban moi nay!")
    print("=" * 65)
else:
    print("LOI KHI PUSH LEN GITHUB:")
    print(push_res.stderr)

print()
input("Nhan phim bat ky de ket thuc...")
