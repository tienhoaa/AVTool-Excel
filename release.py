import os, sys, json, subprocess, zipfile, datetime

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
os.chdir(workspace)

print("=" * 65)
print("     HE THONG PHAT HANH BAN MOI TU DONG 100% - AVTOOL PRO")
print("=" * 65)
print()

# 1. Tu dong doc va tang so phien ban
v_file = os.path.join(workspace, "version.json")
cur_ver = "1.0.2"
if os.path.exists(v_file):
    try:
        with open(v_file, "r", encoding="utf-8") as f:
            data = json.load(f)
            cur_ver = data.get("version", "1.0.2")
    except Exception:
        pass

# Tu dong tang so cuoi: 1.0.2 -> 1.0.3
parts = cur_ver.split(".")
try:
    parts[-1] = str(int(parts[-1]) + 1)
    new_ver = ".".join(parts)
except Exception:
    new_ver = "1.0.3"

now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
print(f"[*] Phien ban cu   : v{cur_ver}")
print(f"[*] PHIEN BAN MOI  : v{new_ver} (Tu dong sinh)")
print()

# 2. Hoi noi dung (hoac an Enter luon de tu sinh)
user_note = input("Nhap ghi chu cap nhat (hoac an Enter de dung mac dinh): ").strip()
if user_note:
    changelog = user_note
else:
    changelog = f"- Cap nhat toi uu he thong ngay {now_str}"

print()
print(f"--> Dang tien hanh phat hanh phien ban v{new_ver}...")
print("-" * 65)

# 3. Ghi version.json
v_data = {
    "version": new_ver,
    "latest_version": new_ver,
    "release_date": datetime.datetime.now().strftime("%Y-%m-%d"),
    "changelog": changelog,
    "download_url": "https://raw.githubusercontent.com/tienhoaa/AVTool-Excel/main/TN_FontViet_Pro_2026.xlam"
}

with open(v_file, "w", encoding="utf-8") as f:
    json.dump(v_data, f, ensure_ascii=False, indent=2)
print("[1/4] Da cap nhat phien ban moi v" + new_ver + " vao version.json.")

# 4. Chay build_master.py
print("[2/4] Dang bien dich lai file TN_FontViet_Pro_2026.xlam...")
res = subprocess.run([sys.executable, "build_master.py"], capture_output=True, text=True)
if res.returncode != 0:
    print("LOI BIEN DICH:")
    print(res.stderr)
    input("\nNhan Enter de thoat...")
    sys.exit(1)
print("      --> Bien dich xong!")

# 5. Dong goi zip ra Desktop
print("[3/4] Dang dong goi file AVTool_Pro_2026.zip...")
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
print("      --> Dong goi ZIP xong!")

# 6. Git Push len GitHub
print("[4/4] Dang day len GitHub tienhoaa/AVTool-Excel...")
git_exe = r"C:\Program Files\Git\cmd\git.exe"
subprocess.run([git_exe, "add", "."], capture_output=True)
subprocess.run([git_exe, "commit", "-m", f"Auto-Release v{new_ver}: {changelog}"], capture_output=True)
push_res = subprocess.run([git_exe, "push", "origin", "main"], capture_output=True, text=True)

print()
if push_res.returncode == 0:
    print("=" * 65)
    print(f" [OK] DA PHAT HANH THANH CONG PHIEN BAN v{new_ver} LEN GITHUB!")
    print(" Nguoi dung chi can bam nut [Cap Nhat Ban Moi] tren Excel la co ngay!")
    print("=" * 65)
else:
    print("LOI PUSH GITHUB:", push_res.stderr)

print()
input("Nhan phim bat ky de dong...")
