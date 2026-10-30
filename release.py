import os, sys, json, subprocess, zipfile, datetime, re

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
os.chdir(workspace)

print("=" * 65)
print("     HE THONG PHAT HANH BAN MOI TU DONG 100% - AVTOOL PRO")
print("=" * 65)
print()

# 1. Tu dong doc phien ban hien tai tu version.json
v_file = os.path.join(workspace, "version.json")
cur_ver = "1.0.4"
if os.path.exists(v_file):
    try:
        with open(v_file, "r", encoding="utf-8") as f:
            data = json.load(f)
            cur_ver = data.get("version", "1.0.4")
    except Exception:
        pass

# 2. Tinh toan phien ban goi y
parts = cur_ver.split(".")
try:
    patch_parts = list(parts)
    patch_parts[-1] = str(int(patch_parts[-1]) + 1)
    ver_patch = ".".join(patch_parts)
except Exception:
    ver_patch = "1.0.5"

try:
    major_num = int(parts[0]) + 1
    ver_major = f"{major_num}.0.0"
except Exception:
    ver_major = "2.0.0"

print(f"[*] Phien ban hien tai : v{cur_ver}")
print("-" * 65)
print(f"  [1] Cap nhat nho / Sua loi  --> v{ver_patch} (Mac dinh - chi can go Enter)")
print(f"  [2] Them tinh nang moi       --> v{ver_major}")
print(f"  [3] Tu nhap so phien ban khac")
print("-" * 65)

choice = input("Chon loai phat hanh [1/2/3] (Enter = 1): ").strip()
if choice == "2":
    new_ver = ver_major
elif choice == "3":
    custom_v = input("Nhap so phien ban mong muon (vi du 2.0.0): ").strip()
    new_ver = custom_v if custom_v else ver_patch
else:
    new_ver = ver_patch

now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
print()
print(f"--> PHIEN BAN SE PHAT HANH: v{new_ver}")
print()

# 3. Hoi noi dung ghi chu cap nhat
user_note = input("Nhap ghi chu cap nhat (hoac an Enter de dung mac dinh): ").strip()
if user_note:
    changelog = user_note
else:
    changelog = f"- Cap nhat toi uu he thong ngay {now_str}"

print()
print(f"--> Dang tien hanh phat hanh phien ban v{new_ver}...")
print("-" * 65)

# 4. Ghi version.json
v_data = {
    "version": new_ver,
    "latest_version": new_ver,
    "release_date": datetime.datetime.now().strftime("%Y-%m-%d"),
    "changelog": changelog,
    "download_url": "https://raw.githubusercontent.com/tienhoaa/AVTool-Excel/main/TN_FontViet_Pro_2026.xlam"
}

with open(v_file, "w", encoding="utf-8") as f:
    json.dump(v_data, f, ensure_ascii=False, indent=2)
print("[1/5] Da cap nhat phien ban moi v" + new_ver + " vao version.json.")

# 5. Cap nhat currentVer trong Mod_FontViet_Core.bas de Addin dong bo phien ban
bas_core = os.path.join(workspace, "Mod_FontViet_Core.bas")
if os.path.exists(bas_core):
    with open(bas_core, "r", encoding="utf-8") as bf:
        bcode = bf.read()
    bcode = re.sub(r'currentVer = "[0-9\.]+"', f'currentVer = "{new_ver}"', bcode)
    with open(bas_core, "w", encoding="utf-8") as bf:
        bf.write(bcode)
    print(f"[2/5] Da cap nhat currentVer = \"{new_ver}\" vao Mod_FontViet_Core.bas")

# 6. Chay build_master.py de bien dich xlam
print("[3/5] Dang bien dich lai file TN_FontViet_Pro_2026.xlam...")
res = subprocess.run([sys.executable, "build_master.py"], capture_output=True, text=True)
if res.returncode != 0:
    print("LOI BIEN DICH:")
    print(res.stderr)
    input("\nNhan Enter de thoat...")
    sys.exit(1)
print("      --> Bien dich xong!")

# 7. Dong goi zip ra Desktop va Workspace
print("[4/5] Dang dong goi file AVTool_Pro_2026.zip...")
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
print("      --> Dong goi ZIP xong (da xuat ra Desktop)!")

# 8. Git Push len GitHub
print("[5/5] Dang day len GitHub tienhoaa/AVTool-Excel...")
git_exe = r"C:\Program Files\Git\cmd\git.exe"
subprocess.run([git_exe, "add", "."], capture_output=True)
subprocess.run([git_exe, "commit", "-m", f"Release v{new_ver}: {changelog}"], capture_output=True)
push_res = subprocess.run([git_exe, "push", "origin", "main"], capture_output=True, text=True)

print()
if push_res.returncode == 0:
    print("=" * 65)
    print(f" [OK] DA PHAT HANH THANH CONG PHIEN BAN v{new_ver} LEN GITHUB!")
    print("=" * 65)
    print("Moi nguoi dung mo Excel hoac bam Cap nhat se tu dong nhan ban moi nay!")
else:
    print("LOI PUSH GITHUB:", push_res.stderr)

print()
input("Nhan Enter de hoan tat va dong...")
