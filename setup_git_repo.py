import subprocess, os, json

workspace = r"C:\Users\GAMING-PC\Downloads\TN-FONTVIET-R8"
git_exe = r"C:\Program Files\Git\cmd\git.exe"
gh_exe = r"C:\Program Files\GitHub CLI\gh.exe"

# 1. Config git identity
subprocess.run([git_exe, "config", "--global", "user.name", "tienhoaa"], cwd=workspace)
subprocess.run([git_exe, "config", "--global", "user.email", "tienhoaa@users.noreply.github.com"], cwd=workspace)
subprocess.run([git_exe, "config", "user.name", "tienhoaa"], cwd=workspace)
subprocess.run([git_exe, "config", "user.email", "tienhoaa@users.noreply.github.com"], cwd=workspace)

# 2. Get token from gh
token = subprocess.run([gh_exe, "auth", "token"], capture_output=True, text=True).stdout.strip()
remote_url = f"https://tienhoaa:{token}@github.com/tienhoaa/AVTool-Excel.git"

# 3. Add and commit
subprocess.run([git_exe, "add", "-A"], cwd=workspace)
res = subprocess.run([git_exe, "commit", "-m", "Release AVTool Pro v1.0.1 with DValue and Auto-Update"], cwd=workspace, capture_output=True, text=True)
print("Git commit output:\n", res.stdout, res.stderr)

# 4. Set remote
subprocess.run([git_exe, "remote", "remove", "origin"], cwd=workspace)
subprocess.run([git_exe, "remote", "add", "origin", remote_url], cwd=workspace)
subprocess.run([git_exe, "branch", "-M", "main"], cwd=workspace)

# 5. Push
push_res = subprocess.run([git_exe, "push", "-u", "origin", "main", "--force"], cwd=workspace, capture_output=True, text=True)
print("Git push output:\n", push_res.stdout, push_res.stderr)

if push_res.returncode == 0:
    print("SUCCESS: Code pushed to https://github.com/tienhoaa/AVTool-Excel!")
else:
    print("Push failed:", push_res.returncode)
