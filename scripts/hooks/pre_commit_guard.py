#!/usr/bin/env python3
"""Hook Claude Code: trước 'git commit' / 'gh pr create' thì chạy pii_scan; dính thì chặn (exit 2)."""
import json
import os
import subprocess
import sys

data = json.load(sys.stdin)
cmd = (data.get("tool_input") or {}).get("command", "")
if "git commit" not in cmd and "gh pr create" not in cmd and "git push" not in cmd:
    sys.exit(0)
root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
r = subprocess.run([sys.executable, os.path.join(root, "scripts", "pii_scan.py")], capture_output=True, text=True, cwd=root)
if r.returncode != 0:
    sys.stderr.write("Chặn: pii_scan phát hiện dữ liệu cá nhân/dữ liệu khách.\n" + r.stdout[-3000:])
    sys.exit(2)
