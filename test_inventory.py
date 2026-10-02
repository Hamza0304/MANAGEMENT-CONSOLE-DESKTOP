import subprocess
from pathlib import Path

SCRIPT_PATH = Path(__file__).parent / "powershell" / "inventory.ps1"

print("PowerShell script:")
print(SCRIPT_PATH)
print()


result = subprocess.run(
    [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPT_PATH),
    ],
    capture_output=True,
    text=True,
    encoding="utf-8",
    errors="replace",
)


print("RETURN CODE:", result.returncode)

print("\n--- POWERSHELL OUTPUT ---")
print(result.stdout)

print("\n--- POWERSHELL ERRORS ---")
print(result.stderr)