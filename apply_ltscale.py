"""
Apply LTSCALE 15 to all DWG files in the current folder.

Usage:
    1. Copy this file to the folder containing your DWG files
    2. Run: python apply_ltscale.py
    3. Then run: run_ltscale.bat
"""

import os
import subprocess
from pathlib import Path


def find_accoreconsole():
    """Find AutoCAD Core Console installation."""
    for year in range(2026, 2019, -1):
        path = f"C:\\Program Files\\Autodesk\\AutoCAD {year}\\accoreconsole.exe"
        if os.path.exists(path):
            return path
    return None


def main():
    folder = Path(os.getcwd())

    # Find all DWG files
    dwg_files = list(folder.glob("*.dwg")) + list(folder.glob("*.DWG"))
    dwg_files = sorted(set(dwg_files))

    if not dwg_files:
        print(f"No DWG files found in: {folder}")
        return

    print(f"Found {len(dwg_files)} DWG files")

    # Find AutoCAD Core Console
    accore = find_accoreconsole()
    if not accore:
        print("ERROR: AutoCAD Core Console not found!")
        return

    print(f"Using: {accore}")
    print()

    # Create a single script that sets LTSCALE and saves
    scr_path = folder / "_ltscale_cmd.scr"
    with open(scr_path, "w", encoding="utf-8") as f:
        f.write("_.LTSCALE\n")
        f.write("15\n")
        f.write("_.QSAVE\n")

    # Process each file individually
    success = 0
    failed = 0

    for i, dwg_file in enumerate(dwg_files, 1):
        print(f"[{i}/{len(dwg_files)}] Processing: {dwg_file.name}...", end=" ", flush=True)

        try:
            result = subprocess.run(
                [accore, "/i", str(dwg_file), "/s", str(scr_path)],
                capture_output=True,
                text=True,
                timeout=120  # 2 minute timeout per file
            )

            if "error" in result.stdout.lower() or result.returncode != 0:
                print("FAILED")
                failed += 1
            else:
                print("OK")
                success += 1

        except subprocess.TimeoutExpired:
            print("TIMEOUT")
            failed += 1
        except Exception as e:
            print(f"ERROR: {e}")
            failed += 1

    # Cleanup temp script
    try:
        scr_path.unlink()
    except:
        pass

    print()
    print("=" * 50)
    print(f"Completed: {success} successful, {failed} failed")
    print("=" * 50)


if __name__ == "__main__":
    main()
