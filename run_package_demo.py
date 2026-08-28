"""
End-to-end demonstration of the packaging engine.

Runs both packagers against the synthetic sample data and prints
the output file paths. Useful for portfolio demos and CI smoke tests.
"""

import os
import subprocess
import sys


def main():
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    os.chdir(repo_root)

    print("=" * 70)
    print("Federal FA Packaging Engine — End-to-End Demo")
    print("=" * 70)

    print("\n[1/2] Generating clock-hour packet (MSMA sample)...")
    result = subprocess.run(
        [sys.executable, "generate_package.py"],
        capture_output=True, text=True,
    )
    print(result.stdout)
    if result.returncode != 0:
        print(f"ERROR: {result.stderr}")
        return 1

    print("\n[2/2] Generating quarter-program packet (Culinary Diploma sample)...")
    result = subprocess.run(
        [sys.executable, "generate_package_quarter.py"],
        capture_output=True, text=True,
    )
    print(result.stdout)
    if result.returncode != 0:
        print(f"ERROR: {result.stderr}")
        return 1

    print("\n" + "=" * 70)
    print("Demo complete. Output PDFs are in ./output/FA_Info_Packets/")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    sys.exit(main())