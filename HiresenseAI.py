#!/usr/bin/env python3
"""
Master script to run the recruitment pipeline sequentially:
1. AIresumereader.py – extracts and scores candidates
2. calling.py – calls shortlisted candidates
3. emailautomation.py – sends emails to shortlisted/non‑shortlisted
"""

import subprocess
import sys
import os

SCRIPTS = [
    "AIresumereader.py",
    "Risk.py",
    "calling.py",
    "emailautomation.py"
]

def run_script(script_name):
    """Execute a Python script and return success status."""
    print(f"\n{'='*60}")
    print(f" Running: {script_name}")
    print(f"{'='*60}\n")

    python_executable = sys.executable
    try:
        result = subprocess.run(
            [python_executable, script_name],
            capture_output=True,
            text=True,
            check=False
        )
        if result.stdout:
            print("--- STDOUT ---")
            print(result.stdout)
        if result.stderr:
            print("--- STDERR ---")
            print(result.stderr)

        if result.returncode == 0:
            print(f"\n {script_name} completed successfully.\n")
            return True
        else:
            print(f"\n {script_name} failed with exit code {result.returncode}.\n")
            return False
    except FileNotFoundError:
        print(f"\n Script not found: {script_name}\n")
        return False
    except Exception as e:
        print(f"\n Unexpected error running {script_name}: {e}\n")
        return False

def main():
    missing = [s for s in SCRIPTS if not os.path.isfile(s)]
    if missing:
        print("The following scripts are missing:")
        for s in missing:
            print(f"  - {s}")
        sys.exit(1)

    for script in SCRIPTS:
        if not run_script(script):
            print("Pipeline stopped due to failure.")
            sys.exit(1)

    print("\nAll scripts completed successfully.")

if __name__ == "__main__":
    main()