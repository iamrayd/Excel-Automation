"""
Phase 3: Node will spawn this script as a subprocess, passing:
  --source <path to the FROM .xlsx/.xlsm>
  --code   <path to a .txt/.bas file containing the final VBA code>
  --mode   overwrite | save-as
  --output <path to write the result to, when mode=save-as>

This script will use pywin32 to:
  1. Launch Excel (win32com.client.Dispatch("Excel.Application"))
  2. Open the source workbook
  3. Add a VBA module via workbook.VBProject.VBComponents.Add(1)  # 1 = standard module
  4. Write the code into that module's CodeModule
  5. Save as .xlsm (overwrite in place, or save-as a new path)
  6. Quit Excel cleanly (including on error, via try/finally)

Requires (one-time setup on this machine):
  - pip install pywin32
  - Excel Trust Center > Macro Settings > "Trust access to the VBA project object model" enabled

Not implemented yet - this is a placeholder for Phase 3.
"""

import argparse
import sys


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--code", required=True)
    parser.add_argument("--mode", choices=["overwrite", "save-as"], default="save-as")
    parser.add_argument("--output", required=False)
    args = parser.parse_args()

    print("inject_vba.py is a placeholder - Phase 3 will implement this.", file=sys.stderr)
    sys.exit(1)


if __name__ == "__main__":
    main()
