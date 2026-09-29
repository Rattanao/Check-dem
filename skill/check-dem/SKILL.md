---
name: check-dem
description: >
  Calculates the demurrage start date (DEM.) and container receipt date (REC.) for every
  B/L in a Freedays sheet (.xls/.xlsx with a "Freedays" column such as "7/5 Add 7/9" or
  "3/3 Add 2/11 TOS 3") from the vessel's ATA date, and produces a color-coded
  "Check DEM.xlsx" report laid out for A4 portrait printing. Use whenever the user asks
  to "check dem", "เช็ค DEM", "หาวันเริ่มเก็บค่าใช้จ่าย", "คำนวณ freedays", or uploads/points
  to a DEM-*.xls file, even if they only give the ATA date or only the file.
---

# Check DEM

Rattana's confirmed report format. Run the bundled script; do not rebuild the layout by hand.

## Rules (already implemented in the script)

- **TOTAL** = first number of Freedays + first number after `Add` → `7/5 Add 7/9` = 7 + 7 = 14
- **DEM.** = ATA + TOTAL days → ATA 27/09/2026 → 2026-10-11
- **REC.** = ATA + 3 days, only when the row contains `TOS` (otherwise blank). The number after TOS is ignored.
- DEM. cells with the same date share a color (gradient red → blue from earliest to latest). The Freedays cell takes the same color as its DEM. cell.

## Steps

1. Get the input file and the ATA date (dd/mm/yyyy). If ATA is missing, ask for it. Never guess.
2. Run:
   ```bash
   PYTHONIOENCODING=utf-8 python "<skill dir>/scripts/check_dem.py" "<input file>" <ATA> "<output dir>"
   ```
   The default output dir is `output/`, next to the input file's folder. The file is always named `Check DEM.xlsx`.
   Requires `pandas`, `xlrd` (for .xls) and `openpyxl`.
3. Read the result back with openpyxl and report the following to the user in Thai:
   - the number of B/Ls
   - each DEM. date with how many B/Ls fall on it
   - which rows have TOS, with their REC. date
   - any row whose DEM. shows `CHECK`, meaning Freedays could not be parsed
4. Send the user the output file.

## SUR form (SUR.xls, in input/)

Source columns: A POL | B Status | C B/L No. | D Freedays | E RECEIPT B/L | F Consignee.
Use only columns **C-F**, and save the report into `input/` as **START_DEM.xlsx**:
```bash
PYTHONIOENCODING=utf-8 python "<skill dir>/scripts/check_dem.py" "input/SUR.xls" <ATA> "input" --cols C-F --name START_DEM
```
`--cols` restricts the source to that column range, and `--name` sets the output file name (".xlsx" is added).

## Report layout

- Row 1 has a navy **CHECK DEM** banner.
- Row 2 is one info line: ต้นฉบับ | file | ATA | date | จำนวน B/L | count.
- Row 4 holds the table header, with columns in this order: No. | Freedays | TOTAL | DEM. | REC. | then the source's other columns (B/L No., RECEIPT B/L, Consignee).
- RECEIPT B/L has a yellow fill with red bold text. TOS rows have red bold text, and their REC. cell has an orange fill.
- Column widths auto-fit the text, including bold/12pt ATA so it never shows `####`. The sheet is set to A4 portrait, 1 page wide, with the header row repeated on every printed page.

If the user asks for layout changes, edit `scripts/check_dem.py`. The canonical copy lives in https://github.com/Rattanao/Check-dem, so keep both copies in sync.
