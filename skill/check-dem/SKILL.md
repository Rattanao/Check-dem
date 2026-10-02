---
name: check-dem
description: >
  Calculates the demurrage start date (DEM.) and container receipt date (REC.) for every
  B/L in a Freedays sheet (.xls/.xlsx with a "Freedays" column such as "7/5 Add 7/9" or
  "3/3 Add 2/11 TOS 3") from the vessel's ATA date, and produces a color-coded
  "CHECK DEM.xlsx" report laid out for A4 portrait printing. Source forms vary; the needed
  headings (Freedays, TPSZ, B/L No., RECEIPT B/L, Consignee) are found by name. Use whenever the user asks
  to "check dem", "เช็ค DEM", "หาวันเริ่มเก็บค่าใช้จ่าย", "คำนวณ freedays", or uploads/points
  to a DEM-*.xls file, even if they only give the ATA date or only the file.
---

# Check DEM

Rattana's confirmed report format. Run the bundled script; do not rebuild the layout by hand.

## Rules (already implemented in the script)

- **TOTAL** = first number of Freedays + first number after `Add` → `7/5 Add 7/9` = 7 + 7 = 14
- **DEM.** = ATA + TOTAL days → ATA 27/09/2026 → 11/10/2026 (shown as DD/MM/YYYY)
- **REC.** = ATA + 3 days, when the row contains `TOS` **or the TPSZ starts with R** (reefer, e.g. `R40Hx1`). Otherwise blank. The number after TOS is ignored.
- A reefer container is always treated as TOS 3, because it always has an electricity charge. This applies even when Freedays does not say TOS.
- DEM. cells with the same date share a color (gradient red → blue from earliest to latest). The Freedays cell takes the same color as its DEM. cell.

## Steps

1. Get the input file and the ATA date (dd/mm/yyyy). If ATA is missing, ask for it. Never guess.
2. Run:
   ```bash
   PYTHONIOENCODING=utf-8 python "<skill dir>/scripts/check_dem.py" "<input file>" <ATA> "<output dir>"
   ```
   The default output dir is `output/`, next to the input file's folder. Rattana wants the report in `output/`, so do not pass an output dir. The file is named `CHECK DEM.xlsx` (`--name` changes it).
   Requires `pandas`, `xlrd` (for .xls) and `openpyxl`.
3. Read the result back with openpyxl and report the following to the user in Thai:
   - the number of B/Ls
   - each DEM. date with how many B/Ls fall on it
   - which rows have a REC. date, split into rows with TOS in Freedays and reefer rows (TPSZ starting with R)
   - any row whose DEM. shows `CHECK`, meaning Freedays could not be parsed
4. Send the user the output file.

## Source forms (headings found by name)

The Excel source comes in several layouts (for example `SUR.xls` with POL | TPSZ | STATUS | B/L No. | Freedays | RECEIPT B/L | Consignee). The script does not depend on column positions:

- The header row is the first row that contains a `Freedays` heading.
- Headings used, matched by name in any column order: **Freedays** (required), **TPSZ**, **B/L No.**, **RECEIPT B/L**, **Consignee**. A heading that is missing from the source is left out of the report.
- Every other heading (POL, STATUS, ...) is ignored.
- The report always comes out in the same form, whatever the source looks like.

```bash
PYTHONIOENCODING=utf-8 python "<skill dir>/scripts/check_dem.py" "input/SUR.xls" <ATA>
```

To add a heading, add it to `FIELDS` in `scripts/check_dem.py` and in `index.html`.

## Report layout

- Row 1 has a navy **CHECK DEM** banner.
- Row 2 is one info line on light blue: จำนวน B/L + count in the first two columns, ต้นฉบับ + file name under Freedays and DEM., ATA + date in the last two columns. The cells between are empty.
- Row 4 holds the table header, always in this order: No. | TPSZ | REC. | TOTAL | Freedays | DEM. | B/L No. | RECEIPT B/L | Consignee.
- RECEIPT B/L has a yellow fill with red bold text. TOS and reefer rows have red bold Freedays text, and their REC. cell has an orange fill.
- Column widths auto-fit the text, including bold/12pt ATA so it never shows `####`. The sheet is set to A4 portrait, 1 page wide, with the header row repeated on every printed page.

If the user asks for layout changes, edit `scripts/check_dem.py`. The canonical copy lives in https://github.com/Rattanao/Check-dem, so keep both copies in sync.
