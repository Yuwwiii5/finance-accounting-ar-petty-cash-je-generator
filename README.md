# Petty Cash JE Generator

Turns a Mews **Payment report** into a ready-to-upload **Business Importer** file for petty cash journal entries in QuickBooks Online — one click per hotel instead of copying lines by hand.

## What it does
- Reads the cash lines from the Mews Payment report (External payments tab) and skips test bookings
- Detects the hotel from the report and picks the matching tab in the Business Importer template
- Looks up each guest's OTA from the Mews Reservation report to build the QuickBooks customer name (`OTA:Guest`)
- Opens a form for guests with no OTA so they can be filled in before generating; answers are reused next time
- Compares the counted payout with the Mews total and books any gap to the reconciliation discrepancy account
- Highlights lines in **yellow** for review: missing accounting category, refunds, lines added by hand, unknown OTA or a discrepancy
- Suggests the next JE number, names the file by journal date and hotel, and copies it to the month's Documentation folder
- Moves used reports and old upload files to the Recycle Bin (recoverable)

## Built with
Python (pandas, openpyxl, tkinter) · Mews · QuickBooks Online · Business Importer · Windows

## Setup
1. Install Python, then run `pip install openpyxl pandas`.
2. Put `petty_cash_bi.py` and the Business Importer template in a folder named `_system`.
3. Put `Run Petty Cash.bat` in the folder above it, where the Mews reports are saved.
4. In `petty_cash_bi.py`, set `DOC_ROOT` and `MASTER` to your own folder paths, and replace the `XXXXXXX` account numbers with your reconciliation discrepancy accounts.

## How to use
1. Export the Payment report and Reservation report from Mews for the same hotel and period, and save them in the main folder.
2. Double-click `Run Petty Cash.bat`.
3. Pick the report, journal date and JE number, enter the counted payout, and click **Generate**.
4. Review any yellow lines, then upload the file to QuickBooks with Business Importer.

> Folder paths and account numbers in this copy are placeholders. No guest or company data is included.
