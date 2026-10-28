# LCD Catalogue Desktop — V3.1

Offline LCD compatibility, selling-price and local-stock lookup application.

## On Mac

```bash
cd ~/LCD-Catalogue-Desktop
source .venv/bin/activate
python app.py
```

## Updating catalogue data

The application uses `data/lcd_catalogue.json`, a validated snapshot created by the existing n8n `LCD - Catalogue Refresh` workflow. It does **not** parse or watch Excel files directly.

1. Update the three Excel workbooks in the existing n8n input folder.
2. Run the n8n catalogue refresh workflow, validate row counts and selling-price warnings, and produce the new JSON file.
3. Back up the last working `data/lcd_catalogue.json` file.
4. Replace the JSON file with the new one or use **Load Catalogue** to select the new JSON from a separate location.
5. Use **Reload Catalogue** if the same selected file was updated.
6. Verify `SM A06` (17.25; stock 80 in this example snapshot), `Realme 10 5G`, and a known missing-price case before relying on updated quotes. These reference numbers may legitimately change after workbook updates.

**Important:** `Load Catalogue` opens a file and does not copy or merge it into the app's `data` directory. A newly selected path will be forgotten after restarting the app; for lasting updates, replace `data/lcd_catalogue.json`. Changes to Excel alone do not update this app.

## Build on Windows

Run `build_windows.bat` in Windows, with Python 3 installed and reachable through `py -3`.
The portable output is `dist/LCDCatalogueDesktop/`. Deliver the **whole folder**, not just the EXE. The catalogue lives in the `data` subfolder next to the EXE.

Test on an actual Windows PC that does not have Python installed. Check that the whole folder is extracted before launch and that normal user permissions permit reading `data/lcd_catalogue.json`.

## Privacy and portfolio

The included example `data/lcd_catalogue.json` contains business-specific pricing and inventory. Do not push it to any public repository or publish the data in screenshots/demos without the owner's permission. Use a sanitized synthetic sample and keep the real files separately for customer delivery.

## Current limitations

- Desktop interface reads the validated JSON snapshot; Excel refresh still requires the existing n8n process.
- No automatic real-world warehouse verification.
- Original 4,202 vs 3,872 compatibility-count discrepancy remains unresolved.
- Results can require manual variant checks; tests cover representative cases, not all records.
