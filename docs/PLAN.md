# Plan

Updated: 2026-04-27

## Current Turn

1. Fix the NAV-per-share source for cross listings in the current PRINOS/class-sheet logic.
2. Rerun the current input set and validate the corrected cross-listing outputs.
3. Record the new output location and validation outcome.

## Status

- Completed: Identified that the repo is a single-file Streamlit app with entry point `Tradeweb.py`.
- Completed: Confirmed expected dependencies from `requirements.txt`: `streamlit`, `pandas`, `openpyxl`, `lxml`.
- Completed: Confirmed the app expects:
  - one PRINOS XML file
  - one or more PORTFELJ XML files
  - optional `ADDITIONAL_DATA.xlsx`
  - optional `template.xlsx`
- Completed: Confirmed the core processing logic lives in `generate_outputs(...)` inside `Tradeweb.py`.
- Completed: Installed a local Python 3.12 runtime under the user profile and installed project dependencies.
- Completed: Refactored `Tradeweb.py` so the Streamlit UI runs through `run_streamlit_app()` and the processing logic can be imported safely.
- Completed: Added `run_local.py` to run the generator directly against the `input` folder and write outputs to `output`.
- Completed: Diagnosed that the provided `PRINOS` file matches portfolio date `2026-04-24`, not `2026-04-27`.
- Completed: Diagnosed and fixed a `GENERAL` sheet compatibility issue where the code assumed one fixed column order and failed on the provided `input/ADDITIONAL_DATA.xlsx`.
- Completed: Executed the generator successfully with the corrected `input/ADDITIONAL_DATA.xlsx` and `portfolio_date=2026-04-24`.
- Completed: The corrected additional-data mapping produced eleven CSV files plus a ZIP archive and processing log.
- Completed: Stored the successful rerun in `output_rerun_20260427_090136`.
- Completed: Deleted the older `output` folder from the previous run.
- Completed: Diagnosed that `SHARES_OUTSTANDING` is currently sourced from the wrong `PRINOS` column in `find_prinos_data(...)`.
- Completed: Confirmed the code reads `row.iloc[24]` as `[NUMBER_OF_UNITS]`, but in the current `PRINOS` layout index 24 is `Fiz.osobe%`, which is why `8.33` appears in the CSV.
- Completed: Confirmed the current logic is hard-coded to one `PRINOS` column layout and does not yet implement separate class-level unit sourcing for outputs like `7BET`.
- Completed: Patched the XML parser to read all `PRINOS` worksheets instead of only the first worksheet.
- Completed: Patched the first-sheet `PRINOS` lookup to read by header names (`Datum`, `Klijent`, `Valuta`, `Broj udjela`, `Cijena udjela`) instead of stale positional indexes.
- Completed: Patched class-level outputs to override `[NUMBER_OF_UNITS]` and `[NAV_PER_SHARE]` from the per-class worksheets (`A17`, `B17`, `C13`, etc.) in worksheet/class order.
- Completed: Reran the generator into `output_rerun_20260427_091500`.
- Completed: Verified all 11 generated CSV files against the class worksheets for both `SHARES_OUTSTANDING` and `NAV_PER_SHARE`.
- Completed: Confirmed the template only consumes three `PRINOS`-fed placeholders: `[FUND_CURRENCY]`, `[NUMBER_OF_UNITS]`, and `[NAV_PER_SHARE]`.
- Completed: Per user request, created an additional fresh rerun in `output_rerun_20260427_091433`.
- Completed: Confirmed the new rerun also produced 11 CSV files, a ZIP archive, and a processing log.
- Completed: Confirmed `PRINOS` first-sheet parsing now uses header names for `Datum`, `Klijent`, `Valuta`, `Broj udjela`, and `Cijena udjela`.
- Completed: Confirmed class-level `PRINOS` parsing now uses header names for `Datum`, `Broj udjela`, and `Cijena udjela` / `Cijena udjela dv`.
- Completed: Confirmed `GENERAL` lookup now prefers headers like `FUND NAME`, `FUND ISIN`, `BLOOMBERG TICKER`, and `OUTPUT FILE NAME`, with positional fallback if headers are missing.
- Completed: Confirmed the code is not fully schema-agnostic yet: `PORTFELJ` parsing still relies on fixed column positions, `ADDITIONAL_DATA.xlsx` still requires specific sheet names, and class-output assignment still depends on the order of `GENERAL` rows matching the sorted class worksheet order.
- Completed: Confirmed the sampled `PORTFELJ` files all contain stable headers such as `Vrsta pozicije`, `Valuta`, `Pozicija`, `ISIN`, `Tip`, `Količina`, `Cijena`, `Tečaj DV`, `Kamata`, and `Iznos DV`.
- Completed: Confirmed those `PORTFELJ` headers repeat by section in every sampled file, so the right refactor is header-based plus section-aware, not a naive single-header-table assumption.
- Completed: Recorded `output_rerun_20260427_091433` as the baseline output set to match during future testing and refactors, unless the user changes the inputs or expected logic.
- Completed: Confirmed the current requested output fix and rerun are done; any additional `PORTFELJ` header-based refactor is optional future hardening work, not required for the current result set.
- Completed: Created branch `codex/prinos-class-fixes`, committed the current PRINOS/local-runner work, pushed it, and opened upstream PR `https://github.com/novicegitgud/tradeweb/pull/2`.
- Completed: Created follow-up branch `codex/portfelj-header-parser` for the `PORTFELJ` parser refactor.
- Completed: Replaced the positional `PORTFELJ` parser with a header-based, section-aware record extractor that reuses repeated header rows within each file.
- Completed: Reran the generator into `output_portfelj_headers_20260427_092300`.
- Completed: Verified that every generated CSV in `output_portfelj_headers_20260427_092300` matches the baseline folder `output_rerun_20260427_091433` exactly; only ZIP artifacts may differ as archives.
- Completed: Detected that the new `PRINOS` file is for portfolio date `27.04.2026`.
- Completed: Reran the generator against the updated `input` folder into `output_rerun_20260427_163556`.
- Completed: Confirmed the processing log contains 5 `Created` entries and no skipped files.
- Completed: Validated the generated `NAV_PER_SHARE` and `SHARES_OUTSTANDING` values against the current PRINOS class sheets for all 5 generated CSV files; no mismatches were found.
- Completed: Reran the generator again after the corrected additional-data replacement into `output_rerun_20260427_163835`.
- Completed: Confirmed the newest processing log contains 11 `Created` entries and no skipped files.
- Completed: Validated the generated `NAV_PER_SHARE` and `SHARES_OUTSTANDING` values against the current PRINOS class sheets for all 11 generated CSV files; no mismatches were found.
- Completed: Verified that the pushed commits are on remote branches `codex/prinos-class-fixes` and `codex/portfelj-header-parser`, while `main` on the fork remains unchanged.
- Completed: Changed the PRINOS NAV-per-share source to prefer `Cijena udjela dv` instead of `Cijena udjela`, while still taking `Broj udjela` from the respective class sheet.
- Completed: Reran the generator into `output_rerun_20260427_205742`.
- Completed: Confirmed the processing log contains 11 `Created` entries and no skipped files.
- Completed: Validated the generated `NAV_PER_SHARE` and `SHARES_OUTSTANDING` values against the current PRINOS class sheets for all 11 generated CSV files; no mismatches were found after the NAV-source change.
- Completed: Spot-checked the reported cross-listing cases and confirmed the outputs now use the depositary-value NAVs (for example `7SLO` / `ICSLOETF` and `7BET` / `ICBETNET` are now aligned on `FUND_BASE=EUR` and near-equal `NAV_PER_SHARE` values, while retaining class-specific shares outstanding).

## Next Step

- Use `run_local.py` for direct folder-based execution:
  - `C:\Users\LukaPerko\AppData\Local\Programs\Python\Python312\python.exe run_local.py --portfolio-date 2026-04-24`
- The newest rerun output set is in `output_rerun_20260427_091433\`.
- The newest rerun ZIP archive is `output_rerun_20260427_091433\tradeweb_csvs_20260427.zip`.
- The verified reference rerun remains `output_rerun_20260427_091500\`.
- The earlier rerun folder `output_rerun_20260427_090136\` contains the pre-fix output and should not be used as the final verified result.
- The current upstream PR for the PRINOS/local-runner fixes is `https://github.com/novicegitgud/tradeweb/pull/2`.
- The current follow-up refactor output is in `output_portfelj_headers_20260427_092300\`.
- The next hardening step after this would be to remove the remaining class-assignment-by-order assumption and map `GENERAL` rows to `PRINOS` classes explicitly.
- For future testing, treat `output_rerun_20260427_091433\` as the expected comparison baseline.
- The latest user-input rerun output is in `output_rerun_20260427_163556\`.
- The newest corrected-additional-data rerun output is in `output_rerun_20260427_163835\`.
- The fork `main` branch still points to `68d5d29`; pushed work currently lives on feature branches until merged.
- The latest NAV-source-corrected rerun output is in `output_rerun_20260427_205742\`.
