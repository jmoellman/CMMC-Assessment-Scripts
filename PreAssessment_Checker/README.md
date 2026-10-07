# CMMC Pre-Assessment Package Checker v1.0.0

Windows desktop app that checks an OSC's pre-assessment package before Phase 2 and before the eMASS Pre-Assessment upload. It reads the files, lists every issue it finds, and lets you view, print, or export a report. It runs fully offline and never changes the source files.

## Files it checks

| File | What it is |
|---|---|
| Pre-Assessment Form | eMASS CMMC_Level2_PreAssessment_Form (first tab, template v3.8) |
| Asset Inventory Scoping | Ignyte CMMC L2 Asset Inventory Scoping (Main Scoped Assets, Consolidated ESP, Out of Scope Areas) |
| Document Request List | Ignyte CMMC L2 Assessment Document Request List (Evidence Plan, Physical Facilities, Physical Security, 171 R2 Control Ownership, OSC Assessment Team) |
| Security Assessment Plan | The OSC-completed CMMC L2 SAP (.docx) |

Any file can be left out. Cross-file checks run when two or more files are loaded.

## Setup (one time)

1. Install Python 3.10 or newer from python.org. Check "Add python.exe to PATH" during install.
2. Unzip this folder anywhere (for example `Documents\PreAssessment_Checker`).
3. Double-click `Run_PreAssessment_Checker.bat`. The first run installs three packages (openpyxl, python-docx, lxml) into a local `.venv` folder. After that it opens right away.

Optional: `Build_EXE.bat` builds a standalone app folder so the checker runs on a PC without Python.

## How to use

1. **Load files.** Click **Load a folder...** and pick the OSC's folder, or **Add files...** and select the files. Each file is recognized automatically. Use **Browse...** on a row to assign a file by hand.
2. **Run checks.** Click **Run checks**. The counts show Errors, Warnings, Info, and how many items can be auto-fixed.
3. **Review.** Filter by severity, file, or owner, or type in Search. Click a row to see what was found, what is expected, and how to resolve it.
4. **Report.**
   - **View report** opens the HTML report in your browser.
   - **Print report** opens it with the print dialog. Pick a printer or "Save as PDF".
   - **Report contents** lets you produce an OSC-only report (to send to the customer) or an Ignyte-only report.
   - **Export to Excel / CSV** saves the issue log.
5. **Corrected copy (optional).** **Save corrected Pre-Assessment...** shows each proposed correction with a checkbox. Uncheck anything you do not want, then save. The copy is saved as `<name>_corrected_v1.xlsx` (v2, v3 on later saves). The original is never overwritten, and the template dropdowns are preserved. You are then offered a re-check of the corrected copy.

Reports are saved in a `Checker_Reports` folder beside the loaded files, so they stay with the engagement files (they may contain CUI).

## Severity and owner

- **Error**: blocks the eMASS import, or a required SAP/template item is missing.
- **Warning**: likely problem to review.
- **Info**: confirm with the OSC or Ignyte team.
- **Owner**: OSC or Ignyte. Per SAP section 5.1, Ignyte completes the C3PAO Contract Date, Planning Start/Completion Dates, and Assessment Standard rows, so those are tagged Ignyte.

## What it checks

**Pre-Assessment Form**
- Values typed in the "Acceptable Values" column (C) instead of "Input" (D). eMASS reads column D only. Auto-fix moves them and restores the guidance text.
- Required fields blank or holding placeholders (TBD, N/A, etc.).
- UEI: 12 letters/numbers, uppercase, no I or O, not starting with 0.
- HLO CAGE: one code only. CAGE codes in scope: 5 characters, semicolon separated (SAP 5.1), no duplicates, no I or O, HLO listed or confirmed out of scope.
- Scope: Enterprise or Enclave. Enclave requires Scope Description, 1000 characters max.
- Sector: eMASS lookup values only, semicolon separated; "Other" requires Sector (Other).
- URL must include http:// or https://. Zip and phone stored as text. Number of Employees is a whole number.
- Dates stored as DD-MMM-YYYY text, in order (contract, planning start, planning completion).
- Assessment Standard is a lookup value; Revision 3 is flagged for confirmation.
- Template version label and the hidden Lookup Values sheet.

**Asset Inventory Scoping**
- Main Scoped Assets: columns A to F required; category and type from the dropdowns; quantity numeric; FedRAMP listed without a package ID; rows copied from the example tab; duplicates; no CUI assets or no Security Protection Assets.
- Consolidated ESP: sample record still present; all columns except I required; email format; CMMC Status None/Level 2/Level 3; sector lookup; duplicates.
- Out of Scope Areas: listed, with name and description.

**Document Request List**
- All five tabs present.
- Evidence Plan: 110 requirements and 320 objectives intact; Document and Artifact objectives mapped to file names; each mapped file name starts with its ERL # (SAP section 8).
- 171 R2 Control Ownership: Control Owner (G) and Control Operator (H) completed for every objective; names not on the OSC Assessment Team tab.
- Physical Facilities: at least one facility; columns A to F; ownership Rented/Leased/Own/Other; Yes/No fields; CAGE format.
- Physical Security: rows still showing "[SELECT FROM".
- OSC Assessment Team: at least one person; first and last name, email, title, responsibility; valid and unique emails. Reminder that eMASS needs an Assessment Official and Technical POC with phone numbers.

**Security Assessment Plan**
- Section 3: SPRS submission type, SPRS score between -203 and 110, DIBCAC answer and its follow-up items.
- Every Yes/No affirmation in sections 4 and 9 answered; any "No" flagged; exactly one COI statement marked Yes.
- Unreplaced placeholders: `<System Name>`, `[insert ...]`, `SSP Version X.X`, `Section X`, `Page X`, blanks.
- Plan Acceptance name, title, and date for the OSC and Ignyte signers.

**Cross-file**
- OSC name on the SAP matches the Pre-Assessment form.
- SAP SPRS submission type matches Pre-Assessment Scope.
- Facility CAGE codes appear in the Pre-Assessment CAGE list.
- Cloud-hosted assets match an ESP record, and ESPs exist when cloud hosting is listed.
- SAP "Yes" affirmations conflict with empty tabs (assets, out-of-scope areas, team, document mapping).
- The SAP affirming official appears on the OSC Assessment Team tab.

## Changing rules

Edit `config\rules.json` (Notepad is fine; keep valid JSON). You can change accepted template versions, lookup lists, SAP affirmation items, placeholder patterns, and which Pre-Assessment fields Ignyte owns. Restart the app after saving.

## Command line (optional)

```
.venv\Scripts\python.exe app.py --cli "C:\path\to\OSC folder\*.xlsx" --html report.html --xlsx issues.xlsx
.venv\Scripts\python.exe app.py --cli --pa form.xlsx --ai assets.xlsx --drl drl.xlsx --sap sap.docx --owner OSC --html osc_report.html
```

## Files

`app.py` (GUI and CLI), `engine.py` (checks), `reports.py` (HTML, Excel, CSV), `xlsx_patch.py` (safe corrected-copy writer), `config\rules.json` (editable rules), `assets\app.ico`, `requirements.txt`, `Run_PreAssessment_Checker.bat`, `Build_EXE.bat`.
