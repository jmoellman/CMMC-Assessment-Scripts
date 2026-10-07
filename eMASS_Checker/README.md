# CMMC eMASS Results Checker

A Windows desktop app that checks a CMMC Level 2 eMASS Assessment Results workbook (template v3.9) before upload. It finds template, lookup, limit, and consistency problems, applies your eMASS reporting rules, checks spelling, grammar, and punctuation, lets you accept or edit fixes cell by cell, saves a corrected copy as a new version, and prints a report that names every issue by sheet, cell, row, and column with the reason it is an issue.

Everything runs on your laptop. No assessment text is sent anywhere. The optional LanguageTool grammar engine runs locally on 127.0.0.1.

## Setup (one time)

1. Install Python 3.10 or newer from https://www.python.org/downloads/windows/. During setup check **Add python.exe to PATH** and keep **tcl/tk and IDLE** selected.
2. Unzip this folder somewhere permanent (for example `Documents\eMASS_Checker`).
3. Double-click **Run_eMASS_Checker.bat**. The first run sets up a private environment, installs `openpyxl`, and runs a self-test. After that it opens straight to the app.

Optional: run **Build_EXE.bat** to produce `dist\eMASS_Checker\eMASS_Checker.exe`, which runs without Python installed.

### Optional: LanguageTool (deeper grammar checks, still offline)

1. Install Java 17 or newer (for example Microsoft Build of OpenJDK or Eclipse Temurin). Confirm with `java -version`.
2. In the app, open **Settings** and click **Download LanguageTool** (about 250 MB, one time; only the software is downloaded). Or download `LanguageTool-stable.zip` from https://languagetool.org/download/, unzip it, and point the **LanguageTool folder** setting at it.
3. Click **Test LanguageTool**, then tick **Grammar: LanguageTool (local)** in the Scope bar.

## Daily use

1. **Open Results File** (Ctrl+O), or drag the .xlsx onto `Run_eMASS_Checker.bat`.
2. Set the **Scope**: tick only the assigned domains, or pick an objective range (From / To). Untick sheets you are not reviewing.
3. **Run Checks** (F5). Issues appear color coded: red Error (eMASS import blocker or prohibited content), yellow Warning (guideline violation to resolve), blue Info (style).
4. Click an issue. The bottom panel shows the issue, why it is an issue, the current cell text with the flagged words highlighted, and the proposed text.
   * **Accept Change** (Ctrl+Enter) applies the proposed text. You can edit the proposed text first.
   * **Ignore Issue** (Del) closes it. Ignored issues stay ignored when you re-run checks.
   * **Add Word to Dictionary** for valid names and terms flagged as spelling.
   * **Revert Cell to Original** undoes all changes to that cell.
5. **Apply Safe Fixes** handles mechanical fixes in bulk (see below). Review everything else one by one.
6. **Save Corrected Copy** (Ctrl+S) writes `<name>_corrected_v1.xlsx` (v2, v3 on later saves) next to the original, plus a `_change_log.html`. The original file is never overwritten.
7. **Export Report** (HTML, Excel, or CSV) or **Print Report** (Ctrl+P, opens the report in your browser with the print dialog).

Use the filters (Severity, Category, Sheet, Status, Has proposed fix, Search) to work through the list.

## What it will and will not change

The app only writes to narrative text cells: **Examine, Test, Overall Comments, Findings** (rationale only), **ESP Name**, and the Assessment sheet **C3PAO Executive Summary** and **CPN** values. It never modifies Requirement Number, Objective Number, Artifacts, Interviews, Time to Assess, Score, Inherited, Standards Acceptance, Date Assessed, Requirement in POA&M, or the Findings result (MET / NOT MET / N/A). Problems in those fields are flagged for assessor review with the word "Protected" in the reason.

The corrected copy is written by editing only the changed cells inside the .xlsx package. The template's dropdown lists (Score, Inherited, POA&M, Standards Acceptance), formulas, formatting, classification labels, and hidden Lookup Values sheet are preserved byte for byte. (Saving through Excel libraries such as openpyxl would silently strip the dropdowns, which is why the app does not do that.)

**Safe fixes** (bulk-applied by Apply Safe Fixes): leading/trailing spaces, double spaces, hidden or non-breaking characters, space before punctuation, missing space after punctuation, doubled punctuation, repeated words, missing final period, sentence capitalization, lowercase "i", contractions, product and program name capitalization (for example "sharepoint" to "SharePoint", "poam" to "POA&M"), over-capitalized "Assessment Team", and "the assigned certified assessor" to "the assessment team". Everything else requires you to review and accept.

## Checks performed

### eMASS template (Instructions and Glossary)

| Rule | Severity | Check |
|---|---|---|
| EM-SHEET, EM-HEADER, EM-CLASS, EM-COLUMNS | Error | Sheets, column headers, and the CUI classification label are unchanged; no added sheets or columns |
| EM-OBJ-ID, EM-REQ-ID | Error | Requirement Number and Objective Number match the v3.9 template exactly |
| EM-REQUIRED, EM-ROW-EMPTY | Error | Required fields are filled (Assessment dates, CPNs, Executive Summary, hash fields, SSP rows, Requirement in POA&M, Inherited, Time to Assess, Score, Date Assessed, Assessed By, Findings) |
| EM-LOOKUP | Error | Exact lookup values, including case: Score (MET / NOT MET / NOT APPLICABLE; SC.L2-3.13.11 also NOT MET -3 / -5), Inherited (None / Partial / Full), POA&M (Yes / No / N/A), Standards Acceptance |
| EM-LIMIT, EM-LIMIT-ITEM | Error | Character limits: OSC Name 100, CPN 50, Executive Summary 2,000, Hash Value 100, Hashed Data List 2,000, Hash Algorithm 100, Examine / Test / Overall Comments / Findings 4,000, ESP Name 2,000, Assessed By 100, SSP Name and Version 250, Artifacts and Interviews 400 per value and 4,000 overall |
| EM-DATE, EM-DATE-ORDER, EM-DATE-RANGE, EM-HASH-DATE | Error / Warning | DD-MMM-YYYY dates; end date after start; Date Assessed inside the assessment period; hash date after start |
| EM-CPN | Error | CPN is the ID number only ("CCP-50" becomes "50") |
| EM-HASH, EM-HASH-LEN | Warning | Hash value is hexadecimal and its length matches the stated algorithm |
| EM-POAM-* | Error / Warning | Requirement in POA&M agrees with POA&M Allowed and with the objective scores |
| EM-ESP-MISSING, EM-ESP-CONFLICT | Warning | Partial / Full inheritance has an ESP name, and the reverse |
| EM-LIST-* | Warning / Info | Artifacts and Interviews use semicolons, with no empty or duplicate entries |
| EM-TIME | Error | Time to Assess is a positive number |

### Your reporting rules

| Rule | Severity | Check |
|---|---|---|
| AG-OPEN-EXAMINE / TEST / OVERALLCOMMENTS | Warning | Examine begins "The assessment team reviewed" or "examined"; Test begins "The assessment team validated"; Overall Comments begins "The assessment team confirmed" |
| AG-TEST-NOTTEST | Warning | Test text that describes a document review instead of a test |
| AG-ACTOR-PROHIBITED, AG-ACTOR-CONSISTENT, AG-ACTOR-PRONOUN | Error / Warning | "The assigned certified assessor", "the assessor", "the CCA", "we", and similar replaced by or flagged for "the assessment team" |
| AG-ADVISORY, AG-ADVISORY-REVIEW | Error / Warning | Should consider, continuous improvement, encouraged, recommend, best practice, additional enhancements, plans to improve, although this did not affect the assessment, for awareness, and similar. The proposed fix removes the clause or sentence |
| AG-STATUS-TERM, AG-TDEE | Error | PARTIAL, PENDING, partially met, "MET via TD/EE", and Temporary Deficiency or Enduring Exception used as a result category |
| AG-QUALIFIER, AG-FILLER | Warning / Info | Unsupported qualifiers (appears, seems, likely, apparently, probably...) and intensifiers (very, clearly, basically...) |
| AG-TIMEREF | Warning | Session days, weekdays, dates, and times ("Day 2", "AM session", "10:30 AM"). Proposed fix uses "during the assessment session" |
| AG-FILENAME, AG-EVIDENCE-ID | Warning | Specific filenames or evidence request IDs cited in narrative fields |
| AG-EMDASH | Warning | Em dashes |
| AG-TERM, AG-TERM-CASE, AG-CAPS, AG-ALLCAPS | Warning / Info | Exact product, program, and template terminology; inconsistent capitalization; all-caps text |
| AG-FIND-OPENER, AG-FIND-FORMAT, AG-FIND-CONFLICT | Error / Warning | Findings open with MET:, NOT MET:, or N/A: and the result agrees with the Score column |
| AG-FIND-MET-WEAKNESS | Warning | MET findings containing weakness, observation, or gap language |
| AG-FIND-NOTMET-OBJ, AG-FIND-NOTMET-DEF | Warning | NOT MET findings cite the exact unsatisfied objective and state the factual evidence deficiency |
| AG-FIND-NA-REASON | Warning | N/A findings explain why the objective does not apply |
| AG-FIND-THIN, AG-FIND-LONG | Warning / Info | Findings with no rationale; Findings longer than the summary threshold (1,000 characters by default) |
| AG-MET-EVIDENCE | Warning | MET with no Examine, Test, or Overall Comments narrative |
| AG-ROW-BLANK | Warning | Blank Artifacts or Interviews on an objective row |

### Spelling, grammar, punctuation

Offline spelling uses a bundled 166,000-word US English dictionary (SCOWL / Hunspell, see `data/DICTIONARY_LICENSE.txt`) plus CMMC, NIST, Microsoft, and security product terms. Built-in rules catch double spaces, spacing around punctuation, doubled punctuation, unbalanced parentheses and quotes, repeated words, a/an errors, sentence capitalization, missing final periods, long sentences, and contractions. LanguageTool, when enabled, adds full grammar checking.

## Customizing

**Tools > Edit Rules** opens your own copy of `rules.json` (stored in `%APPDATA%\eMASS_Checker`). You can change the required openers, the Findings label format and separator, prohibited phrases, product terms, the Findings length threshold, and the spelling whitelist. Save the file, then **Tools > Reload Rules**. Delete your copy to return to the defaults.

**Tools > Edit Custom Dictionary** opens your word list. **Add Word to Dictionary** in the app writes to the same file.

## Command line (optional)

```
.venv\Scripts\python.exe emass_checker.py check "Results.xlsx" --domains AC,AU --html report.html --xlsx report.xlsx
.venv\Scripts\python.exe emass_checker.py check "Results.xlsx" --from "AC.L2-3.1.1[a]" --to "AC.L2-3.1.22[e]" --csv report.csv
.venv\Scripts\python.exe emass_checker.py check "Results.xlsx" --apply-safe --output "Results_corrected_v1.xlsx"
.venv\Scripts\python.exe emass_checker.py --selftest
```

## Notes

* Reports and the print file contain excerpts of cell text, so they carry the CUI banner. Handle them like the results file. The print file is written to your Windows temp folder and overwritten on each print.
* Writing-quality checks are aids, not determinations. A flag on Findings, qualifiers, or advisory language means "review this", and the app never decides a result for you.
* Row and lookup checks are built for eMASS template v3.9. If eMASS publishes a new template version, the app shows a version notice.
