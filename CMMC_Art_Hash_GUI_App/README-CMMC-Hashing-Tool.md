# CMMC Artifact Hashing Tool

A point-and-click Windows Forms wrapper around the `ArtifactHash.ps1` v1.11 logic from WI-C3PAO-HASH-001, so assessment participants don't need to run PowerShell commands by hand.

## Requirements

- Windows, with PowerShell 7 or higher installed (same requirement as the original script).

## How to run

1. Copy `CMMC-Artifact-Hashing-Tool.ps1` to the laptop.
2. Right-click the file and choose **Run with PowerShell**.
   - If your organization blocks script execution, run it from a PowerShell 7 window instead: `pwsh -ExecutionPolicy Bypass -File .\CMMC-Artifact-Hashing-Tool.ps1`
3. In the window that opens:
   - Browse to the artifact folder (the final, locked evidence folder).
   - Leave "Use the same folder as the artifacts" checked, or pick a different output folder.
   - Leave "Create HashedArtifacts.zip" checked (needed for the 6-year OSA retention copy).
   - Check the confirmation box that the folder is final and locked.
   - Click **Run Hashing**.
4. When it finishes, the four eMASS fields are filled in on screen. Use **Copy All Fields** to copy all of them at once, or the individual **Copy** buttons next to each field.

## What it produces

Same three outputs as the original script, in the output folder you chose:

- `CMMCAssessmentArtifacts.log` — every artifact file with its SHA-256 hash → eMASS "Hashed Data List"
- `CMMCAssessmentLogHash.log` — the hash of that log file → eMASS "Hash Value"
- `HashedArtifacts.zip` — a copy of the artifact folder (including the two log files) for the OSA's 6-year retention requirement

## Notes on this version

- If a previous run's log/zip files are already sitting in the artifact folder, the tool excludes them from the new hash listing so a re-run never hashes its own prior output.
- The WI documents mentioned `HashedArtifacts.zip` as an output, but neither attached script (`ArtifactHash.ps1` / `ArtifactHash.py`) actually created it — this tool adds that step.
- This wraps the PowerShell version's logic specifically. If you'd rather distribute a Python-based build (Tkinter GUI + a `build.bat` to produce a standalone `.exe` via PyInstaller), let me know and I'll put that together too.

## v1.1 fix: files with brackets in the name were being skipped

Testing turned up files with square brackets in the filename (a common pattern in this evidence set, e.g. `3.1.10[b] E-AC-32 ....png`) coming out with a blank Algorithm/Hash in the log, even though the path still showed up.

Cause: the tool called `Get-FileHash -Path <file>`. The `-Path` parameter treats its input as a wildcard pattern, and PowerShell reads `[b]` as "match a single literal `b`" rather than the literal characters `[`, `b`, `]`. Since no real file matches that pattern, the cmdlet silently found nothing for that file, and Algorithm/Hash came out blank instead of erroring. v1.1 switches every file-resolving cmdlet (`Get-FileHash`, `Get-ChildItem`, `Test-Path`) to `-LiteralPath`, which uses the string exactly as given — no wildcard interpretation.

If you already ran v1.0 against a folder with bracketed filenames, that log and hash value are incomplete and should not be submitted. Delete the old `CMMCAssessmentArtifacts.log`, `CMMCAssessmentLogHash.log`, and `HashedArtifacts.zip` from the artifact folder, then re-run v1.1+ once against the clean folder.

## v1.2: log output is now byte-for-byte identical to the original script's

After the v1.1 fix, every individual file's hash matched between the app and `ArtifactHash.ps1` exactly — but the two log files still weren't identical, so the final "Hash Value" (a hash of the whole log file) still differed. That was a formatting difference, not an accuracy problem: the app's log had a metadata header block the original script's didn't, used different column spacing, and listed files in a different order (alphabetical vs. the folder's natural scan order).

v1.2 removes all three differences by using the exact same approach as the original script: pipe `Get-ChildItem` results straight into `Get-FileHash`, then write the real result objects with `Out-File -Encoding ASCII -Width 1024` — no custom header, no re-sorting, PowerShell's own native table formatting. For the same artifact folder, this tool and `ArtifactHash.ps1` now produce byte-for-byte identical `CMMCAssessmentArtifacts.log` and `CMMCAssessmentLogHash.log` files, so either tool's output can be used interchangeably.

This also turned out to be a cleaner fix for the v1.1 bracket issue: piping `Get-ChildItem` objects into `Get-FileHash` avoids the wildcard problem entirely, the same way the original script always did.

One deliberate difference remains: this tool still skips its own previous output files (log/hash-log/zip) if they're already sitting in the artifact folder from an earlier run, which the original script does not do. That only comes into play if the folder isn't cleaned before the final run — always start from a clean folder with no prior `CMMCAssessmentArtifacts.log`, `CMMCAssessmentLogHash.log`, or `HashedArtifacts.zip` in it, and both tools will agree.
