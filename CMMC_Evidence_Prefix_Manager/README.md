# CMMC Evidence Prefix Manager

This desktop utility aligns evidence filenames to a Document Request List (DRL). It accepts one or more complete evidence folders, including all nested subfolders, or individually selected files. It searches each filename and its relative folder path for an Objective ID or ERL number, maps the other identifier, previews the result, and copies approved files into a separate output folder with refined filenames.

Output naming format:

`ObjectiveID_ERL#_OriginalFilename.ext`

Example:

`3.1.1[a]_ERL-AC-001_Active_User_List.xlsx`

## Safety and audit features

- Source evidence is never renamed, moved, or edited.
- The original folder structure is retained under the output folder.
- Multiple evidence folders can be added to one run; each remains separated under its own top-level output folder.
- Unmatched and ambiguous files are not copied automatically.
- A reviewer can assign a mapping row manually or skip a file in the preview.
- Name collisions receive a deterministic `__dup01`, `__dup02`, and similar suffix.
- Filenames longer than 240 characters are shortened with a deterministic eight-character name hash.
- Each run produces a UTF-8 CSV change log.
- SHA-256 is calculated before and after copying to verify the file content did not change.
- Existing correct prefixes are not duplicated.

This utility assists evidence administration. It does not decide whether evidence is sufficient, adequate, applicable, MET, or NOT MET.

## Mapping file requirements

Import a `.csv`, `.tsv`, `.xlsx`, or `.xlsm` file. The tool locates the header row within the first 25 rows. At minimum, it needs these columns:

- `Objective ID`
- `ERL #`

Optional columns improve matching and reviewer context:

- `Domain`
- `Control ID`
- `Description`

Common header variations such as `Assessment Objective`, `ERL ID`, `Evidence Request ID`, `Family`, `Practice ID`, and `Requested Artifact` are recognized. Duplicate Objective ID and ERL pairs are ignored.

Descriptive header suffixes are supported. For example, the Ignyte header `ERL # (Suggested identifier for the evidence)` is recognized automatically as the ERL column.

If one ERL maps to several objectives, include the Objective ID in the filename or a matching domain/control folder name. Otherwise, the item is intentionally marked `Ambiguous` for manual assignment.

## Windows quick start

1. Install Python 3.11 or later from python.org. During setup, select `Add Python to PATH`.
2. Extract this package.
3. Double-click `Launch_Evidence_Prefix_Manager.bat`.
4. Select the DRL mapping file.
5. Use `Browse folder` once for each complete evidence folder you want to add. You may also use `Select files` for standalone evidence.
6. Select the output folder. This is where the newly named copies will be created.
7. Choose `Preview`.
8. Review all `Unmatched` and `Ambiguous` rows. Select rows and use `Apply selected mapping` or `Skip selected` as appropriate.
9. Choose `Export renamed copies`.
10. Retain the generated `change_log_<UTC timestamp>.csv` with the assessment evidence handling records.

The launcher attempts to install the optional `openpyxl` dependency if it is missing. CSV and TSV mapping files do not require that dependency.

## Command-line use

Preview only:

```powershell
python evidence_prefix_manager.py --mapping mapping.csv --source evidence --output evidence_renamed
```

Export after preview:

```powershell
python evidence_prefix_manager.py --mapping mapping.csv --source evidence --output evidence_renamed --execute
```

Use multiple `--source` arguments to process multiple folders in one run:

```powershell
python evidence_prefix_manager.py --mapping mapping.csv --source AC_Evidence --source AU_Evidence --output evidence_renamed --execute
```

Open the GUI explicitly:

```powershell
python evidence_prefix_manager.py --gui
```

## Matching behavior

The matching order is:

1. A unique Objective ID and ERL pair found in the path
2. A unique Objective ID
3. A unique ERL number
4. A unique candidate resolved by an optional domain or control folder hint
5. Manual review when multiple candidates remain

Objective notation is normalized, including lowercase objective letters, such as `3.1.1[a]`. An optional full CMMC identifier such as `AC.L2-3.1.1[a]` is preserved when supplied by the DRL.

## Change log fields

The CSV log records the UTC timestamp, processing status, match basis, Objective ID, ERL number, mapping row, source and output relative paths, old and new filename, byte size, source and output SHA-256 values, hash verification result, and reviewer note.

`Unmatched`, `Ambiguous`, and `Skipped` items appear in the log but are not copied.

## Operational recommendations

- Work from a controlled copy of the received evidence package.
- Treat the DRL mapping as assessor-controlled input and validate its Objective ID-to-ERL relationships before export.
- Review the preview instead of relying exclusively on filename inference.
- Preserve the original package, renamed output, mapping file, and change log together.
- Re-run into a new empty output folder when producing a final evidence set.

## Limitations

- Legacy `.xls` files are not supported. Save them as `.xlsx` or `.csv` first.
- The utility does not read file contents to infer controls.
- Files without a recognizable Objective ID or ERL in their filename or folder path require manual assignment.
- The desktop GUI uses Python's built-in Tkinter. Some minimal Linux installations require the separate `python3-tk` operating-system package.
