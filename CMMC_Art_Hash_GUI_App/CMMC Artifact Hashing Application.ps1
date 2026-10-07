<#
CMMC Artifact Hashing Tool
Version 1.2 (GUI wrapper)

Wraps the core hashing logic of ArtifactHash.ps1 v1.11 (Ignyte Federal) in a
Windows Forms interface so assessment participants don't have to run the
script from a PowerShell prompt.

Produces the same three outputs described in WI-C3PAO-HASH-001:
  - CMMCAssessmentArtifacts.log   -> eMASS field "Hashed Data List"
  - CMMCAssessmentLogHash.log     -> eMASS field "Hash Value"
  - HashedArtifacts.zip           -> retained by the OSA for 6 years

Requires PowerShell 7+ on Windows (same prerequisite as ArtifactHash.ps1).

HOW TO RUN
  Right-click this file -> "Run with PowerShell"
  -or-
  From a terminal: pwsh -ExecutionPolicy Bypass -File .\CMMC-Artifact-Hashing-Tool.ps1

VERSION HISTORY
  1.2 - Log output is now byte-for-byte identical to ArtifactHash.ps1 v1.11's
        output for the same folder. Replaced the hand-rolled log formatter
        with the same approach the original script uses: pipe Get-ChildItem
        results directly into Get-FileHash, then write the resulting objects
        with Out-File -Encoding ASCII -Width 1024 -- no custom header block,
        no re-sorting, native PowerShell table formatting. This also happens
        to be a cleaner fix for the v1.1 bracket-filename issue, since piping
        FileInfo objects into Get-FileHash (rather than calling it with
        -Path <string>) avoids the wildcard problem entirely. One retained,
        intentional difference: this tool still skips its own previous
        output files (log/hash-log/zip) if they're already sitting in the
        artifact folder from an earlier run, which the original script does
        not do. That only matters if the folder isn't cleaned before the
        final run -- see the README.
  1.1 - Fixed a bug where files with square brackets in the filename (e.g.
        "3.1.10[b] E-AC-32 ....png", a common CMMC control-reference naming
        pattern) were silently skipped: Algorithm/Hash came out blank while
        the path still appeared in the log. Root cause: Get-FileHash was
        called with -Path, which treats the string as a wildcard pattern, so
        PowerShell read "[b]" as a character class instead of a literal
        bracket and found no matching file. Switched all file-resolving
        cmdlets to -LiteralPath, which treats the string as-is.
  1.0 - Initial release.
#>

# ---------------------------------------------------------------------------
# Relaunch in STA mode if needed (WinForms requires a single-threaded apartment)
# ---------------------------------------------------------------------------
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $exePath = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exePath -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ToolVersion = "1.2"
$CoreScriptVersion = "1.11"
$LogFileName = "CMMCAssessmentArtifacts.log"
$HashLogFileName = "CMMCAssessmentLogHash.log"
$ZipFileName = "HashedArtifacts.zip"
$ReservedOutputNames = @($LogFileName, $HashLogFileName, $ZipFileName)

# ---------------------------------------------------------------------------
# Core hashing functions (mirrors ArtifactHash.ps1 v1.11 logic)
# ---------------------------------------------------------------------------

function Get-ArtifactFileHashes {
    param(
        [string]$RootPath,
        [string]$OutputPath,
        [System.Windows.Forms.TextBox]$LogBox
    )

    # IMPORTANT: -LiteralPath (not -Path) on Get-ChildItem, and piping the
    # resulting FileInfo objects directly into Get-FileHash (rather than
    # calling Get-FileHash -Path <string> ourselves), matches exactly what
    # ArtifactHash.ps1 v1.11 does. This matters for two reasons:
    #   1. It avoids PowerShell treating brackets in a filename (e.g.
    #      "3.1.10[b] E-AC-32 ....png", a common CMMC control-reference
    #      naming pattern) as a wildcard character class, which silently
    #      failed to resolve those files in an earlier version of this tool.
    #   2. Piping real FileInfo/FileHashInfo objects into Out-File uses
    #      PowerShell's native default table formatting -- the exact same
    #      column widths, spacing, and file order the original script
    #      produces -- so the two tools' log files are byte-for-byte
    #      identical for the same folder. A hand-rolled string formatter
    #      (used in earlier versions of this tool) could never guarantee that.
    $files = @(
        Get-ChildItem -LiteralPath $RootPath -Recurse -Force -File -ErrorAction SilentlyContinue |
        Where-Object {
            # Skip this tool's own output files if the artifact root and output
            # directory are the same folder, so a re-run never hashes its own
            # prior log/zip output. (The one intentional behavior difference
            # from the original script, which has no such protection. This
            # only changes anything if stale output files are already sitting
            # in the artifact folder -- which shouldn't be the case if the
            # folder is cleaned before the final hashing run.)
            $isReservedOutput = $ReservedOutputNames -contains $_.Name
            $inOutputFolder = $_.DirectoryName.TrimEnd('\') -ieq $OutputPath.TrimEnd('\')
            -not ($isReservedOutput -and $inOutputFolder)
        }
        # No Sort-Object: preserving Get-ChildItem's natural enumeration order,
        # same as the original script, so file order in the log matches too.
    )

    $total = $files.Count
    $progressCount = 0

    # ForEach-Object's script block runs in this same scope (it's an inline
    # pipeline body, not a new child scope), so $progressCount here is the
    # same variable the "if" check reads below -- no $script: scoping needed.
    # The files themselves pass through unchanged; this step only exists to
    # drive the progress display between Get-ChildItem and Get-FileHash.
    $hashList = @(
        $files | ForEach-Object {
            $progressCount++
            if ($LogBox -and ($progressCount % 25 -eq 0 -or $progressCount -eq $total)) {
                Append-Log $LogBox "  Hashing file $progressCount of $total ..."
                [System.Windows.Forms.Application]::DoEvents()
            }
            $_
        } | Get-FileHash -Algorithm SHA256 -ErrorAction SilentlyContinue -ErrorVariable hashErrors
    )

    # Files that couldn't be hashed (e.g. permission denied) are captured here
    # for on-screen visibility, but -- matching the original script exactly --
    # they are simply absent from $hashList and therefore from the log file,
    # rather than appearing as a placeholder row.
    foreach ($e in $hashErrors) {
        Append-Log $LogBox "  WARNING: $($e.Exception.Message)"
    }

    return $hashList
}

function Append-Log {
    param([System.Windows.Forms.TextBox]$LogBox, [string]$Message)
    if (-not $LogBox) { return }
    $LogBox.AppendText("$Message`r`n")
    $LogBox.SelectionStart = $LogBox.Text.Length
    $LogBox.ScrollToCaret()
}

# ---------------------------------------------------------------------------
# Build the form
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "CMMC Artifact Hashing Tool v$ToolVersion"
$form.Size = New-Object System.Drawing.Size(760, 760)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Font = $font

$y = 10

# --- Intro label ---
$lblIntro = New-Object System.Windows.Forms.Label
$lblIntro.Text = "Generates SHA-256 hashes of CMMC assessment artifacts for eMASS submission, per WI-C3PAO-HASH-001 (32 CFR 170.17 / 170.18). C3PAO / DCMA DIBCAC assessments only."
$lblIntro.Location = New-Object System.Drawing.Point(15, $y)
$lblIntro.Size = New-Object System.Drawing.Size(720, 40)
$form.Controls.Add($lblIntro)
$y += 45

# --- Group: Artifact root ---
$grpArtifact = New-Object System.Windows.Forms.GroupBox
$grpArtifact.Text = "1. Artifact Folder (final evidence to be hashed)"
$grpArtifact.Location = New-Object System.Drawing.Point(15, $y)
$grpArtifact.Size = New-Object System.Drawing.Size(720, 70)
$form.Controls.Add($grpArtifact)

$txtArtifactRoot = New-Object System.Windows.Forms.TextBox
$txtArtifactRoot.Location = New-Object System.Drawing.Point(15, 30)
$txtArtifactRoot.Size = New-Object System.Drawing.Size(590, 24)
$grpArtifact.Controls.Add($txtArtifactRoot)

$btnBrowseArtifact = New-Object System.Windows.Forms.Button
$btnBrowseArtifact.Text = "Browse..."
$btnBrowseArtifact.Location = New-Object System.Drawing.Point(615, 28)
$btnBrowseArtifact.Size = New-Object System.Drawing.Size(90, 27)
$grpArtifact.Controls.Add($btnBrowseArtifact)

$y += 80

# --- Group: Output directory ---
$grpOutput = New-Object System.Windows.Forms.GroupBox
$grpOutput.Text = "2. Output Folder (where the log files and zip will be written)"
$grpOutput.Location = New-Object System.Drawing.Point(15, $y)
$grpOutput.Size = New-Object System.Drawing.Size(720, 100)
$form.Controls.Add($grpOutput)

$chkSameFolder = New-Object System.Windows.Forms.CheckBox
$chkSameFolder.Text = "Use the same folder as the artifacts"
$chkSameFolder.Location = New-Object System.Drawing.Point(15, 25)
$chkSameFolder.Size = New-Object System.Drawing.Size(300, 24)
$chkSameFolder.Checked = $true
$grpOutput.Controls.Add($chkSameFolder)

$txtOutputDir = New-Object System.Windows.Forms.TextBox
$txtOutputDir.Location = New-Object System.Drawing.Point(15, 55)
$txtOutputDir.Size = New-Object System.Drawing.Size(590, 24)
$txtOutputDir.Enabled = $false
$grpOutput.Controls.Add($txtOutputDir)

$btnBrowseOutput = New-Object System.Windows.Forms.Button
$btnBrowseOutput.Text = "Browse..."
$btnBrowseOutput.Location = New-Object System.Drawing.Point(615, 53)
$btnBrowseOutput.Size = New-Object System.Drawing.Size(90, 27)
$btnBrowseOutput.Enabled = $false
$grpOutput.Controls.Add($btnBrowseOutput)

$y += 110

# --- Options ---
$chkZip = New-Object System.Windows.Forms.CheckBox
$chkZip.Text = "Create HashedArtifacts.zip (recommended - required for the OSA's 6-year retention copy)"
$chkZip.Location = New-Object System.Drawing.Point(15, $y)
$chkZip.Size = New-Object System.Drawing.Size(700, 24)
$chkZip.Checked = $true
$form.Controls.Add($chkZip)
$y += 30

$chkLock = New-Object System.Windows.Forms.CheckBox
$chkLock.Text = "I confirm the artifact folder is final - no files will be added, removed, or edited after this point"
$chkLock.Location = New-Object System.Drawing.Point(15, $y)
$chkLock.Size = New-Object System.Drawing.Size(700, 24)
$chkLock.ForeColor = [System.Drawing.Color]::DarkRed
$form.Controls.Add($chkLock)
$y += 35

# --- Run button + progress bar ---
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run Hashing"
$btnRun.Location = New-Object System.Drawing.Point(15, $y)
$btnRun.Size = New-Object System.Drawing.Size(150, 36)
$btnRun.Enabled = $false
$btnRun.BackColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($btnRun)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(180, $y + 6)
$progressBar.Size = New-Object System.Drawing.Size(555, 24)
$progressBar.Style = "Marquee"
$progressBar.MarqueeAnimationSpeed = 0
$form.Controls.Add($progressBar)

$y += 46

# --- Log box ---
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Activity Log"
$lblLog.Location = New-Object System.Drawing.Point(15, $y)
$lblLog.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblLog)
$y += 20

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$txtLog.Location = New-Object System.Drawing.Point(15, $y)
$txtLog.Size = New-Object System.Drawing.Size(720, 130)
$form.Controls.Add($txtLog)
$y += 140

# --- Results group (eMASS fields) ---
$grpResults = New-Object System.Windows.Forms.GroupBox
$grpResults.Text = "3. eMASS Submission Fields"
$grpResults.Location = New-Object System.Drawing.Point(15, $y)
$grpResults.Size = New-Object System.Drawing.Size(720, 190)
$form.Controls.Add($grpResults)

function New-ResultRow {
    param([string]$LabelText, [int]$RowY, [string]$DefaultValue = "")

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $LabelText
    $lbl.Location = New-Object System.Drawing.Point(15, $RowY)
    $lbl.Size = New-Object System.Drawing.Size(230, 20)
    $grpResults.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(250, ($RowY - 3))
    $txt.Size = New-Object System.Drawing.Size(370, 24)
    $txt.ReadOnly = $true
    $txt.Text = $DefaultValue
    $grpResults.Controls.Add($txt)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "Copy"
    $btn.Location = New-Object System.Drawing.Point(630, ($RowY - 4))
    $btn.Size = New-Object System.Drawing.Size(75, 26)
    $btn.Enabled = $false
    $grpResults.Controls.Add($btn)

    return @{ TextBox = $txt; Button = $btn }
}

$rowHash = New-ResultRow -LabelText "Hash Value:" -RowY 30
$rowLogName = New-ResultRow -LabelText "Hashed Data List (log filename):" -RowY 65 -DefaultValue $LogFileName
$rowAlgo = New-ResultRow -LabelText "Hash Algorithm:" -RowY 100 -DefaultValue "SHA256"
$rowDate = New-ResultRow -LabelText "Hash Date:" -RowY 135

$rowLogName.Button.Enabled = $true
$rowAlgo.Button.Enabled = $true

$btnCopyAll = New-Object System.Windows.Forms.Button
$btnCopyAll.Text = "Copy All Fields (formatted for eMASS)"
$btnCopyAll.Location = New-Object System.Drawing.Point(250, 165)
$btnCopyAll.Size = New-Object System.Drawing.Size(280, 30)
$btnCopyAll.Enabled = $false
$grpResults.Controls.Add($btnCopyAll)

$btnOpenFolder = New-Object System.Windows.Forms.Button
$btnOpenFolder.Text = "Open Output Folder"
$btnOpenFolder.Location = New-Object System.Drawing.Point(15, 165)
$btnOpenFolder.Size = New-Object System.Drawing.Size(150, 30)
$btnOpenFolder.Enabled = $false
$grpResults.Controls.Add($btnOpenFolder)

$y += 200

$lblFooter = New-Object System.Windows.Forms.Label
$lblFooter.Text = "Based on ArtifactHash.ps1 core logic v$CoreScriptVersion (Ignyte Federal). Do not modify artifacts after hashing - notify your Ignyte assessor if a change is required."
$lblFooter.Location = New-Object System.Drawing.Point(15, $y)
$lblFooter.Size = New-Object System.Drawing.Size(720, 30)
$lblFooter.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblFooter)

$form.ClientSize = New-Object System.Drawing.Size(750, ($y + 45))

# ---------------------------------------------------------------------------
# State / script-scoped variables shared across event handlers
# ---------------------------------------------------------------------------
$script:LastOutputDir = $null

function Update-RunButtonState {
    $ready = ($txtArtifactRoot.Text.Trim() -ne "") -and
             ($chkSameFolder.Checked -or $txtOutputDir.Text.Trim() -ne "") -and
             $chkLock.Checked
    $btnRun.Enabled = $ready
    $btnRun.BackColor = if ($ready) { [System.Drawing.Color]::FromArgb(200, 230, 201) } else { [System.Drawing.Color]::LightGray }
}

# ---------------------------------------------------------------------------
# Event handlers
# ---------------------------------------------------------------------------

$btnBrowseArtifact.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select the artifact root folder to hash"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtArtifactRoot.Text = $dlg.SelectedPath
        if ($chkSameFolder.Checked) { $txtOutputDir.Text = $dlg.SelectedPath }
    }
    Update-RunButtonState
})

$btnBrowseOutput.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select the folder where the log files and zip should be written"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputDir.Text = $dlg.SelectedPath
    }
    Update-RunButtonState
})

$chkSameFolder.Add_CheckedChanged({
    $txtOutputDir.Enabled = -not $chkSameFolder.Checked
    $btnBrowseOutput.Enabled = -not $chkSameFolder.Checked
    if ($chkSameFolder.Checked) { $txtOutputDir.Text = $txtArtifactRoot.Text }
    Update-RunButtonState
})

$txtArtifactRoot.Add_TextChanged({
    if ($chkSameFolder.Checked) { $txtOutputDir.Text = $txtArtifactRoot.Text }
    Update-RunButtonState
})
$txtOutputDir.Add_TextChanged({ Update-RunButtonState })
$chkLock.Add_CheckedChanged({ Update-RunButtonState })

$rowHash.Button.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($rowHash.TextBox.Text) })
$rowLogName.Button.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($rowLogName.TextBox.Text) })
$rowAlgo.Button.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($rowAlgo.TextBox.Text) })
$rowDate.Button.Add_Click({ [System.Windows.Forms.Clipboard]::SetText($rowDate.TextBox.Text) })

$btnCopyAll.Add_Click({
    $block = @(
        "Hash Value (hash of the artifact log): $($rowHash.TextBox.Text)"
        "Hashed Data List (log filename): $($rowLogName.TextBox.Text)"
        "Hash Algorithm: $($rowAlgo.TextBox.Text)"
        "Hash Date: $($rowDate.TextBox.Text)"
    ) -join "`r`n"
    [System.Windows.Forms.Clipboard]::SetText($block)
    [System.Windows.Forms.MessageBox]::Show("All four eMASS fields copied to the clipboard.", "Copied", "OK", "Information") | Out-Null
})

$btnOpenFolder.Add_Click({
    if ($script:LastOutputDir -and (Test-Path -LiteralPath $script:LastOutputDir)) {
        Invoke-Item -LiteralPath $script:LastOutputDir
    }
})

$btnRun.Add_Click({
    $artifactRoot = $txtArtifactRoot.Text.Trim()
    $outputDir = if ($chkSameFolder.Checked) { $artifactRoot } else { $txtOutputDir.Text.Trim() }

    if (-not (Test-Path -LiteralPath $artifactRoot)) {
        [System.Windows.Forms.MessageBox]::Show("Artifact folder does not exist:`r`n$artifactRoot", "Error", "OK", "Error") | Out-Null
        return
    }
    if (-not (Test-Path -LiteralPath $outputDir)) {
        $create = [System.Windows.Forms.MessageBox]::Show("Output folder does not exist:`r`n$outputDir`r`n`r`nCreate it?", "Create Folder", "YesNo", "Question")
        if ($create -eq [System.Windows.Forms.DialogResult]::Yes) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        } else {
            return
        }
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will hash every file in:`r`n$artifactRoot`r`n`r`nDo not modify any artifact files after this step. Continue?",
        "Confirm Final Evidence Lock", "OKCancel", "Warning")
    if ($confirm -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $txtLog.Clear()
    $btnRun.Enabled = $false
    $progressBar.MarqueeAnimationSpeed = 30
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        Append-Log $txtLog "CMMC Artifact Hashing Tool v$ToolVersion (core logic v$CoreScriptVersion)"
        Append-Log $txtLog "Artifact root : $artifactRoot"
        Append-Log $txtLog "Output folder : $outputDir"
        Append-Log $txtLog ""
        Append-Log $txtLog "Scanning and hashing files (SHA-256) ..."
        [System.Windows.Forms.Application]::DoEvents()

        $timestamp = Get-Date

        $hashRows = Get-ArtifactFileHashes -RootPath $artifactRoot -OutputPath $outputDir -LogBox $txtLog

        if ($hashRows.Count -eq 0) {
            Append-Log $txtLog "WARNING: no files were found to hash."
        }

        $logFilePath = Join-Path $outputDir $LogFileName
        $hashLogFilePath = Join-Path $outputDir $HashLogFileName
        $zipFilePath = Join-Path $outputDir $ZipFileName

        # Write the artifact log using Out-File on the real Get-FileHash
        # objects, with no header/metadata block prepended and no sorting --
        # identical to how ArtifactHash.ps1 v1.11 writes this file (Out-File
        # -Encoding ASCII -Width 1024). Byte-for-byte identical output for the
        # same folder is the whole point: either tool's log/hash can be used
        # interchangeably for eMASS.
        Append-Log $txtLog "Writing $LogFileName ..."
        Out-File -FilePath $logFilePath -Force -Encoding ASCII -InputObject $hashRows -Width 1024

        Append-Log $txtLog "Hashing $LogFileName to produce $HashLogFileName ..."
        $logHash = Get-FileHash -LiteralPath $logFilePath -Algorithm SHA256
        Out-File -FilePath $hashLogFilePath -Force -Encoding ASCII -InputObject $logHash -Width 1024

        if ($chkZip.Checked) {
            Append-Log $txtLog "Creating $ZipFileName ..."
            if (Test-Path $zipFilePath) { Remove-Item $zipFilePath -Force }
            Compress-Archive -Path (Join-Path $artifactRoot '*') -DestinationPath $zipFilePath -CompressionLevel Optimal -Force
            Append-Log $txtLog "  Zip created: $zipFilePath"
        } else {
            Append-Log $txtLog "Skipped zip creation (unchecked)."
        }

        Append-Log $txtLog ""
        Append-Log $txtLog "SCRIPT COMPLETE"
        Append-Log $txtLog "  Files hashed  : $($hashRows.Count)"
        Append-Log $txtLog "  Artifact log  : $logFilePath"
        Append-Log $txtLog "  Integrity log : $hashLogFilePath"
        if ($chkZip.Checked) { Append-Log $txtLog "  Zip archive   : $zipFilePath" }

        # Populate the eMASS result fields
        $rowHash.TextBox.Text = $logHash.Hash
        $rowHash.Button.Enabled = $true
        $rowDate.TextBox.Text = $timestamp.ToString("yyyy-MM-dd")
        $rowDate.Button.Enabled = $true
        $btnCopyAll.Enabled = $true
        $btnOpenFolder.Enabled = $true
        $script:LastOutputDir = $outputDir

        [System.Windows.Forms.MessageBox]::Show(
            "Hashing complete.`r`n`r`nHash Value:`r`n$($logHash.Hash)`r`n`r`nUse 'Copy All Fields' to copy everything needed for eMASS.",
            "Done", "OK", "Information") | Out-Null
    }
    catch {
        Append-Log $txtLog "ERROR: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("An error occurred:`r`n$($_.Exception.Message)", "Error", "OK", "Error") | Out-Null
    }
    finally {
        $progressBar.MarqueeAnimationSpeed = 0
        $btnRun.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# ---------------------------------------------------------------------------
# Show the form
# ---------------------------------------------------------------------------
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()