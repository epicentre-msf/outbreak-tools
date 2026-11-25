param(
    [Parameter(Mandatory = $true)] [string] $Source,
    [Parameter(Mandatory = $true)] [string] $Target,
    [Parameter(Mandatory = $true)] [string] $Tables,
    [switch] $IncludeHeader,
    [switch] $Visible
)

# Make the fixture case 040 imports from: a copy of a setup whose tables have
# had ROWS DELETED, so the sheet below each one has moved.
#
# -Tables takes a COMMA-SEPARATED list of "Sheet:TableName" specs, e.g.
#   -Tables "Analysis:Tab_Global_Summary,Translations:Tab_Translations"
#
# One delimited string rather than a [string[]]: powershell.exe -File does not
# split a comma-separated argument into array elements the way the parameter
# binder does in-process, so a [string[]] arrived here as one long element and
# the first spec swallowed the rest. Splitting here is predictable from bash.
#
# WHY THIS MANUFACTURES THE FAULT INSTEAD OF FINDING ONE
# -----------------------------------------------------------------------------
# Case 040 used to convert malaria-aweil.xlsb, which raised error 91 on its
# Translations sheet because Tab_Translations had no data body. That fault was
# then FIXED -- case 060 covers it and passes -- so 040 started failing for the
# best possible reason: the bug it pinned was gone. A case that depends on a
# setup happening to be broken dies the day somebody fixes it, and takes its
# coverage with it silently.
#
# So the fault is made here, on a throwaway copy, and 040 keeps testing the
# behaviour after the next fix too.
#
# HOW THIS DIFFERS FROM 050 AND 060, WHICH ALSO EMPTY THINGS
# -----------------------------------------------------------------------------
#   050  clears a whole sheet's CELLS. The tables keep their height and
#        addresses; DataBodyRange stays a real range over blank cells.
#   060  deletes ONE table's DataBodyRange. The cells below it inside that
#        range shift up, the body becomes Nothing, nothing else moves.
#   040  deletes the table's ENTIRE WORKSHEET ROWS. Everything below on that
#        sheet is dragged upward, so every OTHER table beneath it changes
#        address too. On Analysis, where nine tables are stacked down the
#        sheet, emptying the first moves the other eight.
#
# That last shape is the one an old hand-edited setup arrives in, and it is the
# fault 040's header has always described: "a sheet whose table moved".
#
# The script reports each table before and after so the case can assert the
# fixture is the shape it asked for rather than trusting this script.

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "..\lib\excel-helpers.ps1")

if (-not (Test-Path $Source)) { "ERROR: no setup at $Source"; exit 1 }

$targetDir = Split-Path $Target -Parent
if ($targetDir -and -not (Test-Path $targetDir)) {
    [void](New-Item -ItemType Directory -Path $targetDir -Force)
}
Copy-Item $Source $Target -Force

$deleted = 0
$skipped = 0

$xl = $null
$wb = $null
try {
    $xl = New-QuietExcel -Visible:$Visible
    $xl.ScreenUpdating = $false

    # Explicit empty Password and WriteResPassword: an encrypted workbook raises
    # rather than putting up a dialog. See the rule atop ../lib/excel-helpers.ps1.
    $wb = $xl.Workbooks.Open($Target, $false, $false, [Type]::Missing, "", "")

    $specs = @($Tables.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($specs.Count -lt 1) { throw "-Tables named no table" }

    foreach ($spec in $specs) {
        $parts = $spec.Split(":", 2)
        if ($parts.Count -ne 2) { throw "'$spec' is not a Sheet:Table spec" }
        $sheetName = $parts[0]
        $tableName = $parts[1]

        $sh = $null
        try { $sh = $wb.Worksheets($sheetName) } catch { }
        if ($null -eq $sh) {
            $names = @()
            foreach ($s in $wb.Worksheets) { $names += $s.Name }
            throw "no worksheet called '$sheetName'; the setup holds: $($names -join ', ')"
        }

        if (-not (Lift-SheetProtection -Workbook $wb -Sheet $sh -Quiet)) {
            throw "the '$sheetName' worksheet is protected and would not unprotect"
        }

        $lo = $null
        try { $lo = $sh.ListObjects.Item($tableName) } catch { }
        if ($null -eq $lo) {
            $tabs = @()
            foreach ($t in $sh.ListObjects) { $tabs += $t.Name }
            throw "no table called '$tableName' on '$sheetName'; it holds: $($tabs -join ', ')"
        }

        $rangeBefore = $lo.Range.Address($false, $false)

        if ($null -eq $lo.DataBodyRange) {
            "TABLE=${sheetName}:${tableName} BEFORE=Nothing AFTER=Nothing RANGE=$rangeBefore NOTE=already-empty"
            $skipped++
            continue
        }

        $rowsBefore = $lo.DataBodyRange.Rows.Count

        # THE ONE CHANGE. EntireRow, so the sheet below moves with it.
        # [void] because Range.Delete() answers through COM and a bare call
        # leaks "True" into this script's output.
        #
        # -IncludeHeader takes the HEADER rows with the body, which destroys the
        # ListObject outright: the table is gone from the sheet afterwards, and
        # an import looking for it by name finds nothing. That is the harder
        # fault, and the only one left that the import still reports -- body
        # rows alone it now absorbs without complaint.
        if ($IncludeHeader) {
            [void]$lo.Range.EntireRow.Delete()
        } else {
            [void]$lo.DataBodyRange.EntireRow.Delete()
        }

        # The ListObject may not exist any more, so every read of it is guarded.
        $bodyAfter = "gone"
        $rangeAfter = "gone"
        try {
            $bodyAfter = $(if ($null -eq $lo.DataBodyRange) { "Nothing" } else { "$($lo.DataBodyRange.Rows.Count)" })
            $rangeAfter = $lo.Range.Address($false, $false)
        } catch { }

        # Whether the sheet still carries a table of that name at all.
        $stillThere = "no"
        try { if ($null -ne $sh.ListObjects.Item($tableName)) { $stillThere = "yes" } } catch { }

        "TABLE=${sheetName}:${tableName} BEFORE=$rowsBefore AFTER=$bodyAfter RANGE_BEFORE=$rangeBefore RANGE_AFTER=$rangeAfter STILL_THERE=$stillThere"
        $deleted++
    }

    $xl.EnableEvents = $false
    $wb.Save()
    $wb.Close($false)
    $wb = $null

    "DELETED=$deleted"
    "ALREADY_EMPTY=$skipped"
    if ($deleted -lt 1) {
        "ERROR: nothing was deleted, so the fixture proves nothing"
        exit 1
    }
    exit 0
} catch {
    "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    Close-QuietExcel $xl
}
