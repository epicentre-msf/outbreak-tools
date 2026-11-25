param(
    [Parameter(Mandatory = $true)] [string] $Source,
    [Parameter(Mandatory = $true)] [string] $Target,
    [Parameter(Mandatory = $true)] [string] $Sheet,
    [string] $Table = "",
    [switch] $Visible
)

# Make the fixture verify case 060 imports from: a copy of a setup where one
# table has had its DATA ROWS DELETED, so DataBodyRange is Nothing.
#
# THIS IS A DIFFERENT SHAPE FROM CASE 050's FIXTURE
# -----------------------------------------------------------------------------
# 050 clears the CELLS. Excel keeps the table's height, so DataBodyRange stays
# a real range over blank cells, and every sheet imported fine. This one
# DELETES THE ROWS: the table drops to its header plus the phantom insert row,
# DataBodyRange becomes Nothing, and that is the shape that broke the import.
#
# It is the shape a real setup ships in. `setup_20251201_aweil 1.xlsb` carries
# Tab_Translations at B4:E5 with DataBodyRange Is Nothing -- headers for
# English, Francais, Espanol and Portugese, and no translation rows at all,
# because nobody ever translated that setup.
#
# The script prints the table address and the body state before and after, so
# the case can assert the fixture is the shape it asked for rather than
# trusting this script.

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "..\lib\excel-helpers.ps1")

if (-not (Test-Path $Source)) { "ERROR: no setup at $Source"; exit 1 }

$targetDir = Split-Path $Target -Parent
if ($targetDir -and -not (Test-Path $targetDir)) {
    [void](New-Item -ItemType Directory -Path $targetDir -Force)
}
Copy-Item $Source $Target -Force

$xl = $null
$wb = $null
try {
    $xl = New-QuietExcel -Visible:$Visible
    $wb = $xl.Workbooks.Open($Target, $false, $false)

    $sh = $null
    try { $sh = $wb.Worksheets($Sheet) } catch { }
    if ($null -eq $sh) {
        $names = @()
        foreach ($s in $wb.Worksheets) { $names += $s.Name }
        throw "no worksheet called '$Sheet'; the setup holds: $($names -join ', ')"
    }

    # Lift-SheetProtection cannot prompt: see the rule at the top of
    # ../lib/excel-helpers.ps1. Every import sheet of a setup is protected.
    if (-not (Lift-SheetProtection -Workbook $wb -Sheet $sh)) {
        throw "the '$Sheet' worksheet is protected and would not unprotect"
    }

    $lo = $null
    if ($Table) {
        try { $lo = $sh.ListObjects.Item($Table) } catch { }
        if ($null -eq $lo) { throw "no table called '$Table' on '$Sheet'" }
    } else {
        if ($sh.ListObjects.Count -lt 1) { throw "the '$Sheet' worksheet carries no table" }
        $lo = $sh.ListObjects.Item(1)
    }

    # Read BEFORE the close: $lo is a dead reference once the workbook is
    # closed, and the name then comes back empty.
    $tableName = $lo.Name
    "sheet    : $Sheet"
    "table    : $tableName"
    "range    : $($lo.Range.Address($false, $false))"

    if ($null -eq $lo.DataBodyRange) {
        "ERROR: '$($lo.Name)' already has no data body, so the fixture proves nothing"
        exit 1
    }
    $rowsBefore = $lo.DataBodyRange.Rows.Count
    "body     : $rowsBefore row(s)"

    # THE ONE CHANGE. Delete rather than clear.
    # [void] because Range.Delete() answers a value through COM, and a bare
    # call leaks "True" into this script's output.
    [void]$lo.DataBodyRange.Delete()

    $bodyAfter = $(if ($null -eq $lo.DataBodyRange) { "Nothing" } else { "$($lo.DataBodyRange.Rows.Count) row(s)" })
    "range after: $($lo.Range.Address($false, $false))"
    "body after : $bodyAfter"

    $xl.EnableEvents = $false
    $wb.Save()
    $wb.Close($false)
    $wb = $null

    if ($bodyAfter -ne "Nothing") {
        "ERROR: the table still has a data body, so this is not the zero-row shape"
        exit 1
    }

    "ROWS_BEFORE=$rowsBefore"
    "BODY_AFTER=Nothing"
    "EMPTIED=$tableName"
    exit 0
} catch {
    "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    Close-QuietExcel $xl
}
