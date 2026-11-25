param(
    [Parameter(Mandatory = $true)] [string] $Source,
    [Parameter(Mandatory = $true)] [string] $Target,
    [Parameter(Mandatory = $true)] [string] $Sheet,
    [switch] $Visible
)

# Make the fixture verify case 050 imports from: a copy of a setup with ONE of
# its worksheets emptied of data, every other sheet left as it was.
#
# WHAT "EMPTY" MEANS HERE
# -----------------------------------------------------------------------------
# The body of every table on the sheet is cleared and the HEADERS STAY. So the
# sheet is structurally what the import expects and holds no data, which is the
# case a user creates by deleting their rows. A sheet whose headers had gone
# too would be a malformed sheet, a different fault, and a different case.
#
# WHAT THE FIVE SHEETS HOLD, measured 2026-11-25 on ebola-2026-11-19.xlsb
# -----------------------------------------------------------------------------
#   Dictionary    1 table   Tab_Dictionary              381 rows
#   Choices       1 table   Tab_Choices                 259 rows
#   Exports       1 table   Tab_Export                    3 rows
#   Analysis      9 tables  Tab_Global_Summary, Tab_Univariate_Analysis,
#                           Tab_Bivariate_Analysis, Tab_TimeSeries_Analysis,
#                           Tab_Label_TSGraph, Tab_Graph_TimeSeries,
#                           Tab_Spatial_Analysis,
#                           Tab_SpatioTemporal_Analysis,
#                           Tab_SpatioTemporal_Specs
#   Translations  1 table   Tab_Translations            542 rows
#
# Analysis carrying NINE tables is why this clears every table on the sheet
# rather than the first one. Clearing one of nine would leave the sheet full.
#
# Excel keeps at least one body row in a table, so an emptied table still
# answers rows=1 with nothing in it. The measurement that matters is CountA
# over the body, which this prints before and after so the case can assert the
# fixture is what it asked for rather than trusting this script.

$ErrorActionPreference = "Stop"

if (-not (Test-Path $Source)) { "ERROR: no setup at $Source"; exit 1 }

$targetDir = Split-Path $Target -Parent
if ($targetDir -and -not (Test-Path $targetDir)) {
    [void](New-Item -ItemType Directory -Path $targetDir -Force)
}
Copy-Item $Source $Target -Force

. (Join-Path $PSScriptRoot "..\lib\excel-helpers.ps1")

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

    # Counted BEFORE the close. $sh is a dead reference once the workbook is
    # closed, and reading ListObjects.Count off it then answers 0, which the
    # first version of this script reported as TABLES=0 on a sheet with one.
    $tableCount = $sh.ListObjects.Count
    "sheet    : $Sheet"
    "tables   : $tableCount"
    if ($tableCount -lt 1) { throw "the '$Sheet' worksheet carries no table" }

    # EVERY ONE OF THE FIVE SHEETS IS PROTECTED, measured 2026-11-25.
    # Dictionary and Choices leave their table bodies unlocked, so clearing
    # them needs no lift at all. Analysis and Translations are locked and hold
    # the password the setup keeps in its own __pass sheet.
    #
    # Lift-SheetProtection does it without ever prompting, which is the whole
    # reason it exists: see the rule at the top of ../lib/excel-helpers.ps1.
    if (-not (Lift-SheetProtection -Workbook $wb -Sheet $sh)) {
        throw "the '$Sheet' worksheet is protected and would not unprotect, so it cannot be emptied"
    }

    $before = 0
    $after = 0
    foreach ($lo in $sh.ListObjects) {
        if ($null -eq $lo.DataBodyRange) {
            "  {0,-32} body already absent" -f $lo.Name
            continue
        }
        $b = $xl.WorksheetFunction.CountA($lo.DataBodyRange)
        [void]$lo.DataBodyRange.ClearContents()
        $a = $xl.WorksheetFunction.CountA($lo.DataBodyRange)
        $before += $b
        $after += $a
        "  {0,-32} cells {1} -> {2}" -f $lo.Name, $b, $a
    }

    "CELLS_BEFORE=$before"
    "CELLS_AFTER=$after"

    $xl.EnableEvents = $false
    $wb.Save()
    $wb.Close($false)
    $wb = $null

    if ($after -ne 0) { "ERROR: the sheet still holds $after filled cell(s)"; exit 1 }
    if ($before -eq 0) { "ERROR: the sheet was already empty, so the fixture proves nothing"; exit 1 }

    "EMPTIED=$Sheet"
    "TABLES=$tableCount"
    exit 0
} catch {
    "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    Close-QuietExcel $xl
}
