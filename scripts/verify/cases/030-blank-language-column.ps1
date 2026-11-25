param(
    [Parameter(Mandatory = $true)] [string] $Source,
    [Parameter(Mandatory = $true)] [string] $Target,
    [Parameter(Mandatory = $true)] [string] $Language,
    [switch] $Visible
)

# Make the fixture verify case 030 builds from: a copy of a setup whose
# translation column for ONE language has been emptied, with every other
# column left as it was.
#
# WHAT THE SHEET LOOKS LIKE, measured 2026-11-25
# -----------------------------------------------------------------------------
# The Translations worksheet carries one table, Tab_Translations, whose header
# row holds the LANGUAGE NAMES and whose column A, outside the table, holds the
# tag each row translates:
#
#   measles.xlsb   Tab_Translations at B4:D516, 512 rows
#                  header: English | Portuguese | Français
#   generic.xlsb   Tab_Translations at B4:F154, 150 rows
#                  header: English | Français | Espanol | Portugues | Arabic
#
# So emptying a language is clearing one column of the table's body, which is
# what this does. The header stays: a column that disappeared would be a
# different fault, and the case is about a column that is THERE and EMPTY.
#
# It prints the row counts before and after so the shell case can assert the
# fixture is what it asked for rather than trusting this script.

$ErrorActionPreference = "Stop"

# Write UTF-8, so a language name with an accent reaches the log as itself
# rather than as a replacement character. This is for the HUMAN-READABLE lines:
# an assertion still goes through an ASCII marker, because the console code page
# is not something a case should have to depend on.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

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
    $sh = $wb.Worksheets("Translations")

    # Translations is one of the PASSWORD-PROTECTED sheets of a setup, so
    # clearing a column of it needs the protection lifted first. Without this
    # the clear below answered "The cell or chart you're trying to change is on
    # a protected sheet" and this case skipped without ever testing anything.
    # Lift-SheetProtection cannot prompt: see the rule at the top of
    # ../lib/excel-helpers.ps1.
    if (-not (Lift-SheetProtection -Workbook $wb -Sheet $sh)) {
        throw "the Translations sheet is protected and would not unprotect"
    }

    if ($sh.ListObjects.Count -lt 1) { throw "the Translations sheet carries no table" }
    $lo = $sh.ListObjects.Item(1)

    $names = @()
    foreach ($c in $lo.ListColumns) { $names += $c.Name }
    "table    : $($lo.Name) at $($lo.Range.Address($false, $false))"
    "columns  : $($names -join ' | ')"

    $col = $null
    foreach ($c in $lo.ListColumns) { if ($c.Name -eq $Language) { $col = $c; break } }
    if ($null -eq $col) {
        throw "no column called '$Language'; the table holds: $($names -join ', ')"
    }

    $rows = if ($null -eq $lo.DataBodyRange) { 0 } else { $lo.DataBodyRange.Rows.Count }
    $before = $xl.WorksheetFunction.CountA($col.DataBodyRange)
    "rows     : $rows"
    "before   : $Language filled=$before"

    [void]$col.DataBodyRange.ClearContents()

    $after = $xl.WorksheetFunction.CountA($col.DataBodyRange)
    "after    : $Language filled=$after"

    # What is left filled, so the case knows a fallback has somewhere to go.
    $others = @()
    foreach ($c in $lo.ListColumns) {
        if ($c.Name -ne $Language) {
            $n = $xl.WorksheetFunction.CountA($c.DataBodyRange)
            $others += "$($c.Name)=$n"
        }
    }
    "others   : $($others -join ' ')"

    $xl.EnableEvents = $false
    $wb.Save()
    $wb.Close($false)
    $wb = $null

    if ($after -ne 0) { "ERROR: the column is still filled"; exit 1 }

    # BLANKED carries the language name for a person reading the log. The case
    # must NOT assert on it: a language name holds accents -- Francais is
    # spelled with a cedilla -- and this script's stdout reaches bash through
    # the console code page, so the byte for the cedilla does not match the
    # UTF-8 the case holds. Case 030 failed on exactly that, grepping for
    # BLANKED=Fran<ccedil>ais against a log that spelled it differently, while
    # the column really had been emptied.
    #
    # BLANKED_FILLED_AFTER is the assertable one: ASCII, and it is the number
    # the question actually turns on.
    "BLANKED=$Language"
    "BLANKED_FILLED_AFTER=$after"
    "BLANKED_FILLED_BEFORE=$before"
    "ROWS=$rows"
    exit 0
} catch {
    "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    Close-QuietExcel $xl
}
