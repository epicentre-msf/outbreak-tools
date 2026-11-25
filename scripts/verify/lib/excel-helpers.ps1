# What a fixture builder needs from Excel, with no dialog anywhere.
#
# Dot-source it:
#   . (Join-Path $PSScriptRoot "..\lib\excel-helpers.ps1")
#
# THE RULE THIS FILE EXISTS FOR
# -----------------------------------------------------------------------------
# The suite runs unattended, so nothing may put up a modal dialog. The trap:
# Worksheet.Unprotect() WITH NO ARGUMENT on a sheet that has a password makes
# Excel show the password dialog, and someone has to type into it. Passing a
# string -- even the wrong one -- makes Excel raise instead, which a catch
# handles silently.
#
# Every one of a setup's five import sheets is protected. Measured 2026-11-25
# on ebola-2026-11-19.xlsb:
#
#   Dictionary    protected, table body unlocked, needs no lift to clear
#   Choices       protected, table body unlocked, needs no lift to clear
#   Exports       protected, lifts with an empty password
#   Analysis      protected, needs the workbook's own password
#   Translations  protected, needs the workbook's own password
#
# The password is read off the workbook being worked on, so nothing is
# hardcoded and nothing is asked. A fixture always works on a throwaway copy.
#
# WHERE THE PASSWORD LIVES, AND WHY NOT A CELL ADDRESS
# -----------------------------------------------------------------------------
# It is the WORKBOOK-LEVEL NAME "RNG_DebuggingPassword", which is what the
# product itself reads: Passwords.cls holds it as NAMEDEBUGPASSWORD and gets at
# it through ThisWorkbook.Names. The name lives on the __pass sheet but its cell
# is not fixed, so an address is the wrong handle for it.
#
# This file read __pass!A1 until 2026-11-25. That address happens to hold the
# password in the OLD setups, which is why case 060 lifted Translations on
# ebola-2026-11-19.xlsb and nobody noticed. On the CURRENT setups A1 holds
# something else, both lift attempts failed, and case 030 skipped with
# "the Translations sheet is protected and would not unprotect" -- a fixture
# fault that read like a product one. Follow the name, never the address.

Set-StrictMode -Off

# Get-DebugPassword -Workbook <wb>
#
# The workbook's own debugging password, or $null when it carries none.
#
# Answers a STRING or $null and prints nothing, so a caller can try it without
# reporting it. Never raises: a workbook with no such name, or a name that
# refers nowhere, answers $null and the caller moves on to its next attempt.
#
# The workbook-level name comes first because that is where the product puts
# it. A sheet-scoped name is tried after it, for a workbook whose __pass sheet
# was cloned into a copy -- CloneToWorksheet does that -- since cloning can
# leave the name scoped to the sheet.
function Get-DebugPassword {
    param(
        [Parameter(Mandatory = $true)] $Workbook
    )

    $nameOfIt = "RNG_DebuggingPassword"

    # Cells(1,1) rather than Value2 on the range: a single-cell name answers the
    # same string either way, and a name that somehow spans cells answers the
    # first one instead of a 2-D array that would stringify into nonsense.
    try {
        $value = $Workbook.Names($nameOfIt).RefersToRange.Cells(1, 1).Value2
        if ($null -ne $value -and "$value".Length -gt 0) { return "$value" }
    } catch { }

    try {
        $sheet = $Workbook.Worksheets("__pass")
        $value = $sheet.Names($nameOfIt).RefersToRange.Cells(1, 1).Value2
        if ($null -ne $value -and "$value".Length -gt 0) { return "$value" }
    } catch { }

    return $null
}

# Lift-SheetProtection -Workbook <wb> -Sheet <ws>
#
# Answers $true when the sheet ends up writable, $false when it does not, and
# prompts under no circumstance. The order is the workbook's own password, then
# the empty string; both are explicit arguments.
#
# -Quiet suppresses the one line it otherwise prints about how it got in.
# EVERY MESSAGE GOES THROUGH Write-Host, DELIBERATELY.
# A bare string inside a PowerShell function is part of its OUTPUT STREAM,
# which is its return value. So a function that printed its progress and then
# returned $true answered an ARRAY, and `if (-not (Lift-SheetProtection ...))`
# tested the truthiness of a non-empty array -- always true, so the guard never
# fired and a refusal read as a success. Write-Host keeps the diagnostics on the
# host, where a redirected run still captures them, and leaves the return value
# a single Boolean.
function Lift-SheetProtection {
    param(
        [Parameter(Mandatory = $true)] $Workbook,
        [Parameter(Mandatory = $true)] $Sheet,
        [switch] $Quiet
    )

    $say = { param($text) if (-not $Quiet) { Write-Host $text } }

    if (-not $Sheet.ProtectContents) {
        & $say "unprotect: the sheet was not protected"
        return $true
    }

    $password = Get-DebugPassword $Workbook

    if ($null -ne $password -and "$password".Length -gt 0) {
        try {
            $Sheet.Unprotect("$password")
            if (-not $Sheet.ProtectContents) {
                & $say "unprotect: lifted with the password the workbook carries"
                return $true
            }
        } catch { }
    }

    # A sheet protected with no password at all. The empty string is still an
    # explicit argument, so this cannot prompt either.
    try {
        $Sheet.Unprotect("")
        if (-not $Sheet.ProtectContents) {
            & $say "unprotect: lifted with an empty password"
            return $true
        }
    } catch { }

    & $say "unprotect: REFUSED -- neither the workbook's own password nor an empty one lifted it"
    return $false
}

# New-QuietExcel: an Excel that asks nothing.
#
# AutomationSecurity is deliberately left alone. ForceDisable would disable the
# macros of a workbook being opened, and a fixture that later needs
# Application.Run would break on the one call that matters.
function New-QuietExcel {
    param([switch] $Visible)

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = [bool]$Visible
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    # A setup carries Workbook_Open and Workbook_BeforeClose, and a fixture run
    # has no screen for whatever they would put there.
    $xl.EnableEvents = $false
    return $xl
}

# Close-QuietExcel: let go of it, and of the COM reference.
function Close-QuietExcel {
    param($Excel)

    if ($null -eq $Excel) { return }
    try { $Excel.DisplayAlerts = $false; $Excel.Quit() } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) } catch { }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
