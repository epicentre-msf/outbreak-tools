-- build-linelist.applescript
-- ============================================================================
-- Thin macOS trigger for the headless linelist build. All portable logic lives
-- in scripts/headless/build-linelist.R; this only does what ONLY Apple Events
-- can do: open the driver copy, load the source into it, call the build, read
-- back what the build recorded, and quit Excel cleanly.
--
-- Invoked by build-linelist.R:
--   osascript build-linelist.applescript <workbook> <designer> <setup> \
--             <sourceRoot> <formsFolder> <outFolder> <outName> <options> \
--             <reportPath> <obtHome>
--
-- <obtHome> is the ONE folder every other path sits under, and the only one
-- the macOS sandbox is asked about. See the step 0 note below.
--
-- This trigger runs NO test. It calls HeadlessBuild.BuildLinelistFromSetup
-- directly and passes it the eight paths R resolved, which is why the build
-- stops being pinned to the absolute paths a test module carried as constants.
--
-- Order matters, and it is the same order the test harness uses for its own
-- reasons: OBTBuildCodeTables reads the run dir's .generated intermediates and
-- rebuilds the Codes import tables from them, then OBTSilentImport reads those
-- tables to decide what to pull from the staged source. R has already filtered
-- the intermediates down to the classes and modules alone, so what lands in the
-- driver here is the build closure with no test module in it.
--
-- The macros are referenced as "<workbook-name>!<macro>" so they resolve in the
-- opened copy regardless of what else is open, and the build itself is
-- module-qualified because BuildLinelistFromSetup is a name a source module
-- could one day share.
--
-- Prerequisite: build-linelist.R quits any running Excel before calling this,
-- so `open` always starts a fresh instance (run VB macro fails with a Parameter
-- error against an already-open Excel).
-- ============================================================================

on run argv
	if (count of argv) < 11 then
		error "build-linelist.applescript: expected 11 arguments, got " & (count of argv)
	end if

	set wbPath to item 1 of argv
	set designerPath to item 2 of argv
	set setupPath to item 3 of argv
	set sourceRoot to item 4 of argv
	set formsFolder to item 5 of argv
	set outFolder to item 6 of argv
	set outName to item 7 of argv
	set buildOptions to item 8 of argv
	set reportPath to item 9 of argv
	set obtHome to item 10 of argv
	set needPick to item 11 of argv

	set wbName to my basename(wbPath)

	set outcomeText to ""
	set summaryText to ""
	set grantText to ""

	tell application "Microsoft Excel"
		set display alerts to false

		-- Read/write: the run imports components into the copy.
		--
		-- The open carries its own timeout, and it is not decoration. On the
		-- FIRST run on a machine, the grant panel below can put itself on the
		-- screen and wait for a person. AppleScript's default timeout is two
		-- minutes; an unattended first run crossed it and reported "AppleEvent
		-- timed out (-1712)" against an Excel that was sitting on a dialog with
		-- nothing wrong with it. Ten minutes is long enough for somebody to
		-- notice and click Grant, and short enough that an unattended run still
		-- gives up the same day.
		with timeout of 600 seconds
			open workbook workbook file name wbPath read only false
		end timeout

		-- 0) THE FOLDER PICK, and it comes before every other call because every
		--    other call reads a file. OBTRefreshHarness below is itself a read
		--    of run/bootstrap/*.bas, and nothing used to cover that folder,
		--    which is why the harness modules asked on every single run.
		--
		--    One pick covers the whole run: the staged sources, the generation
		--    log, and the OBTApp_ files a run invents as it goes. Measured
		--    2026-08-14 and written up in .obt/gotchas/macos-sandbox-grant.md.
		--    Application.GrantAccessToMultipleFiles raises 438 on Excel for Mac
		--    and has never done anything; only a picked folder counts.
		--
		--    OBTGrantAccess is the macro that does the picking, and it is used
		--    here BECAUSE IT NEVER CHANGES. It has been in the driver workbook
		--    since the harness was set up. A purpose-built entry point in
		--    OBTBootstrap was tried first and failed in the field: OBTBootstrap
		--    is the one module nothing re-imports from disk, so changing its
		--    signature meant the workbook held the old copy, the call errored,
		--    and no picker ever opened. Calling a macro that is already there
		--    cannot go stale.
		--
		--    needPick is "1" only when R found no <obtHome>/.granted, so this is
		--    silent on every run after the first. R does that check because R is
		--    not sandboxed; the same check inside Excel would pop the panel it
		--    is trying to avoid.
		--
		--    Ten minutes: somebody has to pick the folder and clear the two
		--    message boxes OBTGrantAccess shows. Every run after that skips it.
		if needPick is "1" then
			with timeout of 600 seconds
				try
					run VB macro (wbName & "!OBTGrantAccess")
					set grantText to "folder picker was shown for " & obtHome
				on error errText number errNum
					set grantText to "PICKER FAILED " & (errNum as text) & ": " & errText
				end try
			end timeout
		else
			set grantText to "step 0 skipped, no folder picker was opened"
		end if

		-- One hour, the same ceiling the test trigger uses. A build is far
		-- shorter than a full suite, but the ceiling is there to stop a wedged
		-- modal hanging the driver forever rather than to time anything.
		with timeout of 3600 seconds
			-- 1) refresh the harness modules from the run dir, so OBTImport can
			--    be iterated without a manual VBE re-import.
			run VB macro (wbName & "!OBTRefreshHarness")

			-- 2) rebuild the Codes tables from the filtered intermediates. Always,
			--    never optional: the driver may be holding the tables a test run
			--    left behind, and those carry the test folders too.
			run VB macro (wbName & "!OBTBuildCodeTables")

			-- 3) pull the staged source into the driver
			run VB macro (wbName & "!OBTSilentImport")

			-- 4) the build. Wrapped, so a raise still leaves a report on disk
			--    saying how far it got rather than only an osascript exit code.
			--
			--    obtHome goes in as arg8: the build creates its own scratch tree
			--    under the output folder and re-grants the root before doing so,
			--    which costs nothing once step 0 has already granted it.
			try
				set outcomeText to (run VB macro (wbName & "!HeadlessBuild.BuildLinelistFromSetup") arg1 designerPath arg2 setupPath arg3 sourceRoot arg4 formsFolder arg5 outFolder arg6 outName arg7 buildOptions arg8 obtHome) as text
			on error errText number errNum
				set outcomeText to "TRIGGER ERROR " & (errNum as text) & ": " & errText
			end try

			-- 5) what the build recorded. Read it whatever the outcome was: a
			--    run that stopped halfway still holds the narrative up to the
			--    stop, and that narrative is the whole diagnostic.
			--
			--    One call, because the six things worth reading are exposed as
			--    Property Get and Application.Run cannot reach a Property Get.
			--    HeadlessBuild.LastBuildSummary is the Function that can.
			try
				set summaryText to (run VB macro (wbName & "!HeadlessBuild.LastBuildSummary")) as text
			end try
		end timeout

		-- Quit from AppleScript, never from VBA. Nothing here needs the driver
		-- saved: the build writes its own files and the driver is a copy.
		quit saving no
	end tell

	-- The outcome is the one field the summary cannot carry: it is what the
	-- build ANSWERED, and a build that raised never returned to say it. The
	-- grant line is beside it because a run that started prompting again is a
	-- run whose grant did not take, and that has to be readable off the report
	-- rather than guessed at from how long the run took.
	set reportText to "outcome=" & outcomeText & linefeed & ¬
		"grant=" & grantText & linefeed & summaryText & linefeed

	my writeReport(reportPath, reportText)
end run

-- Truncate and rewrite. The file is created by R before the run and kept from
-- one run to the next, for the same reason the test harness keeps its CSV: a
-- file Excel or a script creates fresh has no macOS access grant, and one that
-- is emptied in place keeps the identity the grant is tied to.
on writeReport(reportPath, reportText)
	set fileRef to open for access (POSIX file reportPath) with write permission
	try
		set eof fileRef to 0
		write reportText to fileRef as «class utf8»
	end try
	close access fileRef
end writeReport

-- basename without threading in shell calls (keeps the trigger self-contained).
on basename(p)
	set AppleScript's text item delimiters to "/"
	set parts to text items of p
	set AppleScript's text item delimiters to ""
	return item -1 of parts
end basename
