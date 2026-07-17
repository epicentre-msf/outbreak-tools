-- run-tests.applescript
-- ============================================================================
-- Thin macOS trigger for the test harness. All portable logic lives in
-- scripts/tests/run-tests.R; this only does what ONLY Apple Events can do:
-- drive an already-open Excel workbook, run its parameterless OBT_* macros,
-- and quit Excel cleanly (quit is AppleScript's job, never VBA's).
--
-- Invoked by run-tests.R:
--   osascript run-tests.applescript <workbook-copy-path> <build|nobuild>
--
--   arg 1  absolute path to the per-run workbook COPY (never the original)
--   arg 2  "build"  -> also run OBT_BuildCodeTables (rebuild Codes tables
--                       + ModulesForTesting from the registry .generated files)
--          "nobuild" -> skip the rebuild, just import + run
--
-- The macros are referenced as "<workbook-name>!<macro>" so they resolve in the
-- opened copy regardless of what else is open. OBT_RunAllTests writes
-- test-results.csv next to the workbook and Saves before returning; R reads it.
--
-- Prerequisite (Phase A/B, not yet built): the OBT_* entry points must exist in
-- the workbook. See .obt/plans/test-scripts-status.md.
-- ============================================================================

on run argv
	if (count of argv) < 1 then
		error "run-tests.applescript: missing workbook path argument."
	end if

	set wbPath to item 1 of argv
	set doBuild to false
	if (count of argv) >= 2 then
		if item 2 of argv is "build" then set doBuild to true
	end if

	-- POSIX path -> HFS path for Excel's `open`, and derive the workbook name.
	set wbFile to POSIX file wbPath
	set wbName to my basename(wbPath)

	tell application "Microsoft Excel"
		-- Suppress alerts at the app level as a belt-and-braces guard; the VBA
		-- entry points also set DisplayAlerts = False on their own path.
		set display alerts to false

		open wbFile

		-- Each run is wrapped in a timeout so a wedged modal cannot hang the
		-- driver forever (surfaces as AppleEvent timed out -1712 to R).
		with timeout of 600 seconds
			-- 1) refresh workbook code from src/ (no dialogs)
			run VB macro (wbName & "!OBT_SilentImport")

			-- 2) optionally rebuild Codes tables + ModulesForTesting from registry
			if doBuild then
				run VB macro (wbName & "!OBT_BuildCodeTables")
			end if

			-- 3) run every registered module, serialize testsOutputs to CSV, Save
			run VB macro (wbName & "!OBT_RunAllTests")
		end timeout

		-- Quit from AppleScript (chosen): VBA has already Saved, so no prompt.
		quit saving no
	end tell
end run

-- basename without threading in shell calls (keeps the trigger self-contained).
on basename(p)
	set AppleScript's text item delimiters to "/"
	set parts to text items of p
	set AppleScript's text item delimiters to ""
	return item -1 of parts
end basename
