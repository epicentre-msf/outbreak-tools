-- run-tests.applescript
-- ============================================================================
-- Thin macOS trigger for the automated test harness. All portable logic lives
-- in scripts/tests/run-tests.R; this only does what ONLY Apple Events can do:
-- open the workbook copy, run its parameterless OBT* macros in order, and quit
-- Excel cleanly (quit is AppleScript's job, never VBA's).
--
-- Invoked by run-tests.R:
--   osascript run-tests.applescript <workbook-copy-path> <build|nobuild>
--
--   arg 1  absolute path to the per-run workbook COPY (never the original)
--   arg 2  "build"  -> also run OBTBuildCodeTables FIRST, rebuilding the Codes
--                       import tables + ModulesForTesting from the registry
--                       intermediates in src/tests/.generated
--          "nobuild" -> skip the rebuild; import + run against the tables the
--                       workbook already holds
--
-- Order matters: OBTBuildCodeTables must run BEFORE OBTSilentImport, because the
-- silent import reads the freshly rebuilt Codes tables to decide what to pull
-- from src/. The chain is therefore build (optional) -> import -> run.
--
-- The macros are referenced as "<workbook-name>!<macro>" so they resolve in the
-- opened copy regardless of what else is open. OBTRunAllTests writes
-- test-results.csv next to the workbook and Saves before returning; R reads it.
--
-- Prerequisite: the workbook must carry the Development manager + the OBTImport
-- and OBTHeadless modules, and the Codes-sheet folder named ranges
-- (ModulesCodes / ClassesImplementation / TestsCodes). run-tests.R quits any
-- running Excel before calling this, so `open` always starts a fresh instance
-- (run VB macro fails with a Parameter error against an already-open Excel).
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

	-- Derive the workbook name used to qualify the macro references below.
	set wbName to my basename(wbPath)

	tell application "Microsoft Excel"
		-- Suppress alerts at the app level as a belt-and-braces guard; the VBA
		-- entry points also set DisplayAlerts = False on their own path.
		set display alerts to false

		-- Open read/write explicitly: the loop MUTATES the workbook (imports
		-- components via Development, rebuilds the Codes tables, writes results)
		-- and Saves before returning, so a read-only open would break the run.
		open workbook workbook file name wbPath read only false

		-- Each run is wrapped in a timeout so a wedged modal cannot hang the
		-- driver forever (surfaces as AppleEvent timed out -1712 to R).
		with timeout of 600 seconds
			-- 0) refresh the harness modules (OBTImport/OBTHeadless) from the run
			--    dir, so their code can be iterated without a manual VBE re-import.
			--    Separate call, so the refreshed code is loaded for the steps below.
			run VB macro (wbName & "!OBTRefreshHarness")

			-- 1) optionally rebuild Codes tables + ModulesForTesting from registry
			if doBuild then
				run VB macro (wbName & "!OBTBuildCodeTables")
			end if

			-- 2) refresh workbook code from src/ (no dialogs) via Development
			run VB macro (wbName & "!OBTSilentImport")

			-- 3) run every registered module, serialize testsOutputs to CSV, Save
			run VB macro (wbName & "!OBTRunAllTests")
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
