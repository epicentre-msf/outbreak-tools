-- rundesigner.applescript
-- ============================================================================
-- macOS twin of rundesigner.vbs: drive one linelist generation on a designer
-- workbook. The script writes every entry the Main sheet needs, then runs the
-- generation callback. Excel stays visible: clickGenerate ends with a message
-- box that asks for one click.
--
--   osascript scripts/devtools/rundesigner.applescript \
--     <designer-path> <geo-path> <setup-path> <ll-dir> <ll-name> \
--     <setup-lang> <ll-lang> <ribbon-path>
--
-- The eight arguments are the same, in the same order, as rundesigner.vbs.
-- Paths are absolute POSIX paths.
-- ============================================================================

on run argv
	if (count of argv) < 8 then
		error "rundesigner.applescript: expected 8 arguments (designer, geo, setup, ll dir, ll name, setup lang, ll lang, ribbon)."
	end if

	set desPath to item 1 of argv
	set geoPath to item 2 of argv
	set setupPath to item 3 of argv
	set llDir to item 4 of argv
	set llName to item 5 of argv
	set setupLang to item 6 of argv
	set llLang to item 7 of argv
	set ribbonPath to item 8 of argv

	set wbName to my basename(desPath)

	tell application "Microsoft Excel"
		activate
		set display alerts to false

		open workbook workbook file name desPath read only false

		tell sheet "Main" of workbook wbName
			set value of range "RNG_PathDico" to setupPath
			set value of range "RNG_PathGeo" to geoPath
			set value of range "RNG_LLDir" to llDir
			set value of range "RNG_LLName" to llName
			set value of range "RNG_LLForm" to llLang
			set value of range "RNG_LLTemp" to ribbonPath
			set value of range "RNG_LangSetup" to setupLang
		end tell

		-- A full generation on a large setup can take many minutes; the cap
		-- only stops a wedged modal from holding the driver forever.
		with timeout of 3600 seconds
			run VB macro (wbName & "!clickGenerate")
		end timeout

		close workbook wbName saving no
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
