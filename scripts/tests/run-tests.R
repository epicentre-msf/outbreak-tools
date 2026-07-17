# run-tests.R
# =============================================================================
# Orchestrator for the automated VBA test harness (macOS-driven). One command:
#
#   Rscript scripts/tests/run-tests.R [--build] [--keep]
#
#   --build   also rebuild the Codes tables + ModulesForTesting from the
#             registry before running (runs build-registry.R, then asks the
#             workbook to rebuild via OBT_BuildCodeTables). Omit to run the
#             tables/Name already in the workbook.
#   --keep    never delete the per-run working copy, even on success
#             (for inspecting a green run).
#
# What it does (all portable logic lives here; the OS-specific trigger is thin):
#   1. copy src/bin/test-files/unit_tests_dev.xlsb -> a fresh per-run dir under
#      src/tests/.generated/run-<stamp>/ (NEVER touch the original).
#   2. (optional) regenerate the registry intermediates.
#   3. shell out to the macOS trigger, which opens the copy, runs the OBT_*
#      macros, and quits Excel.
#   4. read test-results.csv written next to the copy; summarise.
#   5. success (csv present, 0 failures) -> delete the run dir.
#      failure -> KEEP the run dir and print the stale-copy path so it can be
#      opened by hand; the stale .xlsb holds whatever partial testsOutputs Excel
#      managed even when the CSV never printed. Ask the user to test by hand
#      only as a last resort (Excel wedged / no CSV / no usable stale state).
#   Exit code: 0 on all-pass, 1 on any failure or incomplete run.
#
# Cross-platform: the copy/collect/summarise logic is OS-agnostic. Only the
# trigger differs — macos/run-tests.applescript today; windows/run-tests.vbs
# (trigger-file watcher) is Phase D. See .obt/plans/test-scripts-status.md.
#
# NOTE: not run yet. The OBT_* VBA entry points (Phase A/B) are still to be
# built; until they exist this script will fail at the macro-run step. That is
# expected and documented.
# =============================================================================

# --- args --------------------------------------------------------------------
args     <- commandArgs(trailingOnly = TRUE)
do_build <- "--build" %in% args
do_keep  <- "--keep"  %in% args

# --- locate repo root + key paths --------------------------------------------
repo_root <- tryCatch(
  system2("git", c("rev-parse", "--show-toplevel"), stdout = TRUE),
  warning = function(w) NA_character_
)
if (length(repo_root) != 1L || is.na(repo_root) || !nzchar(repo_root)) {
  stop("run-tests.R: not inside a git repository (could not resolve repo root).")
}

workbook_src <- file.path(repo_root, "src", "bin", "test-files", "unit_tests_dev.xlsb")
scripts_dir  <- file.path(repo_root, "scripts", "tests")
trigger      <- file.path(scripts_dir, "macos", "run-tests.applescript")
generated    <- file.path(repo_root, "src", "tests", ".generated")

if (Sys.info()[["sysname"]] != "Darwin") {
  stop("run-tests.R: the macOS trigger is the only one implemented. ",
       "Windows parity is Phase D (windows/run-tests.vbs).")
}
if (!file.exists(workbook_src)) stop("run-tests.R: workbook not found: ", workbook_src)
if (!file.exists(trigger))      stop("run-tests.R: trigger not found: ", trigger)

# --- optional: refresh registry intermediates --------------------------------
if (do_build) {
  message("run-tests.R: rebuilding registry intermediates ...")
  rc <- system2("Rscript", file.path(scripts_dir, "build-registry.R"))
  if (rc != 0L) stop("run-tests.R: build-registry.R failed (exit ", rc, ").")
}

# --- 1) per-run working copy -------------------------------------------------
stamp   <- format(Sys.time(), "%Y%m%d-%H%M%S")
run_dir <- file.path(generated, paste0("run-", stamp))
dir.create(run_dir, recursive = TRUE, showWarnings = FALSE)

work_copy <- file.path(run_dir, "unit_tests_run.xlsb")
if (!file.copy(workbook_src, work_copy, overwrite = TRUE)) {
  stop("run-tests.R: failed to copy workbook into ", run_dir)
}
csv_path <- file.path(run_dir, "test-results.csv")   # written by OBT_RunAllTests
message("run-tests.R: working copy -> ", work_copy)

# --- 2/3) drive Excel via the thin trigger -----------------------------------
build_flag <- if (do_build) "build" else "nobuild"
message("run-tests.R: launching Excel via osascript (", build_flag, ") ...")
trigger_rc <- system2("osascript", c(shQuote(trigger), shQuote(work_copy), build_flag))

# --- 4) collect + summarise --------------------------------------------------
# Helper: keep the stale copy, explain, and exit non-zero.
fail <- function(msg) {
  message("\n[FAIL] ", msg)
  message("       Stale working copy kept for inspection:")
  message("       ", work_copy)
  message("       Open it in Excel to read whatever `testsOutputs` was written,")
  message("       or re-run with --build if the Codes tables look stale.")
  message("       Ask for a manual test run only as a last resort.")
  quit(status = 1L, save = "no")
}

if (trigger_rc != 0L) {
  fail(sprintf("osascript trigger returned %d (Excel may have wedged on a dialog / timed out).", trigger_rc))
}
if (!file.exists(csv_path)) {
  fail("no test-results.csv was produced (the run likely died before serialising).")
}

results <- tryCatch(
  read.csv(csv_path, stringsAsFactors = FALSE, check.names = FALSE),
  error = function(e) fail(paste0("test-results.csv is unreadable: ", conditionMessage(e)))
)

# Expected columns (see plan Phase B SerializeTestOutputs): module,title,type,label,message
type_col <- intersect(c("type", "Type"), names(results))
if (!length(type_col)) {
  fail("test-results.csv has no `type` column — serializer format changed?")
}
types    <- tolower(as.character(results[[type_col[1]]]))
failures <- sum(types == "error", na.rm = TRUE)
passes   <- sum(types == "success", na.rm = TRUE)

message(sprintf("\nrun-tests.R: %d success / %d failure row(s) across %d result line(s).",
                passes, failures, nrow(results)))

if (failures > 0L) {
  # Show the failing rows so they can be worked directly from here.
  mod_col <- intersect(c("module", "Module"), names(results))
  msg_col <- intersect(c("message", "Message"), names(results))
  bad <- results[types == "error", , drop = FALSE]
  message("\nFailing rows:")
  for (i in seq_len(nrow(bad))) {
    m  <- if (length(mod_col)) bad[[mod_col[1]]][i] else "?"
    tx <- if (length(msg_col)) bad[[msg_col[1]]][i] else ""
    message(sprintf("  - [%s] %s", m, tx))
  }
  fail(sprintf("%d test failure(s). CSV: %s", failures, csv_path))
}

# --- 5) success: copy the CSV out, then clean up -----------------------------
final_csv <- file.path(generated, "test-results.csv")
file.copy(csv_path, final_csv, overwrite = TRUE)
message("run-tests.R: latest results -> ", final_csv)

if (do_keep) {
  message("run-tests.R: --keep set; leaving working copy at ", run_dir)
} else {
  unlink(run_dir, recursive = TRUE, force = TRUE)
  message("run-tests.R: all tests passed; removed working copy.")
}
quit(status = 0L, save = "no")
