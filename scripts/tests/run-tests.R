# run-tests.R
# =============================================================================
# Orchestrator for the automated VBA test harness (macOS-driven). One command:
#
#   Rscript scripts/tests/run-tests.R [--build] [--keep] [--home=<dir>]
#
#   --build   also rebuild the Codes tables + ModulesForTesting from the
#             registry before running (runs build-registry.R, then asks the
#             workbook to rebuild via OBTBuildCodeTables). Required on the first
#             run and after any registry change; omit to reuse the tables the
#             workbook already holds.
#   --keep    never delete the per-run working copy, even on success
#             (for inspecting a green run).
#   --home    root of the untracked working area holding the driver workbook
#             (unit_tests_dev.xlsb) + the per-run staging tree. Defaults to the
#             OBT_TEST_HOME env var, else <repo>/.test-runner. Gitignored, so its
#             location is a local setting, not fixed in the repo.
#
# The harness + probes now live in src/ (promoted). This script assembles a
# per-run dir under <home>/tests/staging/run/ from src/ (the workbook is the only
# thing kept in the working area, as an untracked binary). The repo-root folder
# grant (OBTGrantAccess, one-time) lets Excel read it all with no per-file prompts.
#
# What it does (all portable logic lives here; the OS-specific trigger is thin):
#   1. copy <home>/unit_tests_dev.xlsb -> the stable run dir (NEVER touch original)
#      and assemble the registered probe sources + manifest + harness from src/.
#   2. (optional, --build) regenerate the registry intermediates (src/tests/.generated).
#   3. quit any running Excel (run VB macro fails against an already-open one),
#      then shell out to the macOS trigger, which opens the copy fresh, runs the
#      OBT* macros (refresh -> build -> import -> run), and quits Excel.
#   4. read test-results.csv written next to the copy; summarise.
#   5. success (csv present, 0 failures) -> delete the run dir.
#      failure -> KEEP the run dir and print the stale-copy path so it can be
#      opened by hand; the stale .xlsb holds whatever partial testsOutputs Excel
#      managed even when the CSV never printed. Ask the user to test by hand
#      only as a last resort (Excel wedged / no CSV / no usable stale state).
#   Exit code: 0 on all-pass, 1 on any failure or incomplete run.
#
# Cross-platform: the copy/collect/summarise logic is OS-agnostic. Only the
# trigger differs — macos/run-tests.applescript today. A Windows trigger
# (windows/run-tests.vbs, a trigger-file watcher) is not wired in here yet.
#
# The workbook must carry the Development manager + the OBTImport / OBTHeadless
# modules and the Codes-sheet folder named ranges (ModulesCodes /
# ClassesImplementation / TestsCodes). See src/tests/automated-testing-macos.md.
# =============================================================================

# --- args --------------------------------------------------------------------
args     <- commandArgs(trailingOnly = TRUE)
do_build <- "--build" %in% args
do_keep  <- "--keep"  %in% args
home_opt <- sub("^--home=", "", grep("^--home=", args, value = TRUE))  # "" when absent

# --- locate repo root + key paths --------------------------------------------
repo_root <- tryCatch(
  system2("git", c("rev-parse", "--show-toplevel"), stdout = TRUE),
  warning = function(w) NA_character_
)
if (length(repo_root) != 1L || is.na(repo_root) || !nzchar(repo_root)) {
  stop("run-tests.R: not inside a git repository (could not resolve repo root).")
}

# The untracked working area (driver workbook + per-run staging). It is gitignored,
# so where it lives is a local choice, not a repo fact: --home=<dir> wins, else the
# OBT_TEST_HOME env var, else <repo>/.test-runner. A relative home resolves against the
# repo root, so the default and a relative --home behave the same from any cwd.
test_home <- if (length(home_opt) && nzchar(home_opt[1])) {
  home_opt[1]
} else if (nzchar(Sys.getenv("OBT_TEST_HOME"))) {
  Sys.getenv("OBT_TEST_HOME")
} else {
  file.path(repo_root, ".test-runner")
}
if (!grepl("^(/|~|[A-Za-z]:)", test_home)) test_home <- file.path(repo_root, test_home)

workbook_src <- file.path(test_home, "unit_tests_dev.xlsb")
scripts_dir  <- file.path(repo_root, "scripts", "tests")
trigger      <- file.path(scripts_dir, "macos", "run-tests.applescript")
generated    <- file.path(repo_root, "src", "tests", ".generated")     # build-registry.R output
staging      <- file.path(test_home, "tests", "staging")                # untracked staging tree
# STABLE run dir (fixed path, cleared+reused each run). macOS grants Excel file
# access per PATH, so a fixed path is prompted for at most once, then never again
# -- a fresh run-<stamp> dir each time is what caused the repeated grant prompts.
run_dir      <- file.path(staging, "run")

if (!dir.exists(staging)) {
  stop("run-tests.R: staging tree not found: ", staging,
       " (set --home / OBT_TEST_HOME, or create <home>/tests/staging).")
}

if (Sys.info()[["sysname"]] != "Darwin") {
  stop("run-tests.R: the macOS trigger is the only one implemented. ",
       "The Windows trigger (windows/run-tests.vbs) is not wired in yet.")
}
if (!file.exists(workbook_src)) stop("run-tests.R: workbook not found: ", workbook_src)
if (!file.exists(trigger))      stop("run-tests.R: trigger not found: ", trigger)

# --- optional: refresh registry intermediates --------------------------------
# build-registry.R writes the canonical manifest under src/tests/.generated; it
# is copied into the run dir (next to the workbook) below, where OBTImport reads
# it. Only needed on --build, which is also when OBTBuildCodeTables runs.
if (do_build) {
  message("run-tests.R: rebuilding registry intermediates ...")
  rc <- system2("Rscript", file.path(scripts_dir, "build-registry.R"))
  if (rc != 0L) stop("run-tests.R: build-registry.R failed (exit ", rc, ").")
}

# --- 1) per-run working copy -------------------------------------------------
# Keep the run dir at a STABLE identity (same folder, not just same path) every
# run. A macOS folder-access grant is tied to the folder's identity (inode), so
# DELETING and recreating it -- even at the same path -- makes a new folder the
# grant no longer covers, and Excel re-prompts. So clear the CONTENTS and keep
# the folder itself. (The grant should be on the repo root, a stable parent.)
if (dir.exists(run_dir)) {
  unlink(list.files(run_dir, all.files = TRUE, full.names = TRUE, no.. = TRUE),
         recursive = TRUE, force = TRUE)
} else {
  dir.create(run_dir, recursive = TRUE, showWarnings = FALSE)
}

work_copy <- file.path(run_dir, "unit_tests_run.xlsb")
if (!file.copy(workbook_src, work_copy, overwrite = TRUE)) {
  stop("run-tests.R: failed to copy workbook into ", run_dir)
}
csv_path <- file.path(run_dir, "test-results.csv")   # written by OBTRunAllTests
message("run-tests.R: working copy -> ", work_copy)

# Assemble a SELF-CONTAINED run dir next to the workbook copy: the probe sources
# (pulled from src/, only the registered folders) + the manifest + the harness
# sources. Excel reads them from ThisWorkbook.Path (the run dir), which the
# repo-root folder grant covers -> no per-file prompts.
#
# Map each Development tag to (candidate source bases, run-dir base). Real suites
# live in src/; the local `draft` PROBE is test-only and lives untracked under
# <home>/tests/staging -- so each folder resolves from src first, then falls back
# to the staging area. Then copy <base>/<folder> -> run/<runbase>/<folder>.
tag_srcs <- list(
  "general classes" = c(file.path(repo_root, "src", "classes"),
                        file.path(staging, "classes")),
  "general modules" = c(file.path(repo_root, "src", "modules"),
                        file.path(staging, "modules")),
  # Test modules AND fixture classes now share one per-suite folder, src/tests/<folder>
  # (the old tests/modules + tests/classes split was retired). Both tags therefore
  # resolve to the same base; scope (module vs class) is still decided by the tag.
  "tests modules"   = c(file.path(repo_root, "src", "tests"),
                        file.path(staging, "tests")),
  "tests classes"   = c(file.path(repo_root, "src", "tests"),
                        file.path(staging, "tests"))
)
tag_runs <- list(
  "general classes" = file.path(run_dir, "classes"),
  "general modules" = file.path(run_dir, "modules"),
  "tests modules"   = file.path(run_dir, "tests"),
  "tests classes"   = file.path(run_dir, "tests")
)
tbl <- read.delim(file.path(generated, "code-tables.tsv"), stringsAsFactors = FALSE)
pairs <- unique(tbl[, c("folder", "tag")])
for (i in seq_len(nrow(pairs))) {
  cands   <- tag_srcs[[pairs$tag[i]]]
  run_base <- tag_runs[[pairs$tag[i]]]
  if (is.null(cands)) next
  found <- NULL
  for (b in cands) {
    if (dir.exists(file.path(b, pairs$folder[i]))) { found <- file.path(b, pairs$folder[i]); break }
  }
  if (is.null(found)) {
    message("run-tests.R: WARNING - no source found for folder '", pairs$folder[i],
            "' (tag '", pairs$tag[i], "') in src/ or the staging area; skipping.")
    next
  }
  dir.create(run_base, recursive = TRUE, showWarnings = FALSE)
  file.copy(found, run_base, recursive = TRUE)
}

dir.create(file.path(run_dir, ".generated"), showWarnings = FALSE)
for (f in c("code-tables.tsv", "modules-for-testing.txt")) {  # manifest (built above)
  src_f <- file.path(generated, f)
  if (file.exists(src_f)) file.copy(src_f, file.path(run_dir, ".generated", f), overwrite = TRUE)
}
# Harness sources for OBTRefreshHarness to re-import at run start (so OBTImport /
# OBTHeadless load the current src/ version without a manual VBE re-import).
dir.create(file.path(run_dir, "bootstrap"), showWarnings = FALSE)
for (f in c("OBTImport.bas", "OBTHeadless.bas")) {
  src_f <- file.path(repo_root, "src", "tests", "rubberduck", f)
  if (file.exists(src_f)) file.copy(src_f, file.path(run_dir, "bootstrap", f), overwrite = TRUE)
}
message("run-tests.R: run dir assembled from src/ (classes/, tests/, .generated/, bootstrap/)")

# --- 2a) ensure Excel is fully closed before we launch a fresh instance ------
# `run VB macro` returns AppleScript Parameter error (-50) against an Excel that
# is already open interactively, so quit any running instance and wait until the
# process is gone. Only quit when it is actually running (a bare `tell ... quit`
# would auto-launch Excel just to close it).
excel_running <- function() {
  identical(system2("pgrep", c("-x", "Microsoft Excel"),
                    stdout = FALSE, stderr = FALSE), 0L)
}
if (excel_running()) {
  message("run-tests.R: quitting the running Excel instance ...")
  system2("osascript",
          c("-e", shQuote('tell application "Microsoft Excel" to quit saving no')),
          stdout = FALSE, stderr = FALSE)
  for (i in seq_len(40)) {          # up to ~20s
    if (!excel_running()) break
    Sys.sleep(0.5)
  }
  if (excel_running()) {
    stop("run-tests.R: Excel would not quit; close it by hand before re-running.")
  }
}

# --- 2b/3) drive Excel via the thin trigger ----------------------------------
build_flag <- if (do_build) "build" else "nobuild"
message("run-tests.R: launching Excel via osascript (", build_flag, ") ...")
trigger_rc <- system2("osascript", c(shQuote(trigger), shQuote(work_copy), build_flag))

# --- 4) collect + summarise --------------------------------------------------
# Helper: dump the VBA-side diagnostics log if the wrappers wrote one.
show_import_log <- function() {
  log_path <- file.path(run_dir, "obt-import.log")
  if (file.exists(log_path)) {
    message("\n       --- obt-import.log (OBTBuildCodeTables/OBTSilentImport) ---")
    for (ln in readLines(log_path, warn = FALSE)) message("       ", ln)
    message("       --- end log ---")
  } else {
    message("\n       (no obt-import.log written — the wrappers may not have run, ",
            "or could not write to ", run_dir, ")")
  }
}

# Helper: keep the stale copy, explain, and exit non-zero.
fail <- function(msg) {
  message("\n[FAIL] ", msg)
  show_import_log()
  message("\n       Stale working copy kept for inspection:")
  message("       ", work_copy)
  message("       Open it in Excel to read whatever `testsOutputs` was written,")
  message("       or re-run with --build if the Codes tables look stale.")
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

# Expected serializer columns (from OBTRunAllTests): module,title,type,label,message
type_col <- intersect(c("type", "Type"), names(results))
if (!length(type_col)) {
  fail("test-results.csv has no `type` column — serializer format changed?")
}
types    <- tolower(as.character(results[[type_col[1]]]))
failures <- sum(types == "error", na.rm = TRUE)
passes   <- sum(types == "success", na.rm = TRUE)

# Surface the serializer's __summary__ diagnostics (modulesRun / testsFound /
# vbe / name) whatever the outcome — they explain an empty run at a glance.
mod_col <- intersect(c("module", "Module"), names(results))
if (length(mod_col)) {
  summ <- results[results[[mod_col[1]]] == "__summary__", , drop = FALSE]
  if (nrow(summ)) {
    lab_col <- intersect(c("label", "Label"), names(results))
    msg_col <- intersect(c("message", "Message"), names(results))
    message("run-tests.R: summary — ",
            if (length(lab_col)) summ[[lab_col[1]]][1] else "",
            " | ",
            if (length(msg_col)) summ[[msg_col[1]]][1] else "")
  }
}

message(sprintf("\nrun-tests.R: %d success / %d failure row(s) across %d result line(s).",
                passes, failures, nrow(results)))

# An all-zero run means no test actually executed (only the summary row was
# written) — that is a FAILURE, not a pass. Do not let it clean up green.
if (passes + failures == 0L) {
  fail("no test result rows were produced — 0 tests ran. See summary + log above.")
}

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
final_csv <- file.path(staging, "test-results.csv")
file.copy(csv_path, final_csv, overwrite = TRUE)
message("run-tests.R: latest results -> ", final_csv)

# Leave the stable run dir in place (it is cleared at the start of the next run).
# Keeping the same path is what avoids re-triggering the macOS file-access grant.
message("run-tests.R: all tests passed. Run dir left at ", run_dir)
quit(status = 0L, save = "no")
