# build-linelist.R
# =============================================================================
# Generate one linelist, headless. One command, and it runs no test:
#
#   Rscript scripts/headless/build-linelist.R --setup=<filled-setup.xlsb> \
#           [--out=<dir>] [--name=<stem>] [--designer=<path>] \
#           [--forms=<dir>] [--home=<dir>] [--no-merge] \
#           [--temppath=<ribbon.xlsb>] [--geopath=<geo.xlsb>] \
#           [--setuplang=<column>] [--lllang=<code>] [--llpassword=<pass>]
#
# This is the build lifted out of the test harness. It used to happen only as a
# side effect of running TestHeadlessLinelistBuild, which meant generating a
# file took a test run, and the paths it generated from were constants compiled
# into a test module -- absolute, and true on one machine. Here every path is an
# argument, nothing is asserted, and the exit code says whether a linelist was
# written.
#
#   Exit 0  the build answered OK and the file is on disk
#   Exit 1  anything else, with the build's own narrative printed
#
# WHAT IT SHARES WITH run-tests.R, AND WHY
# -----------------------------------------------------------------------------
# The driver workbook has to hold the whole class closure BuildLinelistFromSetup
# runs on -- LinelistSpecs, Linelist, CodeTransfer, TemporaryRepos,
# ApplicationState, InitTransfer and everything they type. That closure is
# already described, once, in src/tests/test-registry.yml. So this reads the
# same registry rather than keeping a second list of it: a second list drifts,
# and a drifted list does not fail -- it delivers last month's code in a file
# that looks right.
#
# It reads the `classes:` and `modules:` rows and DROPS the test rows, so no
# test module is imported into the driver at all. Narrowing the registry for a
# probe only ever comments test rows, so a narrowed registry still builds.
#
# THE THREE STEPS, IN ORDER
# -----------------------------------------------------------------------------
#   1. merge-form-code.R, so the exported forms carry the current FormLogic
#      code. Skipped with --no-merge, and skipping it is how a linelist ends up
#      with the right buttons wired to last week's handlers.
#   2. build-registry.R, then the filtered intermediates into the run dir.
#   3. Excel: refresh the harness, rebuild the Codes tables, import the source,
#      call the build, read back what it recorded.
#
# The run dir is a STABLE path that is cleared and reused, and the workbook copy
# inside it is overwritten in place. macOS ties a file-access grant to the file
# identity rather than to the path, so sweeping and recreating hands Excel a
# brand new file the grant no longer covers and the operator is asked again.
# =============================================================================

# --- args --------------------------------------------------------------------
args <- commandArgs(trailingOnly = TRUE)

opt <- function(key, default = NULL) {
  hit <- grep(paste0("^--", key, "="), args, value = TRUE)
  if (!length(hit)) return(default)
  sub(paste0("^--", key, "="), "", hit[1])
}
flag <- function(key) paste0("--", key) %in% args

do_merge <- !flag("no-merge")

# --- locate repo root + the working area -------------------------------------
repo_root <- tryCatch(
  system2("git", c("rev-parse", "--show-toplevel"), stdout = TRUE),
  warning = function(w) NA_character_
)
if (length(repo_root) != 1L || is.na(repo_root) || !nzchar(repo_root)) {
  stop("build-linelist.R: not inside a git repository (could not resolve repo root).")
}

# Same untracked working area the test harness uses, and for the same reason:
# where it lives is a local choice, not a repo fact.
build_home <- opt("home", default = {
  if (nzchar(Sys.getenv("OBT_TEST_HOME"))) Sys.getenv("OBT_TEST_HOME")
  else file.path(repo_root, ".test-runner")
})
if (!grepl("^(/|~|[A-Za-z]:)", build_home)) build_home <- file.path(repo_root, build_home)

workbook_src <- file.path(build_home, "unit_tests_dev.xlsb")
scripts_dir  <- file.path(repo_root, "scripts")

# The trigger is the only OS-specific piece. Everything above and below it is
# the same on both hosts, which is the whole point of keeping it thin: the two
# triggers drive the SAME entry points in the SAME order with the SAME
# arguments, so a difference in what gets built is a difference in the VBA
# rather than in the harness.
on_windows <- identical(.Platform$OS.type, "windows")
trigger <- if (on_windows) {
  file.path(scripts_dir, "headless", "windows", "build-linelist.vbs")
} else {
  file.path(scripts_dir, "headless", "macos", "build-linelist.applescript")
}
merger       <- file.path(scripts_dir, "headless", "merge-form-code.R")
registry_r   <- file.path(scripts_dir, "tests", "build-registry.R")
generated    <- file.path(repo_root, "src", "tests", ".generated")
run_dir      <- file.path(build_home, "headless", "run")

# --- what to build -----------------------------------------------------------
setup_path <- opt("setup")
if (is.null(setup_path) || !nzchar(setup_path)) {
  stop("build-linelist.R: --setup=<path> is required (the filled setup workbook to generate from).")
}

designer_path <- opt("designer", file.path(repo_root, ".mock", "designer_mock.xlsb"))
forms_folder  <- opt("forms",    file.path(build_home, "forms", "merged"))
out_folder    <- opt("out",      file.path(build_home, "build"))
out_name      <- opt("name",     "linelist")

# Every option key is passed on every run, empty when unset. An empty value is
# meaningful to the build and documented as such -- empty temppath is the
# buttons build, empty setuplang is "read it off the setup" -- so there is no
# case where omitting a key says something a blank one does not.
build_options <- paste(
  paste0("temppath=",   opt("temppath",   "")),
  paste0("geopath=",    opt("geopath",    "")),
  paste0("setuplang=",  opt("setuplang",  "")),
  paste0("lllang=",     opt("lllang",     "")),
  paste0("llpassword=", opt("llpassword", "")),
  sep = "|"
)

if (on_windows) {
  message("build-linelist.R: NOTE - the Windows trigger has never been run ",
          "against a real Windows Excel. Read what it reports with that in mind.")
}
for (needed in c(workbook_src, trigger, designer_path, setup_path, registry_r)) {
  if (!file.exists(needed)) stop("build-linelist.R: not found: ", needed)
}

message("build-linelist.R: designer  -> ", designer_path)
message("build-linelist.R: setup     -> ", setup_path)
message("build-linelist.R: output    -> ", file.path(out_folder, paste0(out_name, ".xlsb")))
message("build-linelist.R: ",
        if (nzchar(opt("temppath", ""))) {
          paste0("ribbon build, template ", opt("temppath", ""))
        } else {
          "buttons build (no ribbon template given): action buttons on the sheets, an Admin sheet"
        })

# --- 1) the merged forms -----------------------------------------------------
# The forms in .mock carry whatever code was in the workbook the day they were
# exported; the code that belongs in them lives in src/modules/linelistform. A
# build over unmerged forms delivers every control wired to stale handlers, and
# nothing about the delivered file says so.
if (do_merge) {
  message("build-linelist.R: merging the current form code ...")
  rc <- system2("Rscript", c(shQuote(merger), "--out", shQuote(forms_folder)))
  if (rc != 0L) stop("build-linelist.R: merge-form-code.R failed (exit ", rc, ").")
} else {
  message("build-linelist.R: --no-merge, using the forms already in ", forms_folder)
}
if (!dir.exists(forms_folder)) {
  stop("build-linelist.R: no merged forms at ", forms_folder,
       " (drop --no-merge, or point --forms at a folder that has them).")
}

# --- 2) the source list, filtered down to the build closure ------------------
message("build-linelist.R: rebuilding registry intermediates ...")
rc <- system2("Rscript", shQuote(registry_r))
if (rc != 0L) stop("build-linelist.R: build-registry.R failed (exit ", rc, ").")

tbl <- read.delim(file.path(generated, "code-tables.tsv"), stringsAsFactors = FALSE)
build_tags <- c("general classes", "general modules")
tbl <- tbl[tbl$tag %in% build_tags, , drop = FALSE]
if (!nrow(tbl)) {
  stop("build-linelist.R: the registry described no classes or modules to import.")
}
if (!"headless" %in% tbl$folder) {
  stop("build-linelist.R: src/modules/headless is not registered, so the driver ",
       "would have no HeadlessBuild to call. Add it under `modules:` in ",
       "src/tests/test-registry.yml.")
}

# --- 3) assemble the run dir -------------------------------------------------
# Contents cleared, the folder and the three files Excel touches kept in place.
work_copy   <- file.path(run_dir, "linelist_build.xlsb")
report_path <- file.path(run_dir, "build-report.txt")   # written by the trigger
log_path    <- file.path(run_dir, "obt-import.log")     # written by OBTImport

keep <- c(basename(work_copy), basename(report_path), basename(log_path))

if (dir.exists(run_dir)) {
  stale <- list.files(run_dir, all.files = TRUE, full.names = TRUE, no.. = TRUE)
  unlink(stale[!(basename(stale) %in% keep)], recursive = TRUE, force = TRUE)
} else {
  dir.create(run_dir, recursive = TRUE, showWarnings = FALSE)
}

if (!file.copy(workbook_src, work_copy, overwrite = TRUE)) {
  stop("build-linelist.R: failed to copy the driver workbook into ", run_dir)
}
for (kept_file in c(report_path, log_path)) cat("", file = kept_file, append = FALSE)

tag_bases <- c("general classes" = file.path(repo_root, "src", "classes"),
               "general modules" = file.path(repo_root, "src", "modules"))
tag_runs  <- c("general classes" = file.path(run_dir, "classes"),
               "general modules" = file.path(run_dir, "modules"))

pairs <- unique(tbl[, c("folder", "tag")])
for (i in seq_len(nrow(pairs))) {
  found <- file.path(tag_bases[[pairs$tag[i]]], pairs$folder[i])
  if (!dir.exists(found)) {
    stop("build-linelist.R: the registry names ", pairs$tag[i], " folder '",
         pairs$folder[i], "' and it is not in src/.")
  }
  dir.create(tag_runs[[pairs$tag[i]]], recursive = TRUE, showWarnings = FALSE)
  file.copy(found, tag_runs[[pairs$tag[i]]], recursive = TRUE)
}

# OBTImport points a folder range at run/tests whatever the tables hold, so the
# folder exists and is empty. Empty is the whole point: no test module reaches
# the driver, and a test module that will not compile cannot stop a build.
dir.create(file.path(run_dir, "tests"), showWarnings = FALSE)

dir.create(file.path(run_dir, ".generated"), showWarnings = FALSE)
write.table(tbl, file.path(run_dir, ".generated", "code-tables.tsv"),
            sep = "\t", quote = FALSE, row.names = FALSE)
cat("", file = file.path(run_dir, ".generated", "modules-for-testing.txt"))

dir.create(file.path(run_dir, "bootstrap"), showWarnings = FALSE)
for (f in c("OBTImport.bas", "OBTHeadless.bas")) {
  src_f <- file.path(repo_root, "src", "tests", "rubberduck", f)
  if (file.exists(src_f)) {
    file.copy(src_f, file.path(run_dir, "bootstrap", f), overwrite = TRUE)
  }
}

message("build-linelist.R: staged ", nrow(tbl), " component(s) from ",
        nrow(pairs), " folder(s), no test module among them")

# --- 4) clear the way --------------------------------------------------------
dir.create(out_folder, recursive = TRUE, showWarnings = FALSE)

# The three files a build writes. A stale linelist left in place would let a
# failed run report a file on disk and read as a success.
for (leaf in c(".xlsb", "-generation.txt", "-designer.xlsb")) {
  unlink(file.path(out_folder, paste0(out_name, leaf)), force = TRUE)
}

# Only macOS needs this. There, `run VB macro` returns a Parameter error against
# an Excel that is already open interactively, so the instance has to go first.
# On Windows CreateObject("Excel.Application") makes its own instance and an
# Excel the operator has open is none of its business.
excel_running <- function() {
  if (on_windows) return(FALSE)
  identical(system2("pgrep", c("-x", "Microsoft Excel"),
                    stdout = FALSE, stderr = FALSE), 0L)
}
if (excel_running()) {
  message("build-linelist.R: quitting the running Excel instance ...")
  system2("osascript",
          c("-e", shQuote('tell application "Microsoft Excel" to quit saving no')),
          stdout = FALSE, stderr = FALSE)
  for (i in seq_len(40)) {
    if (!excel_running()) break
    Sys.sleep(0.5)
  }
  if (excel_running()) {
    stop("build-linelist.R: Excel would not quit; close it by hand before re-running.")
  }
}

# --- 5) build ----------------------------------------------------------------
trigger_args <- shQuote(c(
  trigger, work_copy, designer_path, setup_path, repo_root,
  forms_folder, out_folder, out_name, build_options, report_path
))

message("build-linelist.R: launching Excel via ",
        if (on_windows) "cscript" else "osascript", " ...")
trigger_rc <- if (on_windows) {
  system2("cscript", c("//nologo", trigger_args))
} else {
  system2("osascript", trigger_args)
}

# --- 6) read back ------------------------------------------------------------
show_import_log <- function() {
  if (file.exists(log_path) && file.size(log_path) > 0L) {
    message("\n       --- obt-import.log ---")
    for (ln in readLines(log_path, warn = FALSE)) message("       ", ln)
    message("       --- end log ---")
  }
}

fail <- function(msg) {
  message("\n[FAIL] ", msg)
  show_import_log()
  message("\n       Run dir kept for inspection:")
  message("       ", run_dir)
  quit(status = 1L, save = "no")
}

if (!file.exists(report_path) || file.size(report_path) == 0L) {
  fail(sprintf("the trigger wrote no report (osascript returned %d; Excel may have wedged on a dialog).",
               trigger_rc))
}

report_lines <- readLines(report_path, warn = FALSE)
marker <- match("--report--", report_lines)
fields <- if (is.na(marker)) report_lines else report_lines[seq_len(marker - 1L)]
narrative <- if (is.na(marker)) character(0) else report_lines[-seq_len(marker)]

field <- function(key) {
  hit <- grep(paste0("^", key, "="), fields, value = TRUE)
  if (!length(hit)) "" else sub(paste0("^", key, "="), "", hit[1])
}

outcome  <- field("outcome")
built    <- field("linelist")

if (length(narrative)) {
  message("\nWhat the build did:")
  for (ln in narrative) if (nzchar(trimws(ln))) message("  ", ln)
}

message(sprintf("\nbuild-linelist.R: %s sheet(s), %s variable(s), %s component(s) re-imported.",
                field("sheets"), field("variables"), field("components")))

if (!identical(outcome, "OK")) {
  fail(paste0("the build answered: ", outcome))
}
if (!nzchar(built) || !file.exists(built)) {
  fail(paste0("the build answered OK and there is no file at: ", built))
}

message("\nbuild-linelist.R: linelist -> ", built,
        " (", format(file.size(built), big.mark = ","), " bytes)")
if (nzchar(field("log"))) message("build-linelist.R: log      -> ", field("log"))
quit(status = 0L, save = "no")
