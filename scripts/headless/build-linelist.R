# build-linelist.R
# =============================================================================
# Generate one linelist, headless. One command, and it runs no test:
#
#   Rscript scripts/headless/build-linelist.R --setup=<filled-setup.xlsb> \
#           [--out=<dir>] [--name=<stem>] [--designer=<path>] \
#           [--forms=<dir>] [--home=<dir>] [--obt-home=<dir>] [--no-merge] \
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
# ONE FOLDER, ONE GRANT
# -----------------------------------------------------------------------------
# Excel for Mac is sandboxed, and a VBA file read on a path it holds no grant
# for pops a dialog that a headless run has nobody to answer. A probe run with
# the operator watching the screen settled how the grant behaves (written up in
# .obt/gotchas/macos-sandbox-grant.md): one FOLDER grant persists across Excel
# being quit and relaunched, and it cascades to files created inside that tree
# afterwards.
#
# So everything Excel touches is staged under a single root, OBT_HOME, and the
# trigger grants that one folder before its first read:
#
#   <OBT_HOME>/run/      the driver copy, bootstrap/, the filtered sources,
#                        .generated/, build-report.txt, obt-import.log
#   <OBT_HOME>/src/      classes/** + modules/**, the tree the linelist is
#                        built from (this is what the build calls sourceRoot)
#   <OBT_HOME>/forms/    the merged .frm files
#   <OBT_HOME>/in/       designer, setup, ribbon template, geobase
#   <OBT_HOME>/out/      the linelist, its log, the designer used, OBTApp_/
#
# This script is NOT sandboxed, so it copies the operator's setup (and designer,
# template, geobase) IN beforehand and copies the finished linelist back OUT to
# --out afterwards. Excel never sees a path outside OBT_HOME, which is what
# makes a packaged install click-free: one grant panel on a machine that has
# never run it, and none after that.
#
# OBT_HOME is --obt-home, else the OBT_HOME environment variable (an .Renviron
# entry is the deployed way to set it), else <repo>/.headless-runner. Its own
# folder rather than a corner of the test harness's working area: the two runs
# share a registry and nothing else, and keeping the headless tree separate is
# what lets it be swept, moved or granted without touching the test loop.
#
# WHAT THE GRANT ACTUALLY BUYS, which is less than this once assumed. Measured
# on 2026-08-14 (.obt/gotchas/macos-sandbox-grant.md): a grant is PER ITEM and
# lasts for the SESSION. It does not cover subfolders and it does not survive
# Excel being quit. So a run still meets panels; consolidating under one root
# cuts how many and makes them nameable in one place. Excel's own container is
# the only spot that needs no grant at all, and the launcher cannot write there
# (TCC), which is the open problem.
#
# Within it the run dir is cleared and reused rather than deleted and remade,
# and the workbook copy is overwritten in place. That predates the grant work
# and is kept: it costs nothing and a file kept in place keeps its identity.
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

# THE granted root. Everything Excel reads or writes goes under it, and the
# trigger asks the sandbox about this one path and nothing else.
#
# Sys.getenv reads .Renviron, which is how a deployed install pins it without
# anyone having to remember a flag: one OBT_HOME=... line in the operator's
# .Renviron and every run lands in the folder they granted.
obt_home <- opt("obt-home", default = {
  if (nzchar(Sys.getenv("OBT_HOME"))) Sys.getenv("OBT_HOME")
  else file.path(repo_root, ".headless-runner")
})
if (!grepl("^(/|~|[A-Za-z]:)", obt_home)) obt_home <- file.path(repo_root, obt_home)
obt_home <- path.expand(obt_home)
dir.create(obt_home, recursive = TRUE, showWarnings = FALSE)
obt_home <- normalizePath(obt_home, mustWork = TRUE)

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

# The five places under the granted root. Nothing Excel touches lives outside
# these, which is the whole of the sandbox fix.
run_dir      <- file.path(obt_home, "run")
staged_forms <- file.path(obt_home, "forms")
staged_in    <- file.path(obt_home, "in")
build_out    <- file.path(obt_home, "out")

# sourceRoot is the root itself, because HeadlessBuild joins "src/classes" and
# "src/modules" onto whatever it is given. So the tree lands at
# <OBT_HOME>/src/classes and the build finds it where it expects to.
source_root  <- obt_home

# --- what to build -----------------------------------------------------------
# Every path handed on to Excel is made absolute first. R resolves a relative
# path against the shell's working directory; Excel resolves it against its
# own, which is somewhere else entirely. A relative --setup therefore passes
# the file.exists check here and comes back from the build as
# "no setup workbook at <path>", naming a file that is plainly sitting there.
abs_path <- function(p) {
  if (is.null(p) || !nzchar(p)) return(p)
  normalizePath(p, mustWork = FALSE)
}

setup_path <- opt("setup")
if (is.null(setup_path) || !nzchar(setup_path)) {
  stop("build-linelist.R: --setup=<path> is required (the filled setup workbook to generate from).")
}
setup_path <- abs_path(setup_path)

designer_path <- abs_path(opt("designer", file.path(repo_root, ".mock", "designer_mock.xlsb")))
temp_path     <- abs_path(opt("temppath", ""))
geo_path      <- abs_path(opt("geopath",  ""))
out_name      <- opt("name", "linelist")

# Where the OPERATOR wants the three files. Excel never writes here: it builds
# into <OBT_HOME>/out and this script copies the result across at the end.
dest_folder <- abs_path(opt("out", file.path(build_home, "build")))

# --forms, when given, is a folder of already-merged .frm files somewhere else;
# it gets copied into the root like every other input. Left unset, the merge
# below writes straight into the root and there is nothing to copy.
forms_source <- abs_path(opt("forms", ""))

if (on_windows) {
  message("build-linelist.R: NOTE - the Windows trigger has never been run ",
          "against a real Windows Excel. Read what it reports with that in mind.")
}
for (needed in c(workbook_src, trigger, designer_path, setup_path, registry_r)) {
  if (!file.exists(needed)) stop("build-linelist.R: not found: ", needed)
}
for (optional in c(temp_path, geo_path)) {
  if (nzchar(optional) && !file.exists(optional)) {
    stop("build-linelist.R: not found: ", optional)
  }
}

message("build-linelist.R: OBT_HOME  -> ", obt_home, " (the one granted folder)")
message("build-linelist.R: designer  -> ", designer_path)
message("build-linelist.R: setup     -> ", setup_path)
message("build-linelist.R: output    -> ", file.path(dest_folder, paste0(out_name, ".xlsb")))
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
#
# They land in <OBT_HOME>/forms whichever way they arrive, because that is where
# Excel is allowed to read them from.
dir.create(staged_forms, recursive = TRUE, showWarnings = FALSE)

if (do_merge) {
  message("build-linelist.R: merging the current form code ...")
  rc <- system2("Rscript", c(shQuote(merger), "--out", shQuote(staged_forms)))
  if (rc != 0L) stop("build-linelist.R: merge-form-code.R failed (exit ", rc, ").")
} else if (nzchar(forms_source)) {
  message("build-linelist.R: --no-merge, copying the forms from ", forms_source)
  if (!dir.exists(forms_source)) {
    stop("build-linelist.R: no merged forms at ", forms_source)
  }
  unlink(list.files(staged_forms, full.names = TRUE), recursive = TRUE, force = TRUE)
  copied <- file.copy(list.files(forms_source, full.names = TRUE), staged_forms,
                      recursive = TRUE, overwrite = TRUE)
  if (!all(copied)) stop("build-linelist.R: failed to copy the forms into ", staged_forms)
} else {
  message("build-linelist.R: --no-merge, using the forms already in ", staged_forms)
}

if (!length(list.files(staged_forms, pattern = "\\.frm$"))) {
  stop("build-linelist.R: no merged forms at ", staged_forms,
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

# --- 3b) stage the source tree and the inputs into the granted root ----------
# Two different things are staged from src/, and confusing them is easy:
#
#   run/classes, run/modules  the filtered build closure, imported into the
#                             DRIVER so it has a HeadlessBuild to call
#   src/classes, src/modules  the whole tree, re-imported into the LINELIST by
#                             RefreshSourceCode (it reads <sourceRoot>/src/...)
#
# The whole tree rather than the registered folders, because HeadlessBuild picks
# what it wants out of it through TRANSFER_CLASS_FOLDERS. Copying only what that
# constant names today would put a second copy of the list here, and a second
# copy goes stale silently.
sync_tree <- function(from, to) {
  if (!dir.exists(from)) stop("build-linelist.R: nothing to stage from ", from)
  dir.create(dirname(to), recursive = TRUE, showWarnings = FALSE)
  unlink(to, recursive = TRUE, force = TRUE)
  dir.create(to, recursive = TRUE, showWarnings = FALSE)
  copied <- file.copy(list.files(from, full.names = TRUE), to, recursive = TRUE)
  if (!all(copied)) {
    stop("build-linelist.R: failed to stage ", sum(!copied), " item(s) from ",
         from, " into ", to)
  }
  invisible(TRUE)
}

sync_tree(file.path(repo_root, "src", "classes"), file.path(source_root, "src", "classes"))
sync_tree(file.path(repo_root, "src", "modules"), file.path(source_root, "src", "modules"))

# The operator's own files, copied in under fixed names. Excel is handed these
# copies and never the originals, which is what keeps every path it sees inside
# the granted root no matter where the operator keeps their setup.
dir.create(staged_in, recursive = TRUE, showWarnings = FALSE)

stage_in <- function(from, leaf) {
  if (!nzchar(from)) return("")
  to <- file.path(staged_in, leaf)
  if (!file.copy(from, to, overwrite = TRUE)) {
    stop("build-linelist.R: failed to stage ", from, " into ", staged_in)
  }
  to
}

staged_designer <- stage_in(designer_path, "designer.xlsb")
staged_setup    <- stage_in(setup_path,    "setup.xlsb")
staged_template <- stage_in(temp_path,     "template.xlsb")
staged_geo      <- stage_in(geo_path,      "geo.xlsb")

# Built from the STAGED paths, so a ribbon template or geobase the operator
# keeps on their Desktop is read from inside the root like everything else.
#
# Every option key is passed on every run, empty when unset. An empty value is
# meaningful to the build and documented as such -- empty temppath is the
# buttons build, empty setuplang is "read it off the setup" -- so there is no
# case where omitting a key says something a blank one does not.
build_options <- paste(
  paste0("temppath=",   staged_template),
  paste0("geopath=",    staged_geo),
  paste0("setuplang=",  opt("setuplang",  "")),
  paste0("lllang=",     opt("lllang",     "")),
  paste0("llpassword=", opt("llpassword", "")),
  sep = "|"
)

# --- 4) clear the way --------------------------------------------------------
# Excel builds into the root; the operator's folder is filled from there at the
# end of the run.
dir.create(build_out, recursive = TRUE, showWarnings = FALSE)
dir.create(dest_folder, recursive = TRUE, showWarnings = FALSE)

# The three files a build writes, cleared in BOTH places. A stale linelist left
# in either would let a failed run report a file on disk and read as a success.
build_leaves <- c(".xlsb", "-generation.txt", "-designer.xlsb")
for (leaf in build_leaves) {
  unlink(file.path(build_out,   paste0(out_name, leaf)), force = TRUE)
  unlink(file.path(dest_folder, paste0(out_name, leaf)), force = TRUE)
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
  # `quit`, never pkill. The sandbox grant is a bookmark Excel holds and writes
  # out on normal termination; a killed Excel can lose it, and losing it costs
  # the operator the grant panel again on the next run.
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
# Every path here is inside obt_home, which is the last argument and the only
# one the sandbox is asked about. sourceRoot is the staged tree rather than the
# repo, and outFolder is the root's own out/ rather than the operator's folder.
trigger_args <- shQuote(c(
  trigger, work_copy, staged_designer, staged_setup, source_root,
  staged_forms, build_out, out_name, build_options, report_path, obt_home
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
grant    <- field("grant")

if (length(narrative)) {
  message("\nWhat the build did:")
  for (ln in narrative) if (nzchar(trimws(ln))) message("  ", ln)
}

# Printed on every run, because a grant that stopped taking is invisible from
# here otherwise: nothing in an exit code or an elapsed time can see a dialog,
# and the operator is the only one who can.
if (nzchar(grant)) message("\nbuild-linelist.R: sandbox   -> ", grant)

message(sprintf("\nbuild-linelist.R: %s sheet(s), %s variable(s), %s component(s) re-imported.",
                field("sheets"), field("variables"), field("components")))

if (!identical(outcome, "OK")) {
  fail(paste0("the build answered: ", outcome))
}
if (!nzchar(built) || !file.exists(built)) {
  fail(paste0("the build answered OK and there is no file at: ", built))
}

message("\nbuild-linelist.R: built    -> ", built,
        " (", format(file.size(built), big.mark = ","), " bytes)")

# --- 7) hand the files over --------------------------------------------------
# The build wrote into the granted root because that is the only place Excel is
# allowed to write without asking. This script is under no such restriction, so
# it does the last hop to wherever the operator asked for the files.
#
# A copy that fails is a failure of the run: the operator asked for a linelist
# in their folder and there is none there, whatever is sitting in the root.
delivered <- character(0)
if (!identical(normalizePath(build_out), normalizePath(dest_folder))) {
  for (leaf in build_leaves) {
    from <- file.path(build_out, paste0(out_name, leaf))
    if (!file.exists(from)) next
    to <- file.path(dest_folder, paste0(out_name, leaf))
    if (!file.copy(from, to, overwrite = TRUE)) {
      fail(paste0("the build succeeded and the file could not be copied out to ", to))
    }
    delivered <- c(delivered, to)
  }
} else {
  delivered <- file.path(dest_folder, paste0(out_name, build_leaves))
  delivered <- delivered[file.exists(delivered)]
}

message("build-linelist.R: linelist -> ", delivered[1])
for (extra in delivered[-1]) message("build-linelist.R:          -> ", extra)
quit(status = 0L, save = "no")
