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
#             OBT_TEST_HOME env var, else -- on macOS -- Excel's own container at
#             ~/Library/Containers/com.microsoft.Excel/Data/Documents/OBTTestHome,
#             else <repo>/.test-runner. Gitignored, so its location is a local
#             setting, not fixed in the repo.
#
# FILE ACCESS (macOS): THE ANSWER IS TO STAGE INSIDE EXCEL'S CONTAINER.
# Excel for Mac is sandboxed, and a VBA file read on an ungranted path raises a
# modal grant dialog -- one per file. A headless run has nobody to click it, so
# it stops there with Excel at 0% CPU and a zero-byte results CSV. That reads
# exactly like the flaky import wedge; `sample <excel pid>` separates them, and
# `runModalForWindow` over `MbuObtainSecurityBookmarksForFileURLs` in the stack
# is the grant dialog.
#
# Inside Excel's own container nothing is granted or picked, which is why the
# default moved there on 2026-08-31. See the note at excel_container_test_home.
# Application.GrantAccessToMultipleFiles does NOT help: it raises 438 on this
# host and has never worked. Outside the container the one thing that does work
# is a folder pick made inside Excel -- OBTGrantAccess, once, from the VBE --
# and that is the fallback path only.
#
# The harness + probes now live in src/ (promoted). This script assembles a
# per-run dir under <home>/tests/staging/run/ from src/ (the workbook is the only
# thing kept in the working area, as an untracked binary).
#
# What it does (all portable logic lives here; the OS-specific trigger is thin):
#   1. copy <home>/unit_tests_dev.xlsb -> the stable run dir (NEVER touch original)
#      and assemble the registered probe sources + manifest + harness from src/.
#   2. (optional, --build) regenerate the registry intermediates (src/tests/.generated)
#      and merge the FormLogic code into the exported forms (<home>/forms/merged).
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
# trigger differs, and the host picks it — macos/run-tests.applescript through
# osascript, windows/run-tests.vbs through cscript. Both drive the same OBT*
# entry points in the same order, so a difference in results is a difference in
# the VBA rather than in the harness. The Windows side runs against a real
# Windows Excel, hidden for the length of the run so it never takes the screen
# from the operator. A hidden Excel refuses Window.FreezePanes and SheetView
# writes with error 1004; production logs those refusals and carries on since
# commit 0b4c7bb9, and no test asserts a frozen pane, so a hidden run stays
# green with its panes left unfrozen.
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
# so where it lives is a local choice, not a repo fact: --home=<dir> wins, else
# the OBT_TEST_HOME env var, else -- on macOS -- Excel's own container, else
# <repo>/.test-runner. A relative home resolves against the repo root, so the
# default and a relative --home behave the same from any cwd.
# Needed this early: the container default below is macOS-only.
on_windows <- identical(.Platform$OS.type, "windows")

#
# ON macOS THE DEFAULT IS INSIDE EXCEL'S OWN CONTAINER, and that is the whole
# answer to the grant problem. Excel for Mac is sandboxed: a VBA file read
# (Open, Dir, VBComponents.Import) outside the container needs a security-scoped
# grant, and only a folder PICK in a dialog creates one. A headless run has
# nobody to click it, so the run stops dead on a modal with Excel at 0% CPU and
# a zero-byte results CSV -- which looks exactly like the flaky import wedge and
# is not one. Inside the container Excel needs no permission for anything and
# never puts a panel on the screen.
#
# build-linelist.R has staged inside the container since 2026-08-15 (see its
# note at excel_container_home). This loop staged in the repo until 2026-08-31
# and paid the grant for it: a session that only EDITED files never noticed,
# because every path already had a bookmark, while a session that REGISTERED a
# new class, module, test module or form put files there that had none and died
# on the dialog.
#
# The catch is on THIS side of the fence: macOS keeps Excel's container away
# from other programs, so R can only reach it once the operator has given their
# terminal app Full Disk Access (System Settings -> Privacy & Security). Where
# that is not done the probe below fails and the run falls back to
# <repo>/.test-runner, the old path with the old rules -- somebody runs
# OBTGrantAccess once from the VBE and picks the repo root.
excel_container_test_home <- function() {
  if (on_windows) return(NA_character_)
  root <- path.expand("~/Library/Containers/com.microsoft.Excel/Data/Documents")
  if (!dir.exists(root)) return(NA_character_)
  # Written rather than merely tested: dir.exists() answers TRUE for a folder
  # macOS will refuse to open, so only a real write settles it.
  home  <- file.path(root, "OBTTestHome")
  probe <- file.path(home, ".reachable")
  ok <- tryCatch({
    dir.create(home, recursive = TRUE, showWarnings = FALSE)
    writeLines("ok", probe)
    identical(readLines(probe, warn = FALSE), "ok")
  }, error = function(e) FALSE, warning = function(w) FALSE)
  unlink(probe, force = TRUE)
  if (isTRUE(ok)) home else NA_character_
}

chose_container <- FALSE
test_home <- if (length(home_opt) && nzchar(home_opt[1])) {
  home_opt[1]
} else if (nzchar(Sys.getenv("OBT_TEST_HOME"))) {
  Sys.getenv("OBT_TEST_HOME")
} else {
  container <- excel_container_test_home()
  if (!is.na(container)) {
    chose_container <- TRUE
    container
  } else {
    file.path(repo_root, ".test-runner")
  }
}
if (!grepl("^(/|~|[A-Za-z]:)", test_home)) test_home <- file.path(repo_root, test_home)
test_home <- path.expand(test_home)

workbook_src <- file.path(test_home, "unit_tests_dev.xlsb")
# Where merge-form-code.R writes the importable forms. The merger defaults to
# this same path, so both scripts name one tree.
merged_forms <- file.path(test_home, "forms", "merged")
scripts_dir  <- file.path(repo_root, "scripts", "tests")
# The trigger is the only OS-specific piece; both drive the same OBT* entry
# points in the same order, so a difference in results is a difference in the
# VBA rather than in the harness. on_windows itself is set further up, because
# the container default needs it before this point.
trigger      <- if (on_windows) {
  file.path(scripts_dir, "windows", "run-tests.vbs")
} else {
  file.path(scripts_dir, "macos", "run-tests.applescript")
}
generated    <- file.path(repo_root, "src", "tests", ".generated")     # build-registry.R output
staging      <- file.path(test_home, "tests", "staging")                # untracked staging tree
# STABLE run dir (fixed path, cleared+reused each run). macOS grants Excel file
# access per PATH, so a fixed path is prompted for at most once, then never again
# -- a fresh run-<stamp> dir each time is what caused the repeated grant prompts.
run_dir      <- file.path(staging, "run")

# A home this script chose for itself is a home it has to furnish. The operator
# picked nothing, so there is nobody to tell that <home>/tests/staging is
# missing; and the driver workbook is an untracked binary that exists only in
# whatever working area the machine used before, so it is copied across rather
# than asked for. A home given by --home or OBT_TEST_HOME is left exactly as it
# is: an explicit choice is never second-guessed, and the checks below report it.
if (chose_container) {
  dir.create(staging, recursive = TRUE, showWarnings = FALSE)
  if (!file.exists(workbook_src)) {
    legacy_workbook <- file.path(repo_root, ".test-runner", "unit_tests_dev.xlsb")
    if (file.exists(legacy_workbook)) {
      if (!file.copy(legacy_workbook, workbook_src, overwrite = FALSE)) {
        stop("run-tests.R: could not seed the driver workbook into ", test_home)
      }
      message("run-tests.R: seeded the driver workbook into Excel's container ",
              "from ", legacy_workbook)
    }
  }
}

if (!dir.exists(staging)) {
  stop("run-tests.R: staging tree not found: ", staging,
       " (set --home / OBT_TEST_HOME, or create <home>/tests/staging).")
}

if (on_windows) {
  message("run-tests.R: NOTE - Excel runs hidden for the length of the run. ",
          "It is the trigger's own instance, so anything already open is ",
          "left alone.")
}
if (!file.exists(workbook_src)) {
  stop("run-tests.R: workbook not found: ", workbook_src,
       "\n  The driver workbook is an untracked binary. Copy it there, or point",
       "\n  --home / OBT_TEST_HOME at the working area that already holds it.")
}

message("run-tests.R: test home -> ", test_home,
        if (chose_container) "  (inside Excel's container: no file-access grant is needed)" else "")
if (!on_windows && !chose_container) {
  message("run-tests.R: NOTE - this home is OUTSIDE Excel's container, so every ",
          "file Excel\n  imports needs a sandbox grant. A newly registered ",
          "component with no grant stops\n  the run on a dialog nobody can ",
          "click. Run OBTGrantAccess once from the VBE and\n  pick the repo ",
          "root, or let the default choose the container.")
}
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

# --- optional: merge the form code -------------------------------------------
# merge-form-code.R rewrites the code part of every exported .frm with the
# current src/modules/linelistform/FormLogic*.bas body and copies the .frx
# beside it. The pair is what the VBE imports, so a form staged from a merged
# tree runs the code the linelist ships. It reads .mock/forms/designer, a
# tracked tree, so this adds no untracked dependency. --build is the flag that
# already means "regenerate the intermediates", and the merged tree is one.
if (do_build) {
  message("run-tests.R: merging FormLogic code into the exported forms ...")
  merger <- file.path(repo_root, "scripts", "headless", "merge-form-code.R")
  rc <- system2("Rscript", c(shQuote(merger), "--out", shQuote(merged_forms)))
  if (rc != 0L) stop("run-tests.R: merge-form-code.R failed (exit ", rc, ").")

  # The setup's own form joins the merged tree by a straight copy.
  #
  # merge-form-code.R reads .mock/forms/designer only, and it merges because a
  # designer form's code lives apart from it, in src/modules/linelistform. The
  # setup's [Imports] form is not built that way: its .frm carries no code at
  # all -- sixteen lines of header, and its logic sits in
  # src/modules/setup/ImportForm.bas -- so there is nothing to merge into it and
  # the pair is copied as it is.
  #
  # It is here rather than in the merger because the merger also runs as a step
  # of the headless build, which builds a linelist and has no use for a setup
  # form. Only the test workbook needs this one, and it needs it because
  # SetupHelpers declares `Dim formRef As Imports` and reads eleven of its
  # controls: without the form in the project, SetupHelpers does not compile and
  # the whole setup folder goes red with blank counts.
  setup_forms <- file.path(repo_root, ".mock", "forms", "setup")
  if (dir.exists(setup_forms)) {
    dir.create(merged_forms, recursive = TRUE, showWarnings = FALSE)
    staged <- list.files(setup_forms, pattern = "\\.(frm|frx)$", full.names = TRUE)
    if (length(staged)) {
      ok <- file.copy(staged, merged_forms, overwrite = TRUE)
      if (!all(ok)) {
        stop("run-tests.R: could not stage the setup forms from ", setup_forms)
      }
      message(
        "run-tests.R: staged ", length(staged),
        " setup form file(s) into the merged tree"
      )
    }
  } else {
    message("run-tests.R: WARNING - no setup forms tree at ", setup_forms)
  }
}

# --- 1) per-run working copy -------------------------------------------------
# Keep the run dir at a STABLE identity (same folder, not just same path) every
# run. A macOS file-access grant is tied to the identity (inode), not the path,
# so DELETING and recreating a folder -- even at the same path -- makes a new
# folder the grant no longer covers, and Excel re-prompts. So clear the CONTENTS
# and keep the folder itself.
#
# The SAME rule applies to the workbook copy, and it is the file Excel actually
# opens. `unit_tests_dev.xlsb` is only ever read by this script; a grant on it
# can never help Excel. So the copy is kept too: it is left in place and
# file.copy(overwrite = TRUE) truncates and rewrites it, which preserves the
# inode (verified on this machine). Sweeping it away first, as this used to do,
# handed Excel a brand new file every run and threw away any grant on it.
# The same rule reaches the two files EXCEL ITSELF creates -- test-results.csv
# and obt-import.log. Sweeping them handed Excel a brand new file to create on
# every run, and a file Excel creates outside its container needs a grant, so
# the operator was asked for access to obt-import.log every single time. They
# are kept too, and EMPTIED IN PLACE instead: cat(append = FALSE) opens with
# "w", which truncates and preserves the inode (verified on this machine).
# Emptying loses nothing that sweeping bought -- a 0-byte file carries no stale
# result -- so the guards below test for CONTENT rather than for existence.
# Everything else in the run dir IS swept.
work_copy <- file.path(run_dir, "unit_tests_run.xlsb")
csv_path  <- file.path(run_dir, "test-results.csv")   # written by OBTRunAllTests
log_path  <- file.path(run_dir, "obt-import.log")     # written by OBTImport

keep <- c(basename(work_copy), basename(csv_path), basename(log_path))

# The SOURCE TREE is kept for the identity reason above, and it is the reason
# the operator was asked for access over and over. These folders were swept and
# re-copied every run, so every .cls and .bas inside them was a brand new file
# with a brand new inode, and the grants Excel had been given the run before
# covered none of them. The grant store held 144 per-file entries and still
# prompted, because the files they named no longer existed.
#
# They are refreshed in place instead, by copy_tree_in_place below. A file that
# left src/ is left behind here, which costs nothing: Development.ImportAll
# imports what the code-tables manifest names, never what the folder happens to
# hold, so a stale file is never imported and never runs.
keep_dirs <- c("classes", "modules", "tests", "forms", "bootstrap", ".generated")

if (dir.exists(run_dir)) {
  stale <- list.files(run_dir, all.files = TRUE, full.names = TRUE, no.. = TRUE)
  stale <- stale[!(basename(stale) %in% c(keep, keep_dirs))]
  unlink(stale, recursive = TRUE, force = TRUE)
} else {
  dir.create(run_dir, recursive = TRUE, showWarnings = FALSE)
}

# Copy every file of a tree over the one already there, keeping the destination
# files themselves. file.copy(overwrite = TRUE) truncates and rewrites, which
# preserves the inode, so a grant on the file survives the refresh. Copying the
# FOLDER instead -- file.copy(from_dir, to_dir, recursive = TRUE) -- is what
# used to replace the whole subtree.
copy_tree_in_place <- function(from_dir, to_dir) {
  rels <- list.files(from_dir, recursive = TRUE, all.files = TRUE, no.. = TRUE)
  for (rel in rels) {
    src <- file.path(from_dir, rel)
    if (dir.exists(src)) next
    dst <- file.path(to_dir, rel)
    dir.create(dirname(dst), recursive = TRUE, showWarnings = FALSE)
    file.copy(src, dst, overwrite = TRUE)
  }
}

if (!file.copy(workbook_src, work_copy, overwrite = TRUE)) {
  stop("run-tests.R: failed to copy workbook into ", run_dir)
}

# Create them if they are missing, empty them if they are not. Creating them
# here means the FIRST run after this change is the last one that can prompt.
for (kept_file in c(csv_path, log_path)) {
  cat("", file = kept_file, append = FALSE)
}

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
                        file.path(staging, "tests")),
  # The forms come from the merged tree the step above writes, and that tree is
  # FLAT: F_Geo.frm and F_Geo.frx sit in it side by side. The two loops below
  # both walk <base>/<folder>, so the forms are staged by their own loop.
  "general forms"   = c(merged_forms)
)
tag_runs <- list(
  "general classes" = file.path(run_dir, "classes"),
  "general modules" = file.path(run_dir, "modules"),
  "tests modules"   = file.path(run_dir, "tests"),
  "tests classes"   = file.path(run_dir, "tests"),
  "general forms"   = file.path(run_dir, "forms")
)
tbl <- read.delim(file.path(generated, "code-tables.tsv"), stringsAsFactors = FALSE)
pairs <- unique(tbl[, c("folder", "tag")])
for (i in seq_len(nrow(pairs))) {
  if (identical(pairs$tag[i], "general forms")) next   # flat tree; staged below
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
  copy_tree_in_place(found, file.path(run_base, basename(found)))
}

# THE RUN DIR HOLDS EVERY SOURCE FOLDER, NOT ONLY THE ONES THE MANIFEST NAMES.
# A session narrows the registry to the suite it is probing, and a narrowed
# registry drops every other suite's `- module:` rows, so those folders leave
# code-tables.tsv and the loop above never copies them. The run dir then holds
# only the folders the last few narrowings happened to name -- four of nineteen,
# when this was found. The next run under a DIFFERENT narrowing created that
# suite's files from nothing, which gave every one of them a new inode, and
# macOS asked the operator for access to each file again. That is why the grant
# prompts kept coming back after the run dir itself was made stable: the folder
# was stable, the files inside it were not.
#
# So every folder is copied every run, whatever the manifest says. It costs a
# few hundred small files and nothing else, for the reason already given above:
# Development.ImportAll imports what the manifest names, never what the folder
# holds, so a folder that is not under probe is copied, ignored and never run.
copy_every_subfolder <- function(bases, run_base) {
  for (b in bases) {
    if (!dir.exists(b)) next
    for (sub in list.dirs(b, full.names = TRUE, recursive = FALSE)) {
      dir.create(run_base, recursive = TRUE, showWarnings = FALSE)
      copy_tree_in_place(sub, file.path(run_base, basename(sub)))
    }
  }
}
for (tag in setdiff(names(tag_srcs), "general forms")) {
  copy_every_subfolder(tag_srcs[[tag]], tag_runs[[tag]])
}

# THE MERGED FORMS ARE STAGED PER SUITE FOLDER, FROM ONE FLAT TREE.
# Development reads a form at forms/<folder>/<Name>.frm, the same shape it reads
# a class or a module at. The merged tree is flat, so it is copied whole into
# run/forms/<folder> for every folder the manifest names -- the rule the loop
# above follows, and for the same reason: a folder that drops out of a narrowed
# manifest keeps its files and their inodes, so macOS keeps its grants. A folder
# holding forms no table asks for costs a few small files; ImportAll imports
# what the manifest names.
forms_src <- tag_srcs[["general forms"]][1]
forms_run <- tag_runs[["general forms"]]
if (dir.exists(forms_src)) {
  for (suite_folder in unique(tbl$folder)) {
    dir.create(forms_run, recursive = TRUE, showWarnings = FALSE)
    copy_tree_in_place(forms_src, file.path(forms_run, suite_folder))
  }
} else {
  # A missing merged tree warns and lets the rest of the run go on, the way a
  # missing source folder does above. Only a suite carrying a `forms:` key
  # needs it, and --build is what writes it.
  message("run-tests.R: WARNING - no merged forms tree at ", forms_src,
          "; no .frm is staged. Re-run with --build to write it.")
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
# --- the headless suites: a repository mirror INSIDE the run dir (macOS) -----
# THIS IS A SANDBOX FIX, SO IT IS macOS ONLY. Windows has no sandbox: Excel
# reads and writes the repository directly, nothing is ever prompted for, and
# staging 4 MB of binaries per run would buy nothing. There the two headless
# suites keep doing what they always did -- walk up to the repository root and
# read it in place.
#
# On macOS that walk is what broke. TestHeadlessLinelistBuild and
# TestHeadlessSetupFill drive a real designer build, and every path they use
# hangs off a repository root found by walking up from the running workbook to
# the folder holding src/tests/test-registry.yml. That worked while staging sat
# in the repo. Since staging moved into Excel's container the walk finds
# nothing in eight levels and the modules fall back to a hard-coded path, so
# the build reads and writes OUTSIDE the sandbox.
#
# Nothing can grant that for them. Application.GrantAccessToMultipleFiles
# raises 438 on this host -- that is the call behind
# HeadlessBuild.EnsureFileAccess -- so no grant is made and the build carries
# on into CodeTransfer, which creates two files per component under OBTApp_.
# Every one of those is a fresh file with no bookmark, so it is one modal per
# file with nobody to click it and the run stops dead.
#
# So the handful of paths those suites touch are mirrored under the run dir in
# the SAME relative layout the repository uses. Every path expression in the
# modules keeps working; only what it hangs from moves. Inside the container
# none of it is asked for. The mirror is dropped again once the run is read --
# see the cleanup below.
headless_root <- file.path(run_dir, "headless")

if (!on_windows) {
  headless_inputs <- list(
    c(".mock", "designer_mock.xlsb"),
    c("ribbons", "_ribbontemplate_dev.xlsb"),
    c("src", "bin", "setup", "setup_dev.xlsb"),
    c("src", "tests", ".input", "package", "generic-test-setup.xlsb"),
    c("scripts", "headless", "vba", "OBTSetupImportHeadless.bas")
  )
  for (parts in headless_inputs) {
    src_f <- do.call(file.path, as.list(c(repo_root, parts)))
    dst_f <- do.call(file.path, as.list(c(headless_root, parts)))
    dir.create(dirname(dst_f), recursive = TRUE, showWarnings = FALSE)
    if (file.exists(src_f)) {
      file.copy(src_f, dst_f, overwrite = TRUE)
    } else {
      message("run-tests.R: WARNING - headless input missing, the build will ",
              "name it: ", src_f)
    }
  }

  # The class and module folders the transfer re-imports into the designer copy.
  # These lists match HeadlessBuild's TRANSFER_CLASS_FOLDERS and
  # TRANSFER_MODULE_FOLDERS. A folder added there belongs here too, or the build
  # imports fewer components than it used to and says so in its own report.
  for (folder in c("analyses", "dataio", "dictionary", "general", "geo",
                   "graphs", "linelist", "sections", "showhide")) {
    copy_tree_in_place(file.path(repo_root, "src", "classes", folder),
                       file.path(headless_root, "src", "classes", folder))
  }
  copy_tree_in_place(file.path(repo_root, "src", "modules", "linelist"),
                     file.path(headless_root, "src", "modules", "linelist"))

  # The forms the build imports. The suites read them from
  # .test-runner/forms/merged, which is where the merge wrote until staging
  # moved. The merge writes into the container now, so the mirror carries the
  # CURRENT merged tree under that same relative name. Reading the live tree is
  # the point: the copy still sitting in the repo is stale, and four of its
  # twelve forms no longer match the FormLogic source they were merged from.
  if (dir.exists(merged_forms)) {
    copy_tree_in_place(merged_forms,
                       file.path(headless_root, ".test-runner", "forms", "merged"))
  }

  # Where the build WRITES: the filled setup, the linelist, the trace, and the
  # OBTApp_ scratch folder CodeTransfer exports every component into.
  dir.create(file.path(headless_root, ".obt", "draft"), recursive = TRUE,
             showWarnings = FALSE)

  message("run-tests.R: headless mirror staged -> ", headless_root)
}

message("run-tests.R: run dir assembled from src/ (classes/, tests/, forms/, .generated/, bootstrap/)")

# --- 2a) ensure Excel is fully closed before we launch a fresh instance ------
# `run VB macro` returns AppleScript Parameter error (-50) against an Excel that
# is already open interactively, so quit any running instance and wait until the
# process is gone. Only quit when it is actually running (a bare `tell ... quit`
# would auto-launch Excel just to close it).
# macOS only. On Windows CreateObject("Excel.Application") makes its own
# instance, and an Excel the operator has open is none of this script's business.
# shQuote, and the space in the name is why. system2() hands its arguments to
# the shell UNQUOTED, so `c("-x", "Microsoft Excel")` runs as
# `pgrep -x Microsoft Excel` -- two operands, no match, exit 1. The guard
# answered FALSE against a running Excel every single time, so nothing was ever
# quit here and the operator had to do it by hand. Measured 2026-08-15 while
# chasing the same line in build-linelist.R.
excel_running <- function() {
  if (on_windows) return(FALSE)
  identical(system2("pgrep", c("-x", shQuote("Microsoft Excel")),
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
message("run-tests.R: launching Excel via ",
        if (on_windows) "cscript" else "osascript", " (", build_flag, ") ...")
trigger_rc <- if (on_windows) {
  system2("cscript", c("//nologo", shQuote(trigger), shQuote(work_copy), build_flag))
} else {
  system2("osascript", c(shQuote(trigger), shQuote(work_copy), build_flag))
}

# --- 4) collect + summarise --------------------------------------------------
# Helper: dump the VBA-side diagnostics log if the wrappers wrote one.
show_import_log <- function() {
  # The file is always present now (emptied in place before the run), so an
  # empty one is what "the wrappers wrote nothing" looks like.
  if (file.exists(log_path) && file.size(log_path) > 0L) {
    message("\n       --- obt-import.log (OBTBuildCodeTables/OBTSilentImport) ---")
    for (ln in readLines(log_path, warn = FALSE)) message("       ", ln)
    message("       --- end log ---")
  } else {
    message("\n       (no obt-import.log written — the wrappers may not have run, ",
            "or could not write to ", run_dir, ")")
  }
}

# Helper: drop the headless mirror, keeping the one thing worth keeping.
#
# The mirror is staged fresh every run (macOS only, see above), so nothing in
# it needs to survive -- except the build's own trace. That trace is the only
# record of WHERE a headless build stopped: the suite reports the fault through
# an error boundary that strips the location, so a red run reads as an opaque
# message while the trace has the step list. It is rescued to the run dir root,
# beside the results, and the rest of the mirror goes.
#
# Called on both exits. A green run has no use for 4 MB of copied binaries and
# a built linelist; a red one is read from the trace and the results, not from
# the mirror.
clear_headless_mirror <- function() {
  if (!dir.exists(headless_root)) return(invisible(NULL))

  traces <- list.files(file.path(headless_root, ".obt", "draft"),
                       pattern = "trace\\.txt$", full.names = TRUE)
  for (tr in traces) {
    file.copy(tr, file.path(run_dir, paste0("headless-", basename(tr))),
              overwrite = TRUE)
  }
  if (length(traces)) {
    message("run-tests.R: headless trace kept -> ",
            file.path(run_dir, paste0("headless-", basename(traces[1]))))
  }

  unlink(headless_root, recursive = TRUE, force = TRUE)
  message("run-tests.R: headless mirror cleared")
}

# Helper: keep the stale copy, explain, and exit non-zero.
fail <- function(msg) {
  message("\n[FAIL] ", msg)
  clear_headless_mirror()
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
# Emptied in place before the run, so it always exists. An empty one means the
# run wrote nothing, which is the same signal a missing one used to carry.
if (!file.exists(csv_path) || file.size(csv_path) == 0L) {
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
clear_headless_mirror()
message("run-tests.R: all tests passed. Run dir left at ", run_dir)
quit(status = 0L, save = "no")
