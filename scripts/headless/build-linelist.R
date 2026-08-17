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
# ONE ROOT, AND IT IS INSIDE EXCEL'S OWN SANDBOX
# -----------------------------------------------------------------------------
# Excel for Mac is sandboxed, and touching a path it holds no permission for
# pops a dialog that a headless run has nobody to answer. So the run stages
# itself somewhere Excel already owns, where there is nothing to ask about:
#
#   ~/Library/Containers/com.microsoft.Excel/Data/Documents/OBTHome
#
# Everything Excel reads or writes lives under that one root, OBT_HOME:
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
# template, geobase) IN beforehand and copies the finished files back OUT to
# --out afterwards. A successful run then clears all five folders; a failed one
# keeps them, because that is the only account of what went wrong.
#
# The catch is on THIS side of the fence: macOS keeps Excel's container away
# from other programs, so R needs Full Disk Access to reach it. One toggle per
# machine, and it is the deployed install step.
#
# WITHOUT IT the run falls back to <repo>/headless-runner and somebody has to
# pick that folder in Excel BEFORE EVERY BUILD. A pick covers ONE build:
# measured 2026-08-15, a pick lets Excel write over files that were already
# there when it was made, and Excel remakes both output workbooks on every save,
# so the next run is creating files that did not exist at pick time and Excel
# asks about all of them. Reads are different -- a pick covers those anywhere
# below the folder, lastingly, which is why 139 staged source files are deleted
# and recopied every run and never prompt. Full write-up in
# .obt/gotchas/macos-sandbox-grant.md.
#
# The fallback name carries NO LEADING DOT, and that is not cosmetic: macOS
# hides dot-folders in the folder picker, so an operator asked to pick one
# cannot see it. That cost a run on 2026-08-14 -- the picker opened, the folder
# was invisible, something else got picked, and the import died with a Device
# I/O error.
#
# OBT_HOME is --obt-home, else the OBT_HOME environment variable (an .Renviron
# entry is the deployed way to set it), else the container, else the fallback.
# =============================================================================

# --- a text locale, whatever the shell arrived with --------------------------
# EVERY FILE THIS BUILD READS BACK CAN CARRY A NON-ASCII CHARACTER, and in the C
# locale `readLines` cannot decode one: it warns, hands back a mangled line, and
# whatever was reading that line draws its conclusion from the wreckage. The
# generation log carries Excel's own error text, which on a French install is
# accented, and this script reads that log to judge whether the run was good.
# The registry the child script parses is UTF-8 as well, and a C locale there
# stops the build before Excel is ever launched.
#
# Shells that arrive with LC_CTYPE=C or nothing at all: a Terminal profile
# without "Set locale environment variables on startup", cron, CI runners, an
# editor's task shell, an ssh session forwarding nothing. Set here as well as in
# the wrapper, because the wrapper is optional and this file is not.
for (loc in c("en_US.UTF-8", "C.UTF-8", "UTF-8")) {
  if (nzchar(suppressWarnings(Sys.setlocale("LC_CTYPE", loc)))) break
}

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

# Needed this early because the sandbox check below is macOS-only.
on_windows <- identical(.Platform$OS.type, "windows")

# THE ONE ROOT. Everything Excel reads or writes goes under it, and the best
# place for it is INSIDE EXCEL'S OWN SANDBOX, where Excel needs no permission
# for anything and never puts a panel on the screen.
#
#   ~/Library/Containers/com.microsoft.Excel/Data/Documents/OBTHome
#
# Measured 2026-08-15, and it is the whole answer to the grant problem. A build
# run there took 35 seconds, showed nothing, and left Excel's own grant file
# untouched to the byte -- Excel rewrites that file whenever it records an
# answered panel, so an unchanged one is proof it asked for nothing.
#
# The catch is on THIS side of the fence: macOS keeps Excel's container away
# from other programs, so R can only reach it once the operator has given their
# terminal app Full Disk Access (System Settings -> Privacy & Security). That is
# one toggle per machine and it is the deployed install step.
#
# Where the operator has not done it, the run falls back to a plain folder in
# the repo and the old rules apply: somebody picks that folder in Excel, and the
# pick lasts exactly one build (see the note on the marker below).
#
# --obt-home and OBT_HOME both override, and an explicit choice is never
# second-guessed. Sys.getenv reads .Renviron, which is how a deployed install
# pins a folder without anyone having to remember a flag.
excel_container_home <- function() {
  if (on_windows) return(NA_character_)
  root <- path.expand("~/Library/Containers/com.microsoft.Excel/Data/Documents")
  if (!dir.exists(root)) return(NA_character_)
  # Written rather than merely tested: dir.exists() answers TRUE for a folder
  # macOS will refuse to open, so only a real write settles it.
  home  <- file.path(root, "OBTHome")
  probe <- file.path(home, ".reachable")
  ok <- tryCatch({
    dir.create(home, recursive = TRUE, showWarnings = FALSE)
    writeLines("ok", probe)
    identical(readLines(probe, warn = FALSE), "ok")
  }, error = function(e) FALSE, warning = function(w) FALSE)
  unlink(probe, force = TRUE)
  if (isTRUE(ok)) home else NA_character_
}

obt_home_chosen <- opt("obt-home", default = {
  if (nzchar(Sys.getenv("OBT_HOME"))) Sys.getenv("OBT_HOME") else ""
})
in_container <- FALSE

if (nzchar(obt_home_chosen)) {
  obt_home <- obt_home_chosen
} else {
  container <- excel_container_home()
  if (!is.na(container)) {
    obt_home <- container
    in_container <- TRUE
  } else {
    obt_home <- file.path(repo_root, "headless-runner")
  }
}
if (!grepl("^(/|~|[A-Za-z]:)", obt_home)) obt_home <- file.path(repo_root, obt_home)
obt_home <- path.expand(obt_home)
dir.create(obt_home, recursive = TRUE, showWarnings = FALSE)
obt_home <- normalizePath(obt_home, mustWork = TRUE)

# An overridden path can still land inside the container, and then it needs no
# pick either. Decided on the resolved path rather than on how it was chosen.
if (!on_windows) {
  container_root <- path.expand("~/Library/Containers/com.microsoft.Excel/Data")
  in_container <- startsWith(paste0(obt_home, "/"), paste0(container_root, "/"))
}

# --- the folder pick, needed only OUTSIDE the container ----------------------
# Inside Excel's container there is nothing to grant and nothing to pick, so
# this whole section is skipped and no marker is ever read or written.
#
# Outside it, Excel for Mac gets access to a folder one way: somebody PICKS it
# in a folder dialog, in an Excel they opened themselves.
# (Application.GrantAccessToMultipleFiles raises 438 here and has never worked.)
#
# THE PICK CANNOT BE DONE FROM THIS SCRIPT. Two reasons, either one enough:
#
#   * The trigger has to OPEN the driver workbook before it can run any macro,
#     and opening it is itself a file read. The first panel therefore arrives
#     before any grant code could run at all.
#   * A pick made inside an Excel driven by Apple Events did not take. The pick
#     that works was made by hand, F5 in the VBE, in an Excel the operator had
#     opened. The project already knew a `choose folder` sent from OUTSIDE Excel
#     does nothing; an Excel driven wholly from outside behaves the same way.
#
# AND A PICK ONLY LASTS ONE BUILD, which is why the fallback is a fallback and
# not the design. Measured 2026-08-15 across four runs: a pick lets Excel write
# over files that were already there when it was made, and Excel remakes both
# output workbooks on every save, so the next run is writing files that did not
# exist at pick time and asks about every one of them. Excel's own grant file
# shrank back to the same byte count after each run, which is it discarding what
# the pick had added. The marker below therefore records that somebody was told
# what to do, not that the run will be silent.
#
# The marker is written by --granted and never by a finished build. A build can
# finish green while somebody clicks through eighty panels, so success does not
# mean access was granted.
granted_marker <- file.path(obt_home, ".granted")

if (flag("granted")) {
  if (in_container) {
    message("build-linelist.R: nothing to grant -- ", obt_home)
    message("build-linelist.R: that folder is inside Excel's own sandbox, so ",
            "no pick is needed and none is asked for.")
  } else {
    writeLines(c("Somebody picked this folder in Excel.",
                 "Delete this file to be shown the steps again."), granted_marker)
    message("build-linelist.R: wrote ", granted_marker)
  }
  quit(status = 0L, save = "no")
}

if (!on_windows && !in_container && !file.exists(granted_marker)) {
  message("\n=============================================================")
  message(" EXCEL HAS NO ACCESS TO THIS FOLDER")
  message("")
  message("   ", obt_home)
  message("")
  message(" THE BETTER FIX, and it is one toggle:")
  message("")
  message("   System Settings -> Privacy & Security -> Full Disk Access")
  message("   Add the app you run this from, then quit and reopen it.")
  message("")
  message(" That lets the build stage itself inside Excel's own sandbox,")
  message(" where Excel never asks for anything. No picking, ever.")
  message("")
  message(" OR, to keep building in the folder above, do this each time")
  message(" -- a pick only covers ONE build:")
  message("")
  message("   1. Open ", workbook_src, " in Excel")
  message("   2. Press Alt+F11 for the VBE")
  message("   3. Put the cursor inside OBTGrantAccess and press F5")
  message("   4. In the picker, choose the folder named above")
  message("   5. Quit Excel from its File menu")
  message("")
  message("   Rscript scripts/headless/build-linelist.R --granted")
  message("=============================================================\n")
  quit(status = 1L, save = "no")
}

# The trigger is the only OS-specific piece. Everything above and below it is
# the same on both hosts, which is the whole point of keeping it thin: the two
# triggers drive the SAME entry points in the SAME order with the SAME
# arguments, so a difference in what gets built is a difference in the VBA
# rather than in the harness.
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

message("build-linelist.R: OBT_HOME  -> ", obt_home,
        if (in_container) "  (inside Excel's sandbox, no grant needed)"
        else "  (outside Excel's sandbox, a pick covers ONE build)")
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
       ".\n  A successful run CLEARS the staging, so --no-merge on its own has ",
       "nothing left to reuse.\n  Drop --no-merge, or point --forms at a folder ",
       "that holds the .frm files.")
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

#
# THE EXTENSION IS THE OPERATOR'S, never a fixed one. A staged name used to end
# in .xlsb whatever arrived, and a geobase kept as .xlsx was handed to Excel
# under a name that said "binary workbook". Excel for Mac DIES on that file
# rather than refusing it: the build lost its whole process mid-run, the
# AppleScript read a dead connection and reported -609, and the phase it died in
# moved from run to run because it is whenever the geobase is first read.
stage_in <- function(from, stem) {
  if (!nzchar(from)) return("")
  ext <- sub("^.*\\.", "", basename(from))
  if (identical(ext, basename(from)) || !nzchar(ext)) ext <- "xlsb"
  to <- file.path(staged_in, paste0(stem, ".", tolower(ext)))
  if (!file.copy(from, to, overwrite = TRUE)) {
    stop("build-linelist.R: failed to stage ", from, " into ", staged_in)
  }
  to
}

staged_designer <- stage_in(designer_path, "designer")
staged_setup    <- stage_in(setup_path,    "setup")
staged_template <- stage_in(temp_path,     "template")
staged_geo      <- stage_in(geo_path,      "geo")

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

# The three files a build writes, and the two folders get opposite treatment.
#
#   dest_folder  DELETED. R writes there and R needs no grant.
#   build_out    KEPT. Excel writes there, and a file Excel has to CREATE pops a
#                macOS panel while a file it writes OVER does not.
#
# Measured 2026-08-15 with the operator reading the panels: the paths on them
# were out/<name>-designer.xlsb and out/OBTApp_/*, which is exactly what this
# loop used to delete plus the folder Excel makes for itself. The same run
# rebuilt all 139 staged source files from scratch and showed no panel at all,
# because Excel only READS those. Reads under a picked folder are free; creates
# are not.
#
# Deleting bought one thing worth keeping: a stale linelist left in place would
# let a failed run report a file on disk and read as a success. So the run
# stamps the clock here instead, and section 6 refuses any output older than it.
build_leaves <- c(".xlsb", "-generation.txt", "-designer.xlsb")

# THE PREVIOUS RUN'S FILES ARE MOVED ASIDE, NOT DROPPED.
# -----------------------------------------------------------------------------
# Every delivered name is built from --name, so two runs sharing a --name write
# over each other's whole output set: the linelist, the designer copy and all
# four logs. The default --name in the wrapper made that the normal case, and it
# cost real evidence -- a run investigating a failure destroyed the artefacts of
# the failure it was investigating, and the only reason its Excel version was
# ever recovered is that one earlier run happened to use a different name.
#
# So the set is moved into <name>-previous/ first. One generation, overwritten
# each time, which is the generation that matters: the run before this one is
# what a failure gets compared against. Disk use stays bounded, and a rebuild
# loop needs no extra flag.
archive_leaves <- c(build_leaves, "-import.log", "-build-report.txt", "-trace.txt")
prev_folder <- file.path(dest_folder, paste0(out_name, "-previous"))
same_folder <- identical(normalizePath(dest_folder, mustWork = FALSE),
                         normalizePath(build_out, mustWork = FALSE))

if (!same_folder) {
  standing <- file.path(dest_folder, paste0(out_name, archive_leaves))
  standing <- standing[file.exists(standing)]
  if (length(standing)) {
    unlink(prev_folder, recursive = TRUE, force = TRUE)
    dir.create(prev_folder, recursive = TRUE, showWarnings = FALSE)
    for (from in standing) {
      to <- file.path(prev_folder, basename(from))
      # Renamed where it can be, copied where it cannot: --out on another volume
      # makes a rename fail, and losing the file to a tidy-up is the one outcome
      # this block exists to prevent.
      if (!file.rename(from, to)) file.copy(from, to, overwrite = TRUE)
    }
    message("build-linelist.R: previous -> ", length(standing),
            " file(s) moved to ", prev_folder)
  }
}

for (leaf in build_leaves) {
  unlink(file.path(dest_folder, paste0(out_name, leaf)), force = TRUE)
}
run_started <- Sys.time()

# Only macOS needs this. There, `run VB macro` returns a Parameter error against
# an Excel that is already open interactively, so the instance has to go first.
# On Windows CreateObject("Excel.Application") makes its own instance and an
# Excel the operator has open is none of its business.
# --- 4b) the trigger's step 0 ------------------------------------------------
# Kept at "0", so the trigger never opens the folder picker. Inside the
# container there is nothing to pick. Outside it, a pick made in an Excel driven
# by Apple Events does not take, and the picking has to be done by hand anyway
# -- the block near the top of this file prints the steps and stops the run.
need_pick <- "0"

# shQuote, and the space in the name is why. system2() hands its arguments to
# the shell UNQUOTED, so `c("-x", "Microsoft Excel")` runs as
# `pgrep -x Microsoft Excel` -- two operands, no match, exit 1. The guard then
# answered FALSE against a running Excel every single time, which is worse than
# no guard: the trigger goes on to drive the operator's own instance, where
# `run VB macro` does nothing at all. The run then ends with an empty outcome, a
# zero-byte obt-import.log and no file written. Measured 2026-08-15.
excel_running <- function() {
  if (on_windows) return(FALSE)
  identical(system2("pgrep", c("-x", shQuote("Microsoft Excel")),
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
  staged_forms, build_out, out_name, build_options, report_path, obt_home,
  need_pick
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

# The build writes this one line at a time as it goes, so it is the only account
# of a run that took Excel down with it: the report carries the narrative the
# build HANDED BACK, and a build whose process died never handed anything back.
trace_path <- file.path(build_out, paste0(out_name, "-trace.txt"))

show_trace <- function() {
  if (!file.exists(trace_path) || file.size(trace_path) == 0L) return(invisible(NULL))
  message("\n       --- what the build reached, last line first ---")
  for (ln in rev(readLines(trace_path, warn = FALSE))) message("       ", ln)
  message("       --- end trace ---")
}

fail <- function(msg) {
  message("\n[FAIL] ", msg)
  show_trace()
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
platform <- field("platform")

if (length(narrative)) {
  message("\nWhat the build did:")
  for (ln in narrative) if (nzchar(trimws(ln))) message("  ", ln)
}

# Printed on every run, because a run that started prompting again is invisible
# from here otherwise: nothing in an exit code or an elapsed time can see a
# dialog, and the operator is the only one who can.
message("\nbuild-linelist.R: sandbox   -> ",
        if (in_container) "inside Excel's own sandbox, nothing to grant"
        else "OUTSIDE Excel's sandbox, a pick covers ONE build")
if (nzchar(grant)) message("build-linelist.R: trigger   -> ", grant)

# Printed on every run so a report kept or pasted elsewhere still says which
# platform produced it. The two do not behave the same.
if (nzchar(platform)) message("build-linelist.R: platform  -> ", platform)

message(sprintf("\nbuild-linelist.R: %s sheet(s), %s variable(s), %s component(s) re-imported.",
                field("sheets"), field("variables"), field("components")))

if (!identical(outcome, "OK")) {
  fail(paste0("the build answered: ", outcome))
}
if (!nzchar(built) || !file.exists(built)) {
  fail(paste0("the build answered OK and there is no file at: ", built))
}

# The output files are kept from one run to the next now (see section 4), so a
# file being there is no longer evidence THIS run wrote it. Every output must
# carry a timestamp later than the moment the clock was stamped, which is before
# Excel was launched. A file left by an earlier run fails the whole run, exactly
# as an empty out/ used to.
checked <- unique(c(built, file.path(build_out, paste0(out_name, build_leaves))))
stale   <- checked[file.exists(checked) & file.mtime(checked) < run_started]
if (length(stale)) {
  fail(paste0("the build answered OK and left ", length(stale),
              " file(s) from an earlier run, so this run did not write them: ",
              paste(basename(stale), collapse = ", ")))
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

# --- 7b) the run's own record ------------------------------------------------
# The two logs go out with the files. They live in the run dir, which is wiped
# below, and inside Excel's container that folder is somewhere most people will
# never look. They are also the only account of what the build actually did:
# obt-import.log lists every component that went into the driver, and
# build-report.txt is what the trigger recorded, narrative included.
#
# Named after the linelist rather than kept as bare log names, so a folder
# holding several builds still says which run each belongs to.
for (rec in list(c(log_path,    "-import.log"),
                 c(report_path, "-build-report.txt"),
                 c(trace_path,  "-trace.txt"))) {
  if (!file.exists(rec[1]) || file.size(rec[1]) == 0L) next
  to <- file.path(dest_folder, paste0(out_name, rec[2]))
  # The trace is already written under the output name, so a run whose --out IS
  # the build folder would copy the file onto itself and empty it.
  if (identical(normalizePath(rec[1]), normalizePath(to, mustWork = FALSE))) next
  if (file.copy(rec[1], to, overwrite = TRUE)) {
    message("build-linelist.R:          -> ", to)
  } else {
    message("build-linelist.R: WARNING, could not copy the log to ", to)
  }
}

# --- 7c) the generation log has the last word --------------------------------
# THE BUILD SAYING OK IS NOT THE SAME AS THE BUILD BEING GOOD.
# -----------------------------------------------------------------------------
# `outcome=OK` means BuildLinelistFromSetup ran to the end without raising. It
# says nothing about what the generation recorded on the way. A geobase Excel
# refuses is the case that proved it: the geo import logs an error, the build
# carries on, saves a linelist with no geographic data at all, answers OK, and
# exits 0. The operator gets a file that looks right and is missing a quarter of
# its content, and the only sign is a log line in a file they have no reason to
# open.
#
# Checked here, at the very end, so the linelist and every log are already in the
# operator's own folder before the run is failed. The evidence stays put: the run
# dir is kept as well, because fail() quits before the staging is cleared.
gen_path <- file.path(build_out, paste0(out_name, "-generation.txt"))

# Read as bytes and converted from latin1, which cannot fail. This log is
# Excel's own text and Excel for Mac writes it in a single-byte encoding: a
# French install puts a raw 0xE9 in it for the accented word in its own error
# messages, and that is not valid UTF-8. `readLines` would warn and hand back a
# mangled line. Every tag being ASCII, the scan is exact either way.
read_generation <- function(p) {
  if (!file.exists(p) || file.size(p) == 0L) return(character(0))
  txt <- rawToChar(readBin(p, "raw", file.size(p)))
  Encoding(txt) <- "latin1"
  strsplit(enc2utf8(txt), "\r\n|\n|\r")[[1]]
}

gen_lines <- read_generation(gen_path)

if (!length(gen_lines)) {
  fail(paste0("the build answered OK and wrote no generation log at: ", gen_path))
}

# A geobase that was asked for and is not in the log as imported. Checked on its
# own, because a future failure that logs nothing at all would slip past the
# error scan below while still shipping a linelist with no geography.
if (nzchar(geo_path)) {
  if (!any(grepl("was imported", gen_lines, fixed = TRUE))) {
    fail(paste0("a geobase was given (", geo_path,
                ") and the generation log never says it was imported, so the ",
                "linelist carries no geographic data."))
  }
  message("build-linelist.R: geobase  -> imported")
}

gen_errors <- grep("[Error]", gen_lines, fixed = TRUE, value = TRUE)
if (length(gen_errors)) {
  message("\n       --- the ", length(gen_errors), " error line(s) the generation recorded ---")
  for (ln in gen_errors) message("       ", ln)
  message("       --- end ---")
  fail(paste0("the build answered OK and the generation recorded ",
              length(gen_errors), " error(s). The files were delivered, ",
              "and they are not trustworthy."))
}

message("build-linelist.R: log      -> no errors recorded")

# --- 8) clear the staging ----------------------------------------------------
# Only on success, and only after everything is delivered. A run stages about
# 17 MB -- the source tree, the driver copy, the merged forms, copies of the
# operator's designer and setup, and the built workbooks -- and every byte of it
# is rebuilt from scratch by the next run, so keeping it buys nothing.
#
# A FAILED run keeps the lot. fail() prints the run dir and says so, and reading
# what a broken build left behind is how most of this file got written. It never
# reaches here.
staged <- c(run_dir, staged_forms, staged_in, build_out,
            file.path(source_root, "src"))
unlink(staged, recursive = TRUE, force = TRUE)
left <- staged[dir.exists(staged)]
if (length(left)) {
  message("build-linelist.R: WARNING, staging not fully cleared: ",
          paste(left, collapse = ", "))
} else {
  message("build-linelist.R: staging cleared from ", obt_home)
}

quit(status = 0L, save = "no")
