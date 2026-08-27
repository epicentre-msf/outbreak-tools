# build-registry.R
# =============================================================================
# Parse src/tests/test-registry.yml (the single source of truth for the test
# suite AND the Codes-worksheet ListObjects) and emit two flat intermediate
# files that a VBA routine (OBT_BuildCodeTables) consumes to rebuild the tables
# and the `ModulesForTesting` Name.
#
#   src/tests/.generated/code-tables.tsv     one row per component:
#                                            folder <TAB> tag <TAB> component <TAB> interface
#   src/tests/.generated/modules-for-testing.txt   flat list of every test module
#
# YAML is parsed here (not in VBA — VBA YAML parsing is too fragile) matching
# the scripts/devtools/*.R culture. The ListObject API only exists in VBA, so
# the actual table build happens there; this script only flattens the registry.
#
# Usage:  Rscript scripts/tests/build-registry.R
#         (run from anywhere — paths resolve against the git root)
#
# NOTE: nothing here is run against Excel. See src/tests/automated-testing-macos.md
#       for how this fits the harness.
# =============================================================================

# --- a text locale, whatever the shell arrived with --------------------------
# THE REGISTRY IS UTF-8 AND THE SHELL MAY NOT BE.
# -----------------------------------------------------------------------------
# `read_yaml` goes through `readLines`, and in the C locale `readLines` cannot
# decode a multi-byte character: it warns "invalid input found on input
# connection", hands back a mangled line, the YAML then parses to nothing, and
# this script stops with "no `suites:` found" on a file whose `suites:` is
# sitting there in plain sight. Measured on the same file: C locale gives 0
# suites, en_US.UTF-8 gives 15.
#
# Shells that arrive with LC_CTYPE=C or nothing at all: a Terminal profile
# without "Set locale environment variables on startup", cron, CI runners, an
# editor's task shell, an ssh session forwarding nothing.
#
# Three candidate names because no single one exists on every machine. All
# three resolve on macOS. `fileEncoding = "UTF-8"` on the read was tried and
# does NOT work -- in a C locale the conversion target cannot hold the
# character either.
for (loc in c("en_US.UTF-8", "C.UTF-8", "UTF-8")) {
  if (nzchar(suppressWarnings(Sys.setlocale("LC_CTYPE", loc)))) break
}

# --- deps --------------------------------------------------------------------
used_packages <- "yaml"
to_install <- setdiff(used_packages, rownames(installed.packages()))
if (length(to_install)) install.packages(to_install, repos = "https://cloud.r-project.org")
suppressPackageStartupMessages(library(yaml))

# --- locate repo root + paths ------------------------------------------------
repo_root <- tryCatch(
  system2("git", c("rev-parse", "--show-toplevel"), stdout = TRUE),
  warning = function(w) NA_character_
)
if (length(repo_root) != 1L || is.na(repo_root) || !nzchar(repo_root)) {
  stop("build-registry.R: not inside a git repository (could not resolve repo root).")
}

registry_path <- file.path(repo_root, "src", "tests", "test-registry.yml")
out_dir       <- file.path(repo_root, "src", "tests", ".generated")
tsv_path      <- file.path(out_dir, "code-tables.tsv")
list_path     <- file.path(out_dir, "modules-for-testing.txt")

if (!file.exists(registry_path)) {
  stop("build-registry.R: registry not found at ", registry_path)
}
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

# --- parse -------------------------------------------------------------------
registry <- yaml::read_yaml(registry_path)
suites <- registry$suites
if (is.null(suites) || !length(suites)) {
  stop("build-registry.R: no `suites:` found in ", registry_path)
}

# tag mapping (see repo-knowledge/testing.md section 5)
#   classes  -> "general classes"   modules -> "general modules"
#   fixtures -> "tests classes"      tests   -> "tests modules"
#   forms    -> "general forms"
yn <- function(flag) if (isTRUE(flag)) "yes" else "no"

rows <- list()          # each: c(folder, tag, component, interface)
test_modules <- character(0)

for (suite in suites) {
  folder <- suite$folder
  if (is.null(folder) || !nzchar(folder)) {
    stop("build-registry.R: a suite is missing its `folder:` key.")
  }

  # general classes (may carry an interface)
  for (cls in suite$classes) {
    rows[[length(rows) + 1L]] <- c(folder, "general classes", cls$name, yn(cls$interface))
  }

  # general modules (plain strings, never an interface)
  for (mod in suite$modules) {
    rows[[length(rows) + 1L]] <- c(folder, "general modules", mod, "no")
  }

  # test-only fixture classes
  for (fx in suite$fixtures) {
    rows[[length(rows) + 1L]] <- c(folder, "tests classes", fx$name, yn(fx$interface))
  }

  # userforms — the .frm files come from the merged forms tree that
  # merge-form-code.R writes. The tag `general forms` means "import this .frm
  # from disk". The older `general form modules` tag means "copy a module's
  # code into a form already in the workbook", and it stays as it was.
  for (frm in suite$forms) {
    rows[[length(rows) + 1L]] <- c(folder, "general forms", frm$name, "no")
  }

  # test modules — also feed the ModulesForTesting Name
  for (t in suite$tests) {
    rows[[length(rows) + 1L]] <- c(folder, "tests modules", t$module, "no")
    test_modules <- c(test_modules, t$module)
  }
}

if (!length(rows)) stop("build-registry.R: registry parsed but produced no components.")

# --- emit --------------------------------------------------------------------
tbl <- as.data.frame(do.call(rbind, rows), stringsAsFactors = FALSE)
names(tbl) <- c("folder", "tag", "component", "interface")

# TAB-separated, no quoting, LF endings, header row — trivial for VBA to Split().
write.table(
  tbl, file = tsv_path, sep = "\t",
  quote = FALSE, row.names = FALSE, col.names = TRUE, eol = "\n"
)

writeLines(unique(test_modules), con = list_path, sep = "\n")

cat(sprintf(
  "build-registry.R: wrote %d component rows and %d test modules\n  %s\n  %s\n",
  nrow(tbl), length(unique(test_modules)), tsv_path, list_path
))
