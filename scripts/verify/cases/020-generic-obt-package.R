#!/usr/bin/env Rscript
# The obt-package half of verify case 020.
#
#   Rscript 020-generic-obt-package.R <work-folder> <setup> <channel-dir> <name>
#
# Builds one linelist from the generic setup THROUGH THE PACKAGE, the way a
# user does, and prints what it produced. The shell case asserts on the files.
#
# WHY A SEPARATE FILE
# -----------------------------------------------------------------------------
# The package is driven in R and the assertions are in shell. Keeping the R
# here rather than in a heredoc means it can be read, linted and run on its
# own, and the case file stays a list of assertions.
#
# WHY THE CHANNEL IS COPIED IN FIRST
# -----------------------------------------------------------------------------
# obt_designer_add downloads a channel zip into <folder>/obt/<channel>/ and
# REUSES one already there. So a cached zip copied into that folder before the
# commit turns a network step into a file copy, which is what lets this case
# run on a machine with no network and never depend on the release being up.
# This is the pattern of obt's own .r-agent/draft/kit-test/run-obt-generate.R.

args <- commandArgs(trailingOnly = TRUE)
if (length(args) < 4) {
  cat("usage: 020-generic-obt-package.R <work> <setup> <channel-dir> <name>\n")
  quit(status = 2)
}

work <- args[[1]]
setup <- normalizePath(args[[2]], mustWork = TRUE)
channel_src <- normalizePath(args[[3]], mustWork = TRUE)
name <- args[[4]]

suppressPackageStartupMessages(library(obt))

cat("obt version:", as.character(packageVersion("obt")), "\n")
cat("work       :", work, "\n")
cat("setup      :", setup, "\n")
cat("channel src:", channel_src, "\n")

dir.create(work, recursive = TRUE, showWarnings = FALSE)

# 1. THE CHANNEL, COPIED IN so obt_designer_add reuses it.
channel <- file.path(work, "obt", "dev")
dir.create(channel, recursive = TRUE, showWarnings = FALSE)
copied <- list.files(channel_src, pattern = "[.](xlsb|xlsm|zip)$", full.names = TRUE)
if (!length(copied)) {
  cat("FATAL: the channel folder holds no workbook or zip\n")
  quit(status = 2)
}
invisible(file.copy(copied, channel, overwrite = TRUE))
cat("channel    :", length(copied), "file(s) copied into", channel, "\n")

# 2. THE SETUP into the working folder. obt_setup starts a recipe from a setup
#    workbook; overwrite lets this case re-run over its own previous work.
cat("\n--- obt_setup ---\n")
obt_commit(obt_setup(from = setup, folder = work, overwrite = TRUE))

# 3. THE RELEASE, and its designer quieted. A designer that still asks
#    questions sits on a message box no script can answer.
cat("\n--- obt_designer_add ---\n")
obt_commit(obt(folder = work) |> obt_designer_add(type = "dev"))

designer <- list.files(
  channel,
  pattern = "^designer[-_].*[.](xlsb|xlsm)$",
  full.names = TRUE
)
if (length(designer)) {
  cat("quieting   :", basename(designer[[1]]), "\n")
  obt_silent_set(normalizePath(designer[[1]]))
}

# 4. THE LANGUAGE the setup itself declares, read through the package rather
#    than assumed. obt_setup_languages answers the columns of the setup's own
#    translation table, and the first is the one a user would pick.
language <- tryCatch(obt_setup_languages(setup)[[1]], error = function(e) "English")
cat("language   :", language, "\n")

# 5. THE BUILD.
cat("\n--- obt_designer_generate ---\n")
run <- obt(folder = work) |>
  obt_designer_add(type = "dev") |>
  obt_designer_languages(dict = language, form = "ENG") |>
  obt_designer_generate(name = name, overwrite = TRUE)

obt_commit(run)

# 6. WHAT LANDED. The shell case asserts on these paths, so they are printed
#    rather than judged here.
cat("\n--- delivered ---\n")
produced <- list.files(file.path(work, "linelist"), full.names = TRUE)
if (!length(produced)) {
  cat("NOTHING in", file.path(work, "linelist"), "\n")
  quit(status = 1)
}
for (p in produced) {
  cat(sprintf("%10.0f  %s\n", file.info(p)$size, p))
}
cat("OBT_VERIFY_LINELIST=", produced[grepl("[.](xlsm|xlsb)$", produced)][1], "\n", sep = "")
quit(status = 0)
