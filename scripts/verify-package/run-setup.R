#!/usr/bin/env Rscript
# One setup, driven all the way through the obt package.
#
#   Rscript run-setup.R <work> <setup> <channel-dir> <name> [<n-records>]
#
# Generates a linelist from the setup, makes a fake dataset out of the same
# setup, and imports that dataset into the linelist it just generated. Prints
# KEY=value markers the runner asserts on, and the stage name whenever one of
# them fails, so a red names the step rather than the file.
#
# WHY THE PACKAGE AND NOT build-linelist.R
# -----------------------------------------------------------------------------
# This is the path a USER takes. obt drives clickGenerate on a DOWNLOADED
# designer, so the binary under test is the released one rather than the mock
# with src/ imported over it. scripts/verify/ asks whether the product behaves;
# this asks whether the shipped package plus the shipped binaries do.
#
# WHY THE CHANNEL IS COPIED IN FIRST
# -----------------------------------------------------------------------------
# obt_designer_add downloads a channel zip into <folder>/obt/<channel>/ and
# REUSES one already there. A cached zip copied into that folder before the
# commit turns a network step into a file copy, so this never depends on GitHub
# being up. The runner has already checked that the zip is one obt can drive.
#
# WHY THE IMPORT IS FORCED
# -----------------------------------------------------------------------------
# obt_linelist_import reads a MIGRATION file and warns on three things: no
# metadata, no language recorded, or a different language. A person clicking the
# button reads the warning and decides; a scripted run has nobody to decide, so
# it stops and says which of the three stopped it.
#
# obt_fake writes a plain workbook through writexl::write_xlsx -- tables drawn
# from the setup, and no migration metadata or language anywhere in it. So it
# trips the first two warnings by construction, every time, and force = TRUE is
# structural here rather than a convenience. OBT_PACKAGE_FORCE=0 turns it off
# for anyone who wants to see the refusal itself.

args <- commandArgs(trailingOnly = TRUE)
if (length(args) < 4) {
  cat("usage: run-setup.R <work> <setup> <channel-dir> <name> [<n-records>]\n")
  quit(status = 2)
}

work <- args[[1]]
setup <- normalizePath(args[[2]], winslash = "/", mustWork = TRUE)
channel_src <- normalizePath(args[[3]], winslash = "/", mustWork = TRUE)
name <- args[[4]]
n_records <- if (length(args) >= 5) suppressWarnings(as.integer(args[[5]])) else 25L
if (is.na(n_records) || n_records < 1) n_records <- 25L

force_import <- !identical(Sys.getenv("OBT_PACKAGE_FORCE", unset = "1"), "0")

suppressPackageStartupMessages(library(obt))

say <- function(key, value) cat(sprintf("%s=%s\n", key, value), sep = "")

cat("obt version:", as.character(packageVersion("obt")), "\n")
cat("work       :", work, "\n")
cat("setup      :", basename(setup), "\n")
cat("records    :", n_records, "\n")
cat("force      :", force_import, "\n")

dir.create(work, recursive = TRUE, showWarnings = FALSE)

# A stage that fails names itself and stops the run. Without this a failure
# three verbs deep reads as "Error in obt_commit" and says nothing about which
# verb asked for the impossible.
stage <- function(label, expr) {
  cat("\n---", label, "---\n")
  withCallingHandlers(
    tryCatch(force(expr), error = function(e) {
      say("PKG_STAGE_FAILED", label)
      say("PKG_ERROR", conditionMessage(e))
      cat("\n")
      quit(status = 1)
    }),
    warning = function(w) {
      cat("warning:", conditionMessage(w), "\n")
      invokeRestart("muffleWarning")
    }
  )
}

# 1. THE CHANNEL, copied in so obt_designer_add reuses it.
channel <- file.path(work, "obt", "dev")
dir.create(channel, recursive = TRUE, showWarnings = FALSE)
copied <- list.files(channel_src, pattern = "[.](xlsb|xlsm|zip)$", full.names = TRUE)
if (!length(copied)) {
  say("PKG_STAGE_FAILED", "channel-copy")
  say("PKG_ERROR", "the channel folder holds no workbook or zip")
  quit(status = 1)
}
invisible(file.copy(copied, channel, overwrite = TRUE))
cat("channel    :", length(copied), "file(s) copied into", channel, "\n")

# 2. THE SETUP into the working folder.
stage("obt_setup", {
  obt_commit(obt_setup(from = setup, folder = work, overwrite = TRUE))
})

# 3. THE RELEASE, and its designer quieted. A designer that still asks
#    questions sits on a message box no script can answer.
stage("obt_designer_add", {
  obt_commit(obt(folder = work) |> obt_designer_add(type = "dev"))
})

designer <- list.files(channel, pattern = "^designer[-_].*[.](xlsb|xlsm)$", full.names = TRUE)
if (length(designer)) {
  cat("quieting   :", basename(designer[[1]]), "\n")
  obt_silent_set(normalizePath(designer[[1]], winslash = "/"))
}

# 4. THE LANGUAGE the setup declares, read through the package rather than
#    assumed. The first column of its own translation table is what a user picks.
language <- tryCatch(obt_setup_languages(setup)[[1]], error = function(e) "English")
cat("language   :", language, "\n")
say("PKG_LANGUAGE", language)

# 5. THE LINELIST.
stage("obt_designer_generate", {
  obt_commit(
    obt(folder = work) |>
      obt_designer_add(type = "dev") |>
      obt_designer_languages(dict = language, form = "ENG") |>
      obt_designer_generate(name = name, overwrite = TRUE)
  )
})

produced <- list.files(file.path(work, "linelist"), full.names = TRUE)
linelist <- produced[grepl("[.](xlsm|xlsb)$", produced)]
if (!length(linelist)) {
  say("PKG_STAGE_FAILED", "obt_designer_generate")
  say("PKG_ERROR", paste("nothing landed in", file.path(work, "linelist")))
  quit(status = 1)
}
linelist <- linelist[[1]]
say("PKG_LINELIST", linelist)
say("PKG_LINELIST_BYTES", file.info(linelist)$size)

# 6. THE FAKE DATASET, drawn from the same setup. obt_fake needs a recipe that
#    carries a setup, because the setup is what the records are read out of.
stage("obt_fake", {
  obt_commit(
    obt_setup(from = setup, folder = work, overwrite = TRUE) |>
      obt_fake(n = n_records, name = "fake", overwrite = TRUE, seed = 20261125L)
  )
})

fake <- list.files(file.path(work, "export"), pattern = "[.]xlsx$", full.names = TRUE)
if (!length(fake)) {
  say("PKG_STAGE_FAILED", "obt_fake")
  say("PKG_ERROR", paste("no fake workbook in", file.path(work, "export")))
  quit(status = 1)
}
fake <- fake[[1]]
say("PKG_FAKE", fake)
say("PKG_FAKE_BYTES", file.info(fake)$size)

# 7. THE IMPORT, into the linelist this run generated. obt_linelist takes the
#    working folder, the way the package's own example does, so it reads the
#    linelist already standing there rather than being handed a path.
stage("obt_linelist_import", {
  obt_commit(
    obt_linelist(obt(folder = work)) |>
      obt_linelist_import(from = fake, rule = "append", force = force_import)
  )
})

say("PKG_IMPORTED", "yes")
say("PKG_LINELIST_BYTES_AFTER", file.info(linelist)$size)
cat("\n")
quit(status = 0)
