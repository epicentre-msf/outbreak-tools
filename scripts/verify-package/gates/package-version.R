#!/usr/bin/env Rscript
# Is the obt installed on this machine the latest one?
#
#   Rscript package-version.R [<obt-source-repo>]
#
# Prints KEY=value lines and exits 0 whatever it finds. Deciding what to do is
# the runner's job; this only reports.
#
# WHY THIS GATE EXISTS
# -----------------------------------------------------------------------------
# verify-package judges the PACKAGE, so a run over a stale obt answers a
# question nobody asked: it tells you last month's obt still works. Worse, a
# failure then reads as a product fault when the fix is already written and
# merely not installed here.
#
# Measured 2026-11-25: the installed obt was 0.0.1 at sha c2687b02 while
# D:/Projects/obt sat on 0.0.2, two commits ahead. Every verify-package result
# from that install would have described a package the author had already moved
# past.
#
# HOW "LATEST" IS DECIDED, AND WHY BOTH HALVES ARE NEEDED
# -----------------------------------------------------------------------------
# The DESCRIPTION Version alone is not enough: a package gets several commits
# per version bump, so 0.0.2 installed against 0.0.2 in source can still be
# behind by real changes. The installed DESCRIPTION records RemoteSha, which is
# the commit it was built from, so the sha is the precise half and the version
# is the readable half.
#
#   version differs                  -> not current, and which way is printed
#   version matches, sha differs     -> not current when source is ahead
#   source tree has uncommitted work -> cannot be current, because the install
#                                       came from a commit and the tree is not
#                                       at one
#
# A source repo that is not on this machine makes the answer UNKNOWN rather than
# "yes". Guessing "yes" is how a stale install goes unnoticed.

say <- function(key, value) cat(sprintf("%s=%s\n", key, value), sep = "")

args <- commandArgs(trailingOnly = TRUE)

installed_ok <- suppressWarnings(suppressMessages(
    requireNamespace("obt", quietly = TRUE)
))
if (!installed_ok) {
    say("PKG_CURRENT", "no")
    say("PKG_WHY", "obt is not installed on this machine at all")
    quit(status = 0)
}

installed_version <- as.character(utils::packageVersion("obt"))
say("PKG_INSTALLED_VERSION", installed_version)

desc <- file.path(find.package("obt"), "DESCRIPTION")
read_field <- function(path, field) {
    if (!file.exists(path)) return(NA_character_)
    value <- tryCatch(read.dcf(path, fields = field)[1, 1], error = function(e) NA_character_)
    if (is.na(value)) return(NA_character_)
    # unname: read.dcf hands back the field NAME as a names attribute, which
    # trimws keeps, and identical("obt" with names, "obt") is FALSE. That read
    # as "no obt source on this machine" when the repo was right there, and
    # would have made every version comparison below mismatch too.
    unname(trimws(value))
}

installed_sha <- read_field(desc, "RemoteSha")
say("PKG_INSTALLED_SHA", if (is.na(installed_sha)) "" else substr(installed_sha, 1, 12))

# THE SOURCE REPO. Given as an argument, or OBT_SOURCE, or the sibling of this
# repository, which is where it sits on the author's machine.
# This script sits at <repo>/scripts/verify-package/gates/, so the repository
# root is three levels up and the sibling checkout is beside it.
script_arg <- grep("^--file=", commandArgs(), value = TRUE)
here <- if (length(script_arg) == 0) {
    getwd()
} else {
    dirname(normalizePath(sub("^--file=", "", script_arg[[1]]),
                          winslash = "/", mustWork = FALSE))
}
repo_root <- normalizePath(file.path(here, "..", "..", ".."),
                           winslash = "/", mustWork = FALSE)
sibling <- file.path(dirname(repo_root), "obt")

candidates <- c(
    if (length(args) >= 1) args[[1]] else NULL,
    Sys.getenv("OBT_SOURCE", unset = NA_character_),
    sibling
)
candidates <- candidates[!is.na(candidates) & nzchar(candidates)]

source_repo <- NA_character_
for (cand in candidates) {
    if (dir.exists(cand) && file.exists(file.path(cand, "DESCRIPTION"))) {
        pkg <- read_field(file.path(cand, "DESCRIPTION"), "Package")
        if (identical(pkg, "obt")) {
            source_repo <- normalizePath(cand, winslash = "/")
            break
        }
    }
}

if (is.na(source_repo)) {
    say("PKG_CURRENT", "unknown")
    say("PKG_WHY", "no obt source repo on this machine to compare against; set OBT_SOURCE")
    quit(status = 0)
}

say("PKG_SOURCE_REPO", source_repo)
source_version <- read_field(file.path(source_repo, "DESCRIPTION"), "Version")
say("PKG_SOURCE_VERSION", if (is.na(source_version)) "" else source_version)

git <- function(...) {
    out <- suppressWarnings(tryCatch(
        system2("git", c("-C", shQuote(source_repo), ...), stdout = TRUE, stderr = FALSE),
        error = function(e) character()
    ))
    if (length(out) == 0) "" else trimws(out[[1]])
}

source_sha <- git("rev-parse", "HEAD")
say("PKG_SOURCE_SHA", if (nzchar(source_sha)) substr(source_sha, 1, 12) else "")

dirty <- git("status", "--porcelain")
say("PKG_SOURCE_DIRTY", if (nzchar(dirty)) "yes" else "no")

behind <- ""
if (nzchar(source_sha) && !is.na(installed_sha) && nzchar(installed_sha)) {
    counted <- git("rev-list", "--count", paste0(installed_sha, "..", source_sha))
    if (grepl("^[0-9]+$", counted)) behind <- counted
}
say("PKG_COMMITS_BEHIND", behind)

# THE VERDICT, in the order that gives the most useful reason first.
verdict <- "yes"
why <- "the installed obt matches its source"

if (!is.na(source_version) && !identical(source_version, installed_version)) {
    cmp <- utils::compareVersion(source_version, installed_version)
    verdict <- "no"
    why <- if (cmp > 0) {
        sprintf("obt %s is installed but the source is at %s", installed_version, source_version)
    } else {
        sprintf("obt %s is installed, AHEAD of the source at %s", installed_version, source_version)
    }
} else if (nzchar(behind) && as.integer(behind) > 0) {
    verdict <- "no"
    why <- sprintf("obt %s is installed from %s, %s commit(s) behind the source",
                   installed_version, substr(installed_sha, 1, 12), behind)
} else if (nzchar(dirty)) {
    verdict <- "no"
    why <- "the obt source tree has uncommitted changes, so no install can match it"
} else if (is.na(installed_sha) || !nzchar(installed_sha)) {
    verdict <- "unknown"
    why <- "the installed obt records no RemoteSha, so it cannot be placed against the source"
}

say("PKG_CURRENT", verdict)
say("PKG_WHY", why)
say("PKG_REINSTALL", sprintf("Rscript -e 'devtools::install(\"%s\")'", source_repo))
quit(status = 0)
