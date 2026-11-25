#!/usr/bin/env Rscript
# What the cached channel is, and whether obt can drive it.
#
#   Rscript gates/channel-age.R <channel-folder>
#
# Prints KEY=value lines and exits 0 whatever it finds. Deciding what to do
# about a channel is the case's job; this only reports.
#
# WHY IT ASKS obt RATHER THAN DECIDING FOR ITSELF
# -----------------------------------------------------------------------------
# The rule for "can this channel be driven" lives in obt:::check_channel_version
# and nowhere else. Reimplementing it here would mean a second copy that drifts
# the day obt raises its minimum, and the drift would show as a failed run
# that looks like a product fault. So this reads obt's own constants -- the
# moving-version label, the version pattern, the minimum -- through ::: and
# applies them in the same order obt does. Reaching into a namespace is
# deliberate: a wrong answer here is worse than a tight coupling.
#
# THE VERSION IS A FILENAME, NOT A CELL
# -----------------------------------------------------------------------------
# obt:::channel_version() takes everything after the first "-" of the designer's
# basename: designer_dev-latest.xlsb is "latest", designer_dev-2026-06-11.xlsb
# is "2026-06-11". So this needs no Excel and opens nothing.

suppressWarnings(suppressMessages({
    ok <- requireNamespace("obt", quietly = TRUE)
}))

args <- commandArgs(trailingOnly = TRUE)
folder <- if (length(args) >= 1) args[[1]] else ""

say <- function(key, value) cat(sprintf("%s=%s\n", key, value), sep = "")

if (!ok) {
    say("CHANNEL_ERROR", "obt is not installed")
    quit(status = 0)
}

moving <- get("CHANNEL_MOVING_VERSION", asNamespace("obt"))
pattern <- get("CHANNEL_VERSION_PATTERN", asNamespace("obt"))
minimum <- get("OBT_MINIMUM_VERSION", asNamespace("obt"))
version_of <- get("channel_version", asNamespace("obt"))

say("CHANNEL_MINIMUM", minimum)
say("CHANNEL_MOVING", moving)

if (!dir.exists(folder)) {
    say("CHANNEL_ERROR", sprintf("no such folder: %s", folder))
    quit(status = 0)
}

# THE ZIP IS WHAT obt READS, so the designer inside it is the one that decides,
# not a loose .xlsb sitting beside it in the same folder.
zips <- list.files(folder, pattern = "[.]zip$", full.names = TRUE)
if (length(zips) == 0) {
    say("CHANNEL_ERROR", "the folder holds no zip")
    quit(status = 0)
}

zip <- zips[[1]]
say("CHANNEL_ZIP", basename(zip))

listing <- tryCatch(utils::unzip(zip, list = TRUE), error = function(e) NULL)
if (is.null(listing) || nrow(listing) == 0) {
    say("CHANNEL_ERROR", "the zip could not be listed")
    quit(status = 0)
}

designers <- listing$Name[grepl("^designer", basename(listing$Name), ignore.case = TRUE)]
if (length(designers) == 0) {
    say("CHANNEL_ERROR", "the zip holds no designer")
    quit(status = 0)
}

designer <- basename(designers[[1]])
version <- version_of(designer)
say("CHANNEL_DESIGNER", designer)
say("CHANNEL_VERSION", if (is.na(version)) "" else version)

# obt's own order: the dev moving label passes unconditionally, then a
# semantic version at or above the minimum passes, and anything else cannot be
# driven at all.
drivable <- FALSE
why <- ""
if (!is.na(version) && identical(tolower(version), moving)) {
    drivable <- TRUE
    why <- "the dev moving build"
} else if (!is.na(version) && grepl(pattern, version)) {
    if (utils::compareVersion(version, minimum) >= 0) {
        drivable <- TRUE
        why <- sprintf("semantic version at or above %s", minimum)
    } else {
        why <- sprintf("semantic version below the minimum of %s", minimum)
    }
} else if (is.na(version)) {
    why <- "its designer carries no version in its name"
} else {
    why <- sprintf("%s is not a version obt can read; it predates the semantic scheme", version)
}

say("CHANNEL_DRIVABLE", if (drivable) "yes" else "no")
say("CHANNEL_WHY", why)

# THE AGE, which the version cannot give. A dev channel is stamped "latest"
# however old it is, so the only honest age signal is the date of what is
# inside the zip.
stamps <- suppressWarnings(as.POSIXct(listing$Date))
stamps <- stamps[!is.na(stamps)]
if (length(stamps) > 0) {
    say("CHANNEL_NEWEST_EPOCH", as.integer(max(stamps)))
    say("CHANNEL_NEWEST_DATE", format(max(stamps), "%Y-%m-%d"))
}

quit(status = 0)
