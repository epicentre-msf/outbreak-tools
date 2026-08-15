#!/usr/bin/env Rscript
#
# vba-inspect.R -- read the VBA components, and their source, out of a
# vbaProject.bin.
#
#   Rscript scripts/devtools/vba-inspect.R <vbaProject.bin> <mode> [root ...]
#
#     mode = list     name every component and its kind
#     mode = content  also recover each component's source and compare it with
#                     the matching file under the given roots
#
# Writes one tab separated row per component: kind, name, path, status, detail.
# Called by mock-import-drift.sh; useful on its own when a single binary needs
# looking at.
#
# WHAT IS BEING READ
# -----------------------------------------------------------------------------
# vbaProject.bin is an OLE compound file. Two of its streams matter:
#
#   PROJECT        plain text, names every component and its kind --
#                  Class, Module, BaseClass (a form), Document (a sheet).
#   <module name>  one stream per component, holding a performance cache
#                  followed by the source compressed with the MS-OVBA
#                  run-length scheme.
#
# Both are read through the sector chains, never off the raw bytes. A stream is
# scattered across sectors that are not adjacent, and reading the file linearly
# splits a name in half at a sector boundary.
#
# The `dir` stream is decompressed too, for one thing only: the byte offset at
# which each module's source starts. Without it the source has to be found by
# trying every 0x01 byte in the stream, which works but is far slower. That scan
# is still there as a fallback for a module `dir` does not describe.
# =============================================================================

args <- commandArgs(trailingOnly = TRUE)
if (length(args) < 2L) {
  stop("usage: vba-inspect.R <vbaProject.bin> <list|content> [root ...]",
       call. = FALSE)
}
bin_path <- args[1]
mode <- args[2]
roots <- if (length(args) > 2L) args[-(1:2)] else character(0)

FREESECT <- c(4294967294, 4294967295) # ENDOFCHAIN and FREESECT

# --- little endian readers ---------------------------------------------------
# Doubles throughout: a 32 bit sector number does not fit an R integer.
u16 <- function(r, off) {
  b <- as.integer(r[off + 1:2])
  b[1] + b[2] * 256
}
u32 <- function(r, off) {
  b <- as.integer(r[off + 1:4])
  b[1] + b[2] * 256 + b[3] * 65536 + b[4] * 16777216
}
u32v <- function(r) {
  # every 4 byte group of a raw vector, as a numeric vector
  b <- as.integer(r)
  n <- length(b) %/% 4L
  i <- seq_len(n) * 4L - 3L
  b[i] + b[i + 1L] * 256 + b[i + 2L] * 65536 + b[i + 3L] * 16777216
}
u64 <- function(r, off) {
  b <- as.integer(r[off + 1:8])
  sum(b * 256^(0:7))
}

# --- the compound file -------------------------------------------------------
ole_open <- function(path) {
  d <- readBin(path, "raw", n = file.info(path)$size)
  sig <- c(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1)
  if (length(d) < 512 || any(as.integer(d[1:8]) != sig)) {
    stop("not an OLE compound file", call. = FALSE)
  }

  ssz <- bitwShiftL(1L, u16(d, 0x1e))
  msz <- bitwShiftL(1L, u16(d, 0x20))
  n_fat <- u32(d, 0x2c)
  dir_start <- u32(d, 0x30)
  cutoff <- u32(d, 0x38)
  mfat_start <- u32(d, 0x3c)
  n_mfat <- u32(d, 0x40)
  difat_start <- u32(d, 0x44)
  n_difat <- u32(d, 0x48)

  sect <- function(n) {
    off <- (n + 1) * ssz
    if (off + ssz > length(d)) return(raw(ssz))
    d[(off + 1):(off + ssz)]
  }

  difat <- u32v(d[(0x4c + 1):(0x4c + 109 * 4)])
  nxt <- difat_start
  left <- n_difat
  while (!(nxt %in% FREESECT) && left > 0) {
    s <- sect(nxt)
    difat <- c(difat, u32v(s[1:(ssz - 4)]))
    nxt <- u32(s, ssz - 4)
    left <- left - 1
  }

  fat <- numeric(0)
  for (fs in difat[seq_len(min(n_fat, length(difat)))]) {
    if (fs %in% FREESECT) next
    fat <- c(fat, u32v(sect(fs)))
  }

  chain <- function(start, table) {
    out <- numeric(0)
    cur <- start
    guard <- 0
    while (!(cur %in% FREESECT) && cur < length(table) && guard < 1e6) {
      out <- c(out, cur)
      cur <- table[cur + 1]
      guard <- guard + 1
    }
    out
  }

  cat_sectors <- function(secs) {
    if (!length(secs)) return(raw(0))
    do.call(c, lapply(secs, sect))
  }

  dirbytes <- cat_sectors(chain(dir_start, fat))
  entries <- list()
  n_ent <- length(dirbytes) %/% 128
  for (k in seq_len(n_ent)) {
    off <- (k - 1) * 128
    nlen <- u16(dirbytes, off + 64)
    typ <- as.integer(dirbytes[off + 67])
    if (!(typ %in% c(1L, 2L, 5L)) || nlen < 2 || nlen > 64) next
    nm_raw <- dirbytes[(off + 1):(off + nlen - 2)]
    nm <- intToUtf8(u16v_chars(nm_raw))
    entries[[length(entries) + 1L]] <- list(
      name = nm, type = typ,
      start = u32(dirbytes, off + 116),
      size = u64(dirbytes, off + 120)
    )
  }

  root <- Filter(function(e) e$type == 5L, entries)
  mini <- if (length(root)) cat_sectors(chain(root[[1]]$start, fat)) else raw(0)

  mfat <- numeric(0)
  if (n_mfat > 0) {
    for (s in chain(mfat_start, fat)) mfat <- c(mfat, u32v(sect(s)))
  }

  stream <- function(want) {
    for (e in entries) {
      if (e$type == 2L && identical(e$name, want)) {
        if (e$size < cutoff) {
          idx <- chain(e$start, mfat)
          if (!length(idx)) return(raw(0))
          parts <- lapply(idx, function(i) {
            a <- i * msz + 1
            b <- min((i + 1) * msz, length(mini))
            if (a > length(mini)) raw(0) else mini[a:b]
          })
          raw_all <- do.call(c, parts)
        } else {
          raw_all <- cat_sectors(chain(e$start, fat))
        }
        n <- min(e$size, length(raw_all))
        if (n <= 0) return(raw(0))
        return(raw_all[1:n])
      }
    }
    NULL
  }

  list(stream = stream, entries = entries)
}

# UTF-16LE raw -> integer code points (surrogates are not expected in a name)
u16v_chars <- function(r) {
  b <- as.integer(r)
  n <- length(b) %/% 2L
  if (n == 0L) return(integer(0))
  i <- seq_len(n) * 2L - 1L
  b[i] + b[i + 1L] * 256
}

# --- MS-OVBA decompression (MS-OVBA 2.4.1) -----------------------------------
# Returns a raw vector, or NULL when the container does not decode.
ovba_decompress <- function(data, start) {
  n <- length(data)
  if (start >= n || as.integer(data[start + 1]) != 1L) return(NULL)

  cap <- 65536L
  out <- raw(cap)
  len <- 0L
  ensure <- function(need) {
    while (len + need > cap) {
      cap <<- cap * 2L
      tmp <- raw(cap)
      if (len > 0L) tmp[1:len] <- out[1:len]
      out <<- tmp
    }
  }

  pos <- start + 1 # 0 based index of the next byte to read
  while (pos + 1 < n) {
    header <- u16(data, pos)
    pos <- pos + 2
    size <- bitwAnd(header, 0x0FFF) + 3
    compressed <- bitwAnd(header, 0x8000) != 0
    endp <- pos + size - 2
    if (endp > n) return(NULL)
    chunk_start <- len

    if (!compressed) {
      take <- min(4096, n - pos)
      ensure(take)
      out[(len + 1):(len + take)] <- data[(pos + 1):(pos + take)]
      len <- len + take
      pos <- endp
      next
    }

    while (pos < endp) {
      flag <- as.integer(data[pos + 1])
      pos <- pos + 1
      if (flag == 0L && pos + 8 <= endp) {
        # eight literals in a row, the common case: copy them in one go
        ensure(8L)
        out[(len + 1):(len + 8)] <- data[(pos + 1):(pos + 8)]
        len <- len + 8L
        pos <- pos + 8
        next
      }
      for (i in 0:7) {
        if (pos >= endp) break
        if (bitwAnd(bitwShiftR(flag, i), 1L) == 0L) {
          ensure(1L)
          out[len + 1] <- data[pos + 1]
          len <- len + 1L
          pos <- pos + 1
        } else {
          token <- u16(data, pos)
          pos <- pos + 2
          diff <- len - chunk_start
          if (diff < 1) return(NULL)
          bits <- max(4, ceiling(log2(diff)))
          clen <- bitwAnd(token, bitwShiftR(0xFFFF, bits)) + 3
          coff <- bitwShiftR(token, 16 - bits) + 1
          src <- len - coff
          if (src < 0) return(NULL)
          ensure(clen)
          if (coff >= clen) {
            out[(len + 1):(len + clen)] <- out[(src + 1):(src + clen)]
          } else {
            # overlapping copy: it must be byte by byte
            for (k in seq_len(clen)) out[len + k] <- out[src + k]
          }
          len <- len + clen
        }
      }
    }
    pos <- endp
  }
  if (len == 0L) return(NULL)
  out[1:len]
}

# --- the dir stream, for the source offset of each module --------------------
# Records are id (u16), size (u32), payload. PROJECTVERSION is the one record
# whose size field lies: it says 4 and carries 6.
module_offsets <- function(dir_raw) {
  res <- list()
  pos <- 0
  n <- length(dir_raw)
  cur_stream <- NULL
  while (pos + 6 <= n) {
    id <- u16(dir_raw, pos)
    size <- u32(dir_raw, pos + 2)
    pos <- pos + 6
    if (id == 0x0009) {
      pos <- pos + 6
      next
    }
    if (size < 0 || pos + size > n) break
    if (id == 0x001A) {
      cur_stream <- rawToChar(dir_raw[(pos + 1):(pos + size)])
      Encoding(cur_stream) <- "latin1"
    } else if (id == 0x0031 && !is.null(cur_stream)) {
      res[[cur_stream]] <- u32(dir_raw, pos)
    } else if (id == 0x002B) {
      cur_stream <- NULL
    }
    pos <- pos + size
  }
  res
}

# --- source recovery ---------------------------------------------------------
source_of <- function(stream_raw, offset) {
  if (!is.null(offset) && offset < length(stream_raw)) {
    got <- ovba_decompress(stream_raw, offset)
    if (!is.null(got)) {
      txt <- rawToChar(got)
      Encoding(txt) <- "latin1"
      if (grepl("Attribute VB_Name", txt, fixed = TRUE)) return(txt)
    }
  }
  # fallback: try every compressed container signature in the stream
  idx <- which(as.integer(stream_raw) == 1L) - 1L
  for (i in idx) {
    got <- ovba_decompress(stream_raw, i)
    if (is.null(got)) next
    txt <- rawToChar(got)
    Encoding(txt) <- "latin1"
    if (grepl("Attribute VB_Name", txt, fixed = TRUE)) return(txt)
  }
  NULL
}

# Cut both sides down to comparable code.
#
# Two things above the code are written by the exporter and are not part of the
# module as it is stored in the workbook:
#   * the VERSION / BEGIN / END block at the top of an exported .cls
#   * the module level Attribute VB_* block. The stored module carries
#     VB_Base, VB_TemplateDerived and VB_Customizable; the exported file does
#     not. Leaving them in shifts every following line and makes an identical
#     module read as though every one of its lines had changed.
#
# Only the CONTIGUOUS attribute block at the top is dropped, so an Attribute
# line inside a procedure body still counts as code on both sides.
normalise <- function(txt) {
  txt <- gsub("\r\n", "\n", txt, fixed = TRUE)
  txt <- gsub("\r", "\n", txt, fixed = TRUE)
  at <- regexpr("Attribute VB_Name", txt, fixed = TRUE)
  if (at > 1) txt <- substring(txt, at)
  lines <- strsplit(txt, "\n", fixed = TRUE)[[1]]
  lines <- sub("[ \t]+$", "", lines)

  i <- 1L
  while (i <= length(lines) && grepl("^Attribute VB_", lines[i])) i <- i + 1L
  lines <- if (i > length(lines)) character(0) else lines[i:length(lines)]

  while (length(lines) && !nzchar(lines[1])) lines <- lines[-1]
  while (length(lines) && !nzchar(lines[length(lines)])) {
    lines <- lines[-length(lines)]
  }
  lines
}

# The form a line takes once the VBE's own rewrites are undone.
#
# Opening a project in the VBE rewrites the code in ways that are not edits:
#   * the capitalisation of an identifier is made to match its declaration
#     across every module, so `keysTable` becomes `KeysTable` everywhere
#   * whitespace is re-laid around tokens: `Then  Exit` comes back as
#     `Then Exit`, `SubAddress:= x` as `SubAddress:=x`, `If( a` as `If (a`
#
# So a line is compared with its whitespace removed and its case folded --
# but ONLY outside string literals. What is inside the quotes is kept exactly,
# because "0 - 5 months" changing to "0 - 5 Months" is a real edit and this
# repo's whole job is translated strings.
#
# Splitting on the quote character alternates outside/inside, and VBA's escaped
# quote ("") falls out as an empty inside segment, which needs no special case.
#
# The one VBE rewrite this does NOT undo is numeric literals: `1E-13` is stored
# for `0.0000000000001`, and that still reports as a difference.

# The VBE also rewrites numeric literals into its own preferred spelling, so
# `0.0000000000001` is stored as `1E-13`. Every number is put through as.numeric
# and written back the same way, which makes the two spellings one. Digits that
# are part of an identifier (Sheet1) get the same treatment on both sides, so
# they still compare equal to themselves and unequal to anything else.
canon_nums <- function(s) {
  m <- gregexpr("[0-9]+\\.?[0-9]*(e[-+]?[0-9]+)?|\\.[0-9]+(e[-+]?[0-9]+)?",
                s, perl = TRUE)
  regmatches(s, m) <- lapply(regmatches(s, m), function(v) {
    if (!length(v)) return(v)
    n <- suppressWarnings(as.numeric(v))
    ifelse(is.na(n), v, formatC(n, format = "e", digits = 10))
  })
  s
}

canon_line <- function(s) {
  parts <- strsplit(s, "\"", fixed = TRUE)[[1]]
  if (!length(parts)) return("")
  outside <- seq_along(parts) %% 2L == 1L
  parts[outside] <- gsub("[ \t]", "", canon_nums(tolower(parts[outside])))
  paste(parts, collapse = "\"")
}
canon <- function(x) {
  if (!length(x)) return(character(0))
  vapply(x, canon_line, character(1), USE.NAMES = FALSE)
}

find_source <- function(name) {
  hits <- character(0)
  for (root in roots) {
    if (!dir.exists(root)) next
    # A component name is letters, digits and underscores only, so it carries
    # nothing that needs escaping for a regular expression.
    pat <- paste0("^", name, "\\.(cls|bas|frm)$")
    hits <- c(hits, list.files(root, pattern = pat, recursive = TRUE,
                               full.names = TRUE))
  }
  sort(sub("^\\./", "", hits))
}

read_lines_any <- function(path) {
  raw_b <- readBin(path, "raw", n = file.info(path)$size)
  txt <- rawToChar(raw_b)
  Encoding(txt) <- "latin1"
  normalise(txt)
}

# --- run ---------------------------------------------------------------------
ole <- ole_open(bin_path)

project <- ole$stream("PROJECT")
if (is.null(project)) {
  stop("no PROJECT stream in this vbaProject.bin", call. = FALSE)
}
ptxt <- rawToChar(project)
Encoding(ptxt) <- "latin1"

m <- regmatches(ptxt, gregexpr("(?m)^(Document|Module|Class|BaseClass)=[A-Za-z_][A-Za-z0-9_]{0,39}",
                               ptxt, perl = TRUE))[[1]]
kinds <- sub("=.*$", "", m)
names_ <- sub("^[^=]+=", "", m)

keep <- !duplicated(names_)
kinds <- kinds[keep]
names_ <- names_[keep]

offsets <- list()
if (mode == "content") {
  dir_raw <- ole$stream("dir")
  if (!is.null(dir_raw)) {
    dd <- ovba_decompress(dir_raw, 0)
    if (!is.null(dd)) offsets <- module_offsets(dd)
  }
}

out <- character(0)
row <- function(...) out <<- c(out, paste(c(...), collapse = "\t"))

for (i in seq_along(names_)) {
  kind <- kinds[i]
  name <- names_[i]

  if (mode == "list" || kind == "Document") {
    row(kind, name, "", "", "")
    next
  }

  hits <- find_source(name)
  if (!length(hits)) {
    row(kind, name, "-", "nosource", "")
    next
  }
  path <- hits[1]

  sraw <- ole$stream(name)
  if (is.null(sraw) || !length(sraw)) {
    row(kind, name, path, "noextract", "no module stream")
    next
  }
  src <- source_of(sraw, offsets[[name]])
  if (is.null(src)) {
    row(kind, name, path, "noextract", "source not decodable")
    next
  }

  a <- normalise(src)
  b <- read_lines_any(path)

  # The forms in these mocks are shells: the module stream carries the
  # Attribute block and nothing else, and the code lives only in the .frm on
  # disk. That is how they are built, so it is not drift and re-importing is
  # not the answer. Reported on its own rather than as a difference.
  if (!length(a) && length(b)) {
    row(kind, name, path, "empty",
        sprintf("no code in the workbook (repo file has %d lines)", length(b)))
    next
  }

  if (identical(a, b)) {
    row(kind, name, path, "same", if (kind == "BaseClass") "code only" else "")
  } else if (identical(canon(a), canon(b))) {
    k <- min(length(a), length(b))
    n_cos <- sum(a[seq_len(k)] != b[seq_len(k)])
    row(kind, name, path, "cosmetic",
        sprintf("%d line(s) differ only in case or spacing", n_cos))
  } else {
    # Reported on the canonical comparison, so the line pointed at is a real
    # difference and not one of the VBE's rewrites.
    la <- canon(a)
    lb <- canon(b)
    k <- min(length(la), length(lb))
    first <- k + 1L
    if (k > 0) {
      d <- which(la[seq_len(k)] != lb[seq_len(k)])
      if (length(d)) first <- d[1]
    }
    row(kind, name, path, "differs",
        sprintf("first differs at line %d (workbook %d lines, repo %d)%s",
                first, length(a), length(b),
                if (kind == "BaseClass") ", code only" else ""))
  }
}

cat(out, sep = "\n")
if (length(out)) cat("\n")
