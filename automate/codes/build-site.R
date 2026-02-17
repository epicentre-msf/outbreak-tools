library(commonmark)
library(fs)
library(stringr)
library(glue)
library(here)

# Minimal HTML site generator for VBA documentation markdown files.
# Reads .md files produced by class-doc.R, converts to HTML, and
# writes a static site to src/docs/site/.

docs_dir   <- here("src", "docs")
site_dir   <- path(docs_dir, "site")
style_src  <- here("automate", "codes", "style.css")

dir_create(site_dir)

# --- Post-processing helpers ------------------------------------------------

slugify <- function(text) {
  slug <- tolower(trimws(text))
  slug <- gsub("[^a-z0-9]+", "-", slug)
  slug <- gsub("-+", "-", slug)
  slug <- gsub("^-|-$", "", slug)
  slug
}

# Post-process converted HTML to:
#   1. Wrap top-level <h2> section groups in collapsible <details> elements.
#   2. Add id attributes and anchor links to <h3> method headers.
#   3. Add id attributes to headers inside existing <details> blocks.
enhance_html <- function(html) {
  lines <- strsplit(html, "\n", fixed = TRUE)[[1]]
  result <- character()
  details_depth <- 0L
  section_open <- FALSE

  for (line in lines) {
    opens  <- str_count(line, "<details")
    closes <- str_count(line, "</details>")

    # Close any open doc-section before entering an original <details> block
    if (opens > 0L && details_depth == 0L && section_open) {
      result <- c(result, "</div>", "</details>")
      section_open <- FALSE
    }

    was_top <- (details_depth == 0L)
    details_depth <- details_depth + opens

    if (was_top && opens == 0L) {
      # Top-level line (not entering a <details>)
      h2 <- str_match(line, "^<h2>([^<]+)</h2>$")
      h3 <- str_match(line, "^<h3>([^<]+)</h3>$")

      if (!is.na(h2[1])) {
        if (section_open) {
          result <- c(result, "</div>", "</details>")
        }
        id <- slugify(h2[2])
        result <- c(result,
          glue('<details open class="doc-section" id="{id}">'),
          glue('<summary>{h2[2]}</summary>'),
          '<div class="section-body">'
        )
        section_open <- TRUE
      } else if (!is.na(h3[1])) {
        id <- slugify(h3[2])
        result <- c(result,
          glue('<h3 id="{id}">{h3[2]} <a class="anchor-link" href="#{id}">#</a></h3>')
        )
      } else {
        result <- c(result, line)
      }

    } else if (was_top && opens > 0L) {
      # A <details> tag at top level -- pass through unchanged
      result <- c(result, line)

    } else {
      # Inside an existing <details> block -- add IDs only, no wrapping
      h2 <- str_match(line, "^<h2>([^<]+)</h2>$")
      h3 <- str_match(line, "^<h3>([^<]+)</h3>$")

      if (!is.na(h2[1])) {
        id <- slugify(paste0("int-", h2[2]))
        result <- c(result, glue('<h2 id="{id}">{h2[2]}</h2>'))
      } else if (!is.na(h3[1])) {
        id <- slugify(paste0("int-", h3[2]))
        result <- c(result,
          glue('<h3 id="{id}">{h3[2]} <a class="anchor-link" href="#{id}">#</a></h3>')
        )
      } else {
        result <- c(result, line)
      }
    }

    details_depth <- details_depth - closes
    if (details_depth < 0L) details_depth <- 0L
  }

  # Close trailing doc-section if still open
  if (section_open) {
    result <- c(result, "</div>", "</details>")
  }

  paste(result, collapse = "\n")
}

# --- Page template ----------------------------------------------------------

page_template <- function(title, body_html, css_file = "style.css") {
  paste0(
    '<!doctype html>\n',
    '<html lang="en">\n',
    '<head>\n',
    '  <meta charset="utf-8">\n',
    '  <meta name="viewport" content="width=device-width, initial-scale=1">\n',
    '  <title>', title, ' - OBT Dev Docs</title>\n',
    '  <link rel="stylesheet" href="', css_file, '">\n',
    '  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/github.min.css">\n',
    '</head>\n',
    '<body>\n',
    '  <nav><a href="index.html">&larr; Index</a></nav>\n',
    '  <main>\n',
    body_html, '\n',
    '  </main>\n',
    '  <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/highlight.min.js"></script>\n',
    '  <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/languages/vbnet.min.js"></script>\n',
    '  <script>hljs.highlightAll();</script>\n',
    '</body>\n',
    '</html>\n'
  )
}

# --- Read and convert each .md file -----------------------------------------

md_files <- dir_ls(docs_dir, regexp = "\\.md$", type = "file")
if (length(md_files) == 0) {
  cli::cli_abort("No .md files found in {docs_dir}. Run create-docs.R first.")
}

# Build a mapping: class_name -> folder (extracted from <!-- folder: X -->)
class_info <- data.frame(
  name   = character(),
  folder = character(),
  file   = character(),
  stringsAsFactors = FALSE
)

for (md_file in md_files) {
  class_name <- path_ext_remove(path_file(md_file))
  md_text <- paste(readLines(md_file, warn = FALSE), collapse = "\n")

  # Extract folder from HTML comment
  folder_match <- str_match(md_text, "<!--\\s*folder:\\s*([^>]+?)\\s*-->")
  folder <- if (!is.na(folder_match[1])) str_trim(folder_match[2]) else "Other"
  if (!nzchar(folder)) folder <- "Other"

  # Strip the folder comment before converting
  md_clean <- str_remove(md_text, "<!--\\s*folder:[^>]*-->\\s*")

  # Convert markdown to HTML and enhance with anchors + collapsible sections
  html_body <- markdown_html(md_clean, extensions = TRUE)
  html_body <- enhance_html(html_body)

  # Write the page
  out_path <- path(site_dir, paste0(class_name, ".html"))
  writeLines(page_template(class_name, html_body), out_path)

  class_info <- rbind(class_info, data.frame(
    name = class_name, folder = folder, file = basename(out_path),
    stringsAsFactors = FALSE
  ))
}

cli::cli_inform(glue("Converted {nrow(class_info)} pages to HTML."))

# --- Build index.html -------------------------------------------------------

# Sort folders alphabetically, classes alphabetically within each folder
class_info <- class_info[order(class_info$folder, tolower(class_info$name)), ]
folders <- unique(class_info$folder)

index_body <- c(
  '<h1>OutbreakTools &mdash; Developer Reference</h1>',
  glue('<p>{nrow(class_info)} documented classes across {length(folders)} folders.</p>'),
  ''
)

for (folder_name in folders) {
  subset <- class_info[class_info$folder == folder_name, ]
  items <- paste0(
    '    <li><a href="', subset$file, '">', subset$name, '</a></li>'
  )
  index_body <- c(
    index_body,
    glue('<h2>{folder_name}</h2>'),
    '<ul>',
    items,
    '</ul>',
    ''
  )
}

index_html <- page_template(
  "Index",
  paste(index_body, collapse = "\n")
)
writeLines(index_html, path(site_dir, "index.html"))
cli::cli_inform("Written: index.html")

# --- Copy style.css ---------------------------------------------------------

if (file_exists(style_src)) {
  file_copy(style_src, path(site_dir, "style.css"), overwrite = TRUE)
  cli::cli_inform("Copied: style.css")
} else {
  cli::cli_alert_warning(glue("style.css not found at {style_src}"))
}

cli::cli_inform(glue("Site ready at: {site_dir}/index.html"))
