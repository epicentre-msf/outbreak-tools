library(R6)
library(stringr)
library(glue)
library(fs)
library(readr)
library(here)

VBADocParser <- R6Class(
  "VBADocParser",
  public = list(
    initialize = function(
      folder,
      output_folder = here("src", "docs"),
      proj_path = here()
    ) {
      private$folder <- path_abs(folder)
      private$output_folder <- path_abs(output_folder)
      private$proj_path <- proj_path
      dir_create(private$output_folder)
      cli::cli_inform(glue("Parser initialized for folder: {folder}"))
      cli::cli_inform(glue("Parser will write in: {output_folder}"))
    },

    init_folder = function() private$folder,
    write_folder = function() private$output_folder,
    doc_classes = function() private$class_names,

    parse = function(exclude_files = NULL) {
      cls_files <- dir_ls(
        self$init_folder(),
        recurse = TRUE,
        regexp = "\\.cls$"
      )

      cls_files <- setdiff(cls_files, exclude_files)

      private$class_names <- path_ext_remove(path_file(cls_files))

      for (i in seq_along(cls_files)) {
        file <- cls_files[i]
        class_name <- private$class_names[i]
        cli::cli_inform(glue("Parsing: {path_file(file)}"))

        lines <- read_lines(file)
        doc_info <- private$extract_doc_blocks(lines)

        if (length(doc_info$internals) > 0 || length(doc_info$externals) > 0) {
          md <- private$build_markdown(
            class_name,
            doc_info$externals,
            doc_info$internals,
            doc_info$tocs,
            private$class_names
          )
          write_file(md, path(self$write_folder(), glue("{class_name}.md")))
          cli::cli_inform(glue("Written: {class_name}.md"))
        } else {
          cli::cli_alert_warning(glue(
            "No valid documentation blocks found in {class_name}"
          ))
        }
      }
    },

    parse_class = function(class_name, exclude_files = NULL) {
      cls_files <- dir_ls(
        self$init_folder(),
        recurse = TRUE,
        regexp = "\\.cls$"
      )
      cls_files <- setdiff(cls_files, exclude_files)

      # Map class names to files
      class_map <- setNames(
        cls_files,
        path_ext_remove(path_file(cls_files))
      )

      if (!(class_name %in% names(class_map))) {
        cli::cli_abort(glue("Class '{class_name}' not found under {self$init_folder()}"))
      }

      target_file <- class_map[[class_name]]
      private$class_names <- names(class_map)

      cli::cli_inform(glue("Parsing single class: {class_name}"))
      lines <- read_lines(target_file)
      doc_info <- private$extract_doc_blocks(lines)

      if (length(doc_info$internals) > 0 || length(doc_info$externals) > 0) {
        md <- private$build_markdown(
          class_name,
          doc_info$externals,
          doc_info$internals,
          doc_info$tocs,
          private$class_names,
          interface_mode = FALSE
        )
        write_file(md, path(self$write_folder(), glue("{class_name}.md")))
        cli::cli_inform(glue("Written: {class_name}.md"))
      } else {
        cli::cli_alert_warning(glue(
          "No valid documentation blocks found in {class_name}"
        ))
      }
    },

    parse_interface_for_class = function(class_name, exclude_files = NULL) {
      # List all .cls files in scope
      cls_files <- dir_ls(
        self$init_folder(),
        recurse = TRUE,
        regexp = "\\.cls$"
      )
      cls_files <- setdiff(cls_files, exclude_files)

      # Build maps
      all_names <- path_ext_remove(path_file(cls_files))
      file_by_name <- setNames(cls_files, all_names)
      private$class_names <- all_names

      # Preferred interface naming convention
      preferred_iface <- paste0("I", class_name)
      iface_name <- NULL
      if (preferred_iface %in% names(file_by_name)) {
        iface_name <- preferred_iface
      } else if (class_name %in% names(file_by_name)) {
        # Fallback: read Implements lines from the class file
        lines <- read_lines(file_by_name[[class_name]])
        impl <- stringr::str_match(lines, "^\\s*Implements\\s+([A-Za-z0-9_]+)")
        impl <- impl[!is.na(impl[, 2]), 2]
        if (length(impl) > 0) {
          # Pick the first declared interface available in files
          impl <- impl[impl %in% names(file_by_name)]
          if (length(impl) > 0) iface_name <- impl[[1]]
        }
      }

      if (is.null(iface_name)) {
        cli::cli_abort(glue("No interface found for class '{class_name}'. Expected '{preferred_iface}' or an Implements declaration."))
      }

      iface_file <- file_by_name[[iface_name]]
      cli::cli_inform(glue("Parsing interface for class {class_name}: {iface_name}"))

      lines <- read_lines(iface_file)
      doc_info <- private$extract_doc_blocks(lines)

      # Treat all documented members as external API for interfaces
      docs_external <- c(doc_info$externals, doc_info$internals)
      docs_internal <- list()

      if (length(docs_external) > 0) {
        md <- private$build_markdown(
          iface_name,
          docs_external,
          docs_internal,
          doc_info$tocs,
          private$class_names,
          interface_mode = TRUE
        )
        write_file(md, path(self$write_folder(), glue("{iface_name}.md")))
        cli::cli_inform(glue("Written: {iface_name}.md"))
      } else {
        cli::cli_alert_warning(glue(
          "No valid documentation blocks found in interface {iface_name}"
        ))
      }
    },

    detect_usages = function(exclude_files = NULL) {
      all_files <- dir_ls(
        self$init_folder(),
        recurse = TRUE,
        regexp = "\\.(cls|bas)$"
      )
      all_files <- setdiff(all_files, exclude_files)
      usage_map <- setNames(
        vector("list", length(private$class_names)),
        private$class_names
      )

      for (file in all_files) {
        content <- read_file(file)
        for (class_name in private$class_names) {
          iface <- paste0("I", class_name)
          if (str_detect(content, fixed(iface))) {
            usage_map[[class_name]] <- unique(c(
              usage_map[[class_name]],
              path_rel(file, private$proj_path)
            ))
          }
        }
      }

      for (class_name in names(usage_map)) {
        md_path <- path(self$write_folder(), glue("{class_name}.md"))
        if (file_exists(md_path)) {
          md <- read_lines(md_path)
          if (length(usage_map[[class_name]]) > 0) {
            usage_files <- sort(usage_map[[class_name]])
            section <- c(
              "",
              "<details>",
              glue(
                "<summary>Used in ({length(usage_files)} file(s))</summary>"
              ),
              "",
              glue("- [{basename(f)}]({f})", f = usage_files),
              "",
              "</details>"
            )
            md <- c(md, section)
            write_lines(md, md_path)
            cli::cli_inform(glue("Updated usages for: {class_name}"))
          }
        }
      }
    },

    extract_enums = function(exclude_files = NULL) {
      files <- dir_ls(
        self$init_folder(),
        recurse = TRUE,
        regexp = "\\.(cls|bas)$"
      )

      files <- setdiff(files, exclude_files)
      enum_blocks <- list()

      for (file in files) {
        lines <- read_lines(file)
        i <- 1L
        n <- length(lines)

        while (i <= n) {
          if (str_detect(lines[i], "^\\	?\\s*Public\\s+Enum\\s+")) {
            enum_name <- str_match(lines[i], "Public\\s+Enum\\s+(\\w+)")[, 2]
            members <- c()
            i <- i + 1L
            while (i <= n && !str_detect(lines[i], "^\\s*End\\s+Enum")) {
              line <- str_trim(lines[i])
              if (line != "" && !str_starts(line, "'")) {
                members <- c(members, line)
              }
              i <- i + 1L
            }
            enum_blocks[[enum_name]] <- members
          }
          i <- i + 1L
        }
      }

      out <- c("# Enumerations", "")
      for (enum in names(enum_blocks)) {
        out <- c(out, glue("## {enum}"), "")
        out <- c(out, paste0("- `", enum_blocks[[enum]], "`"), "")
      }

      write_lines(out, path(self$write_folder(), "Enumerations.md"))
      cli::cli_inform("Written: Enumerations.md")
    },

    build_master_markdown = function(title = "Code Documentation") {
      out_dir <- self$write_folder()
      md_files <- dir_ls(out_dir, regexp = "\\.md$", type = "file")
      md_files <- md_files[!basename(md_files) %in% c("CodeDocumentation.md")]
      md_files <- md_files[order(tolower(basename(md_files)))]

      header <- c(
        "---",
        glue("title: {title}"),
        "format:",
        "  html:",
        "    toc: true",
        "    theme: 'cosmo'",
        "---",
        "",
        "# Class Index",
        ""
      )

      index <- glue("- [{path_ext_remove(basename(f))}]({basename(f)})", f = md_files)

      sections <- c()
      for (f in md_files) {
        sections <- c(
          sections,
          "",
          glue("\n---\n# {path_ext_remove(basename(f))}\n"),
          read_lines(f)
        )
      }

      out <- c(header, index, sections)
      write_lines(out, path(out_dir, "CodeDocumentation.md"))
      cli::cli_inform("Written: CodeDocumentation.md")
    },

    build_site_index = function(title = "Code Documentation") {
      out_dir <- self$write_folder()
      md_files <- dir_ls(out_dir, regexp = "\\.md$", type = "file")
      md_files <- md_files[order(tolower(basename(md_files)))]

      items <- paste0(
        '<li><a href="', basename(md_files), '">',
        path_ext_remove(basename(md_files)),
        '</a></li>'
      )

      html <- c(
        '<!doctype html>',
        '<html lang="en">',
        '<head>',
        '  <meta charset="utf-8"/>',
        glue('  <title>{title}</title>'),
        '  <meta name="viewport" content="width=device-width, initial-scale=1"/>',
        '  <style>body{font-family:system-ui,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:2rem;max-width:960px} h1{margin-top:0} ul{line-height:1.8} .muted{color:#666;font-size:0.9em} .card{border:1px solid #eee;border-radius:8px;padding:1rem;margin:1rem 0} .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:1rem} </style>',
        '</head>',
        '<body>',
        glue('<h1>{title}</h1>'),
        '<p class="muted">Static index linking to per-class Markdown docs.</p>',
        '<div class="card">',
        '<h2>Classes</h2>',
        '<ul>',
        items,
        '</ul>',
        '</div>',
        '<div class="card">',
        '<h2>Combined Documentation</h2>',
        '<ul>',
        '<li><a href="CodeDocumentation.md">CodeDocumentation.md</a></li>',
        '</ul>',
        '</div>',
        '</body>',
        '</html>'
      )

      write_lines(html, path(out_dir, "index.html"))
      cli::cli_inform("Written: index.html")

      # Also emit a very simple HTML wrapper for the combined markdown
      combined_md <- path(out_dir, "CodeDocumentation.md")
      if (file_exists(combined_md)) {
        md_text <- paste(read_lines(combined_md), collapse = "\n")
        # Render markdown as preformatted text for offline viewing
        book_html <- c(
          '<!doctype html>',
          '<html lang="en">',
          '<head>',
          '  <meta charset="utf-8"/>',
          glue('  <title>{title}</title>'),
          '  <meta name="viewport" content="width=device-width, initial-scale=1"/>',
          '  <style>body{font-family:system-ui,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:2rem;max-width:960px} pre{white-space:pre-wrap;word-wrap:break-word} .muted{color:#666}</style>',
          '</head>',
          '<body>',
          glue('<h1>{title}</h1>'),
          '<p class="muted">Combined markdown rendered as preformatted text. For rich rendering, open the .md in a markdown viewer or Quarto.</p>',
          '<pre>',
          md_text,
          '</pre>',
          '</body>',
          '</html>'
        )
        write_lines(book_html, path(out_dir, "CodeDocumentation.html"))
        cli::cli_inform("Written: CodeDocumentation.html")
      }
    }
  ),

  private = list(
    folder = NULL,
    output_folder = NULL,
    proj_path = NULL,
    class_names = NULL,

    extract_doc_blocks = function(lines) {
      externals <- list()
      internals <- list()
      tocs <- list()
      n <- length(lines)
      current_section <- NULL
      doc <- NULL
      last_tag <- NULL
      last_param_index <- NA_integer_
      last_note_index <- NA_integer_
      last_remark_index <- NA_integer_
      last_throw_index <- NA_integer_
      last_depend_index <- NA_integer_
      i <- 1L

      reset_tracking <- function() {
        last_tag <<- NULL
        last_param_index <<- NA_integer_
        last_note_index <<- NA_integer_
        last_remark_index <<- NA_integer_
        last_throw_index <<- NA_integer_
        last_depend_index <<- NA_integer_
      }
      reset_tracking()

      finalize_doc <- function(entry, signature_lines) {
        if (length(signature_lines) == 0) {
          return()
        }

        entry$signature <- paste(signature_lines, collapse = "\n")
        if (!private$has_text(entry$label)) {
          inferred <- private$infer_label_from_signature(signature_lines[[1]])
          if (private$has_text(inferred)) {
            entry$label <- inferred
          } else {
            entry$label <- str_trim(signature_lines[[1]])
          }
        }

        if (!private$has_text(entry$section)) {
          entry$section <- current_section
        }

        target <- if (isTRUE(entry$export) && !isTRUE(entry$private)) "externals" else "internals"
        if (target == "externals") {
          externals <<- c(externals, list(entry))
        } else {
          internals <<- c(internals, list(entry))
        }
      }

      while (i <= n) {
        line <- lines[i]
        trimmed <- str_trim(line)

        if (str_detect(trimmed, "^'@")) {
          tag_match <- str_match(trimmed, "^'@([\\-A-Za-z0-9_]+):?\\s*(.*)$")
          if (is.na(tag_match[1])) {
            i <- i + 1L
            next
          }

          raw_tag <- tag_match[2]
          content <- str_trim(tag_match[3])
          tag <- str_to_lower(raw_tag)

          if (tag == "pram") tag <- "param"
          if (tag %in% c("params", "parameters")) tag <- "params"
          if (tag %in% c("returns", "returned")) tag <- "return"

          if (tag == "section") {
            current_section <- if (private$has_text(content)) content else NULL
            doc <- NULL
            reset_tracking()
            i <- i + 1L
            next
          }

          module_level <- c(
            "folder",
            "moduledescription",
            "interfacedescription",
            "interface",
            "ignoremodule",
            "ignore",
            "defaultmember",
            "predeclaredid"
          )
          if (tag %in% module_level) {
            i <- i + 1L
            next
          }

          if (tag == "jump" && is.null(doc)) {
            i <- i + 1L
            next
          }

          if (is.null(doc)) {
            doc <- list(section = current_section)
          }

          reset_tracking()

          if (tag == "label") {
            if (private$has_text(content)) {
              doc$label <- content
            }
          } else if (tag == "method") {
            if (private$has_text(content)) {
              doc$label <- content
            }
          } else if (tag %in% c("sub-title", "fun-title", "prop-title")) {
            doc[[tag]] <- content
          } else if (tag %in% c("details", "description")) {
            doc$details <- private$append_text(doc$details, content)
            last_tag <- tag
          } else if (tag == "note") {
            doc$notes <- c(doc$notes, content)
            last_tag <- "note"
            last_note_index <- length(doc$notes)
          } else if (tag == "remarks") {
            doc$remarks <- c(doc$remarks, content)
            last_tag <- "remarks"
            last_remark_index <- length(doc$remarks)
          } else if (tag == "throws") {
            doc$throws <- c(doc$throws, content)
            last_tag <- "throws"
            last_throw_index <- length(doc$throws)
          } else if (tag == "depends") {
            deps <- if (private$has_text(content)) str_split(content, "\\s*,\\s*")[[1]] else ""
            if (length(deps) == 0) deps <- ""
            doc$depends <- c(doc$depends, deps)
            last_tag <- "depends"
            last_depend_index <- length(doc$depends)
          } else if (tag == "export") {
            doc$export <- TRUE
          } else if (tag == "private") {
            doc$private <- TRUE
          } else if (tag == "param") {
            param_info <- private$parse_param_line(content)
            doc$params <- c(doc$params, list(param_info))
            last_tag <- "param"
            last_param_index <- length(doc$params)
          } else if (tag == "params") {
            last_tag <- "params"
            if (private$has_text(content)) {
              param_info <- private$parse_param_line(content)
              doc$params <- c(doc$params, list(param_info))
              last_param_index <- length(doc$params)
            }
          } else if (tag == "return") {
            doc$return <- private$parse_return_line(content)
            last_tag <- "return"
          } else if (tag == "jump") {
            doc$jump <- c(doc$jump, content)
          } else {
            if (str_detect(raw_tag, "^[A-Za-z0-9_]+$")) {
              param_info <- private$parse_param_line(content, raw_tag)
              doc$params <- c(doc$params, list(param_info))
              last_tag <- "param"
              last_param_index <- length(doc$params)
            }
          }

          i <- i + 1L
          next
        }

        if (str_detect(trimmed, "^'(?!@)")) {
          if (!is.null(doc)) {
            text <- str_trim(str_remove(trimmed, "^'"))
            if (identical(last_tag, "details") || identical(last_tag, "description")) {
              doc$details <- private$append_text(doc$details, text)
            } else if (identical(last_tag, "note") && !is.na(last_note_index)) {
              doc$notes[last_note_index] <- private$append_text(doc$notes[last_note_index], text)
            } else if (identical(last_tag, "remarks") && !is.na(last_remark_index)) {
              doc$remarks[last_remark_index] <- private$append_text(doc$remarks[last_remark_index], text)
            } else if (identical(last_tag, "throws") && !is.na(last_throw_index)) {
              doc$throws[last_throw_index] <- private$append_text(doc$throws[last_throw_index], text)
            } else if (identical(last_tag, "depends") && !is.na(last_depend_index)) {
              doc$depends[last_depend_index] <- private$append_text(doc$depends[last_depend_index], text)
            } else if (identical(last_tag, "return") && !is.null(doc$return)) {
              doc$return$details <- private$append_text(doc$return$details, text)
            } else if (identical(last_tag, "param") && !is.na(last_param_index)) {
              doc$params[[last_param_index]]$details <- private$append_text(doc$params[[last_param_index]]$details, text)
            } else if (identical(last_tag, "params")) {
              bullet <- str_remove(text, "^[-*]\\s*")
              if (!private$has_text(bullet)) {
                bullet <- text
              }
              if (private$has_text(bullet)) {
                param_info <- private$parse_param_line(bullet)
                doc$params <- c(doc$params, list(param_info))
                last_param_index <- length(doc$params)
              }
            }
          }
          i <- i + 1L
          next
        }

        if (!is.null(doc)) {
          if (str_detect(trimmed, "^(Public|Private|Friend|Global|Static)\\s+(Sub|Function|Property|Let|Set|Get)")) {
            signature_lines <- character()
            repeat {
              signature_lines <- c(signature_lines, lines[i])
              if (str_detect(lines[i], "\\)") || i >= n) {
                break
              }
              i <- i + 1L
            }
            finalize_doc(doc, signature_lines)
            doc <- NULL
            reset_tracking()
            i <- i + 1L
            next
          }

          if (!nzchar(trimmed)) {
            i <- i + 1L
            next
          }

          if (str_detect(trimmed, "^Attribute\\s")) {
            i <- i + 1L
            next
          }

          doc <- NULL
          reset_tracking()
          next
        }

        i <- i + 1L
      }

      list(externals = externals, internals = internals, tocs = tocs)
    },

    build_markdown = function(
      class_name,
      externals,
      internals,
      tocs,
      class_names,
      interface_mode = FALSE
    ) {
      output <- c(
        "---",
        glue("title: {class_name}"),
        "format:",
        "  html:",
        "    toc: true",
        "    theme: 'cosmo'",
        "---",
        ""
      )

      # Primary section
      external_md <- private$resolve_doc(externals, class_names)
      if (private$has_text(external_md)) {
        output <- c(output, external_md, "")
      }

      # Interface mode: do not emit internals callout
      if (!interface_mode) {
        internal_md <- private$resolve_doc(internals, class_names)
        if (private$has_text(internal_md)) {
          output <- c(
            output,
            "",
            "::: {.callout-note collapse=\"true\" title=\"Additional not exported Subs \"}",
            internal_md,
            ":::"
          )
        }
      }
      paste(output, collapse = "\n")
    },

    resolve_links = function(text, class_names) {
      str_replace_all(text, "see::([A-Za-z0-9_]+)", function(m) {
        cls <- str_match(m, "see::([A-Za-z0-9_]+)")[, 2]
        if (cls %in% class_names) glue("[{cls}]({cls}.md)") else cls
      })
    },

    resolve_doc = function(lst_doc, class_names) {
      if (length(lst_doc) == 0) {
        return("")
      }

      output <- character()
      current_section <- NULL

      for (doc in lst_doc) {
        section_title <- doc$section
        if (private$has_text(section_title) && !identical(section_title, current_section)) {
          current_section <- section_title
          output <- c(output, glue("## {section_title}"), "")
        }

        label <- if (private$has_text(doc$label)) {
          as.character(doc$label)[1]
        } else {
          "UnnamedMember"
        }

        summary <- NULL
        if (private$has_text(doc[["prop-title"]])) {
          summary <- doc[["prop-title"]]
        } else if (private$has_text(doc[["sub-title"]])) {
          summary <- doc[["sub-title"]]
        } else if (private$has_text(doc[["fun-title"]])) {
          summary <- doc[["fun-title"]]
        }

        anchor_source <- if (private$has_text(summary)) summary else label
        anchor <- private$slugify_anchor(anchor_source)

        header <- if (private$has_text(anchor)) {
          glue("### `{label}` {{#sec-{anchor}}}")
        } else {
          glue("### `{label}`")
        }

        block <- c("", header)

        if (private$has_text(summary)) {
          block <- c(block, "", glue("**{summary}**"))
        }

        if (private$has_text(doc$signature)) {
          block <- c(block, "", "**Signature:**", "```vb", doc$signature, "```")
        }

        if (private$has_text(doc$details)) {
          block <- c(block, "", private$resolve_links(doc$details, class_names))
        }

        if (!is.null(doc$params) && length(doc$params) > 0) {
          block <- c(block, "", "**Parameters:**")
          for (param in doc$params) {
            name <- if (private$has_text(param$name)) param$name else "param"
            type_suffix <- ""
            if (private$has_text(param$type)) {
              type_suffix <- glue(" ({param$type})")
            }
            details_text <- ""
            if (private$has_text(param$details)) {
              details_text <- glue(": {private$resolve_links(param$details, class_names)}")
            }
            block <- c(block, glue("  - `{name}`{type_suffix}{details_text}"))
          }
        }

        if (!is.null(doc$return) && (private$has_text(doc$return$type) || private$has_text(doc$return$details))) {
          return_type <- if (private$has_text(doc$return$type)) glue("{doc$return$type} – ") else ""
          return_details <- if (private$has_text(doc$return$details)) private$resolve_links(doc$return$details, class_names) else ""
          block <- c(block, "", glue("**Returns:** {return_type}{return_details}"))
        }

        if (!is.null(doc$notes) && length(doc$notes) > 0) {
          notes_block <- character()
          for (note in doc$notes) {
            if (private$has_text(note)) {
              notes_block <- c(notes_block, glue("  - {private$resolve_links(note, class_names)}"))
            }
          }
          if (length(notes_block) > 0) {
            block <- c(block, "", "**Notes:**", notes_block)
          }
        }

        if (!is.null(doc$remarks) && length(doc$remarks) > 0) {
          remarks_block <- character()
          for (remark in doc$remarks) {
            if (private$has_text(remark)) {
              remarks_block <- c(remarks_block, glue("  - {private$resolve_links(remark, class_names)}"))
            }
          }
          if (length(remarks_block) > 0) {
            block <- c(block, "", "**Remarks:**", remarks_block)
          }
        }

        if (!is.null(doc$throws) && length(doc$throws) > 0) {
          throws_block <- character()
          for (throw_entry in doc$throws) {
            if (private$has_text(throw_entry)) {
              throws_block <- c(throws_block, glue("  - {private$resolve_links(throw_entry, class_names)}"))
            }
          }
          if (length(throws_block) > 0) {
            block <- c(block, "", "**Throws:**", throws_block)
          }
        }

        if (!is.null(doc$depends) && length(doc$depends) > 0) {
          depends_block <- character()
          for (dep in doc$depends) {
            if (private$has_text(dep)) {
              depends_block <- c(depends_block, glue("  - {private$resolve_links(dep, class_names)}"))
            }
          }
          if (length(depends_block) > 0) {
            block <- c(block, "", "**Depends on:**", depends_block)
          }
        }

        block <- c(block, "", "---")
        output <- c(output, block)
      }

      paste(output, collapse = "\n")
    },

    has_text = function(value) {
      if (is.null(value) || length(value) == 0) {
        return(FALSE)
      }
      text <- str_trim(as.character(value))
      text <- text[!is.na(text)]
      if (length(text) == 0) {
        return(FALSE)
      }
      any(nzchar(text))
    },

    append_text = function(existing, addition) {
      if (is.null(addition) || length(addition) == 0) {
        return(existing)
      }
      addition <- as.character(addition)
      addition <- addition[!is.na(addition)]
      addition <- str_trim(addition)
      addition <- addition[nzchar(addition)]
      if (length(addition) == 0) {
        return(existing)
      }
      addition_text <- paste(addition, collapse = "\n")
      if (!private$has_text(existing)) {
        return(addition_text)
      }
      existing_text <- str_trim(paste(as.character(existing), collapse = "\n"))
      if (!nzchar(existing_text)) {
        return(addition_text)
      }
      paste(existing_text, addition_text, sep = "\n")
    },

    parse_param_line = function(text, fallback_name = NULL) {
      value <- ""
      if (!is.null(text) && length(text) > 0) {
        value <- str_trim(as.character(text)[1])
      }

      type <- NULL
      name <- fallback_name
      remainder <- value

      if (private$has_text(remainder)) {
        type_match <- str_match(remainder, "^\\{([^}]+)\\}\\s*(.*)$")
        if (!is.na(type_match[1])) {
          type <- str_trim(type_match[2])
          remainder <- str_trim(type_match[3])
        }

        if (!private$has_text(name)) {
          name_match <- str_match(remainder, "^([A-Za-z0-9_]+)\\s*(.*)$")
          if (!is.na(name_match[1])) {
            name <- name_match[2]
            remainder <- str_trim(name_match[3])
          }
        } else {
          remainder <- str_trim(remainder)
        }

        remainder <- str_remove(remainder, "^[-:\\s]+")
      }

      if (!private$has_text(name)) {
        name <- if (private$has_text(fallback_name)) fallback_name else "param"
      }

      list(
        name = name,
        type = type,
        details = remainder
      )
    },

    parse_return_line = function(text) {
      value <- ""
      if (!is.null(text) && length(text) > 0) {
        value <- str_trim(as.character(text)[1])
      }

      type <- NULL
      remainder <- value

      if (private$has_text(remainder)) {
        type_match <- str_match(remainder, "^\\{([^}]+)\\}\\s*(.*)$")
        if (!is.na(type_match[1])) {
          type <- str_trim(type_match[2])
          remainder <- str_trim(type_match[3])
        }
        remainder <- str_remove(remainder, "^[-:\\s]+")
      }

      list(
        type = type,
        details = remainder
      )
    },

    infer_label_from_signature = function(signature_line) {
      if (is.null(signature_line) || length(signature_line) == 0) {
        return(NA_character_)
      }
      sig <- str_trim(as.character(signature_line)[1])
      match <- str_match(
        sig,
        "^(?:Public|Private|Friend|Global|Static)\\s+(?:Property\\s+(?:Let|Set|Get)\\s+|Function\\s+|Sub\\s+)([A-Za-z0-9_]+)"
      )
      if (!is.na(match[1])) {
        return(match[2])
      }
      NA_character_
    },

    slugify_anchor = function(text) {
      if (!private$has_text(text)) {
        return("")
      }
      slug <- tolower(str_trim(as.character(text)[1]))
      slug <- str_replace_all(slug, "[^a-z0-9]+", "-")
      slug <- str_replace_all(slug, "-+", "-")
      slug <- str_replace(slug, "^-", "")
      slug <- str_replace(slug, "-$", "")
      slug
    }
  )
)

# Example usage:
# parser <- VBADocParser$new("src")
# parser$parse()
# parser$detect_usages()
# parser$extract_enums()
