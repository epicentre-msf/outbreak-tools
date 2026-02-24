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

        # Fallback: use filesystem directory name as folder
        if (!private$has_text(doc_info$header$folder)) {
          doc_info$header$folder <- str_to_title(basename(dirname(file)))
        }

        # For interfaces, treat all members as external API
        is_iface <- str_detect(
          paste(lines, collapse = "\n"), "'@Interface"
        )
        if (is_iface) {
          all_ext <- c(doc_info$externals, doc_info$internals)
          ext <- all_ext
          int <- list()
        } else {
          ext <- doc_info$externals
          int <- doc_info$internals
        }

        if (length(ext) > 0 || length(int) > 0) {
          md <- private$build_markdown(
            class_name,
            ext,
            int,
            private$class_names,
            header = doc_info$header,
            interface_mode = is_iface
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

      if (!private$has_text(doc_info$header$folder)) {
        doc_info$header$folder <- str_to_title(basename(dirname(target_file)))
      }

      if (length(doc_info$internals) > 0 || length(doc_info$externals) > 0) {
        md <- private$build_markdown(
          class_name,
          doc_info$externals,
          doc_info$internals,
          private$class_names,
          header = doc_info$header,
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

      if (!private$has_text(doc_info$header$folder)) {
        doc_info$header$folder <- str_to_title(basename(dirname(iface_file)))
      }

      # Treat all documented members as external API for interfaces
      docs_external <- c(doc_info$externals, doc_info$internals)
      docs_internal <- list()

      if (length(docs_external) > 0) {
        md <- private$build_markdown(
          iface_name,
          docs_external,
          docs_internal,
          private$class_names,
          header = doc_info$header,
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

      # Class-level header metadata
      header <- list(
        class_name = NULL, folder = NULL, module_desc = NULL,
        description = NULL, depends = NULL, version = NULL, author = NULL
      )
      in_header <- FALSE

      n <- length(lines)
      current_section <- NULL
      doc <- NULL
      last_tag <- NULL
      last_param_index <- NA_integer_
      last_note_index <- NA_integer_
      last_remark_index <- NA_integer_
      last_throw_index <- NA_integer_
      last_depend_index <- NA_integer_
      last_todo_index <- NA_integer_
      i <- 1L

      reset_tracking <- function() {
        last_tag <<- NULL
        last_param_index <<- NA_integer_
        last_note_index <<- NA_integer_
        last_remark_index <<- NA_integer_
        last_throw_index <<- NA_integer_
        last_depend_index <<- NA_integer_
        last_todo_index <<- NA_integer_
      }
      reset_tracking()

      finalize_doc <- function(entry, signature_lines) {
        if (length(signature_lines) == 0) {
          return()
        }

        entry$signature <- paste(signature_lines, collapse = "\n")

        # Always infer the function/sub/property name from the signature
        inferred <- private$infer_label_from_signature(signature_lines[[1]])
        if (private$has_text(inferred)) {
          entry$name <- inferred
        } else {
          entry$name <- str_trim(signature_lines[[1]])
        }

        # If no label was set, fall back to the inferred name
        if (!private$has_text(entry$label)) {
          entry$label <- entry$name
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

          # Normalize deprecated tag variants
          if (tag == "pram") tag <- "param"
          if (tag %in% c("params", "parameters")) tag <- "params"
          if (tag %in% c("returns", "returned")) tag <- "return"
          if (tag == "fun-title") tag <- "sub-title"
          if (tag == "remark") tag <- "remarks"

          # --- @class: start header block ---
          if (tag == "class") {
            if (private$has_text(content)) header$class_name <- content
            in_header <- TRUE
            reset_tracking()
            i <- i + 1L
            next
          }

          # --- @folder: extract from ("Name") ---
          if (tag == "folder") {
            folder_match <- str_match(content, '\\("([^"]+)"\\)')
            if (!is.na(folder_match[1])) {
              header$folder <- folder_match[2]
            }
            i <- i + 1L
            next
          }

          # --- @moduledescription: extract from ("text") ---
          if (tag == "moduledescription") {
            md_match <- str_match(content, '\\("([^"]+)"\\)')
            if (!is.na(md_match[1])) {
              header$module_desc <- md_match[2]
            }
            i <- i + 1L
            next
          }

          # --- Other module-level tags: skip ---
          module_level <- c(
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

          # --- @section: end header mode, start section ---
          if (tag == "section") {
            current_section <- if (private$has_text(content)) content else NULL
            in_header <- FALSE
            doc <- NULL
            reset_tracking()
            i <- i + 1L
            next
          }

          # --- Header-mode tags (before first @section) ---
          if (in_header) {
            if (tag %in% c("details", "description")) {
              header$description <- private$append_text(
                header$description, content
              )
              last_tag <- "header_description"
            } else if (tag == "depends") {
              deps <- if (private$has_text(content)) {
                str_split(content, "\\s*,\\s*")[[1]]
              } else {
                character()
              }
              header$depends <- c(header$depends, deps)
              last_tag <- "header_depends"
            } else if (tag == "version") {
              header$version <- content
            } else if (tag == "author") {
              header$author <- content
            } else {
              # Unrecognised header tag; exit header mode
              in_header <- FALSE
            }
            if (in_header) {
              i <- i + 1L
              next
            }
            # Fell through: process as normal member tag below
          }

          # --- @jump outside a doc block: skip ---
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
          } else if (tag %in% c("sub-title", "prop-title")) {
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
            deps <- if (private$has_text(content)) {
              str_split(content, "\\s*,\\s*")[[1]]
            } else {
              ""
            }
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
          } else if (tag == "todo") {
            doc$todo <- c(doc$todo, content)
            last_tag <- "todo"
            last_todo_index <- length(doc$todo)
          } else if (tag == "version") {
            doc$version <- content
          } else if (tag == "author") {
            doc$author <- content
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

        # --- Continuation comments ---
        if (str_detect(trimmed, "^'(?!@)")) {
          text <- str_trim(str_remove(trimmed, "^'"))

          # Header-mode continuation
          if (in_header) {
            if (identical(last_tag, "header_description")) {
              header$description <- private$append_text(
                header$description, text
              )
            } else if (identical(last_tag, "header_depends")) {
              deps <- str_split(text, "\\s*,\\s*")[[1]]
              header$depends <- c(header$depends, deps)
            }
            i <- i + 1L
            next
          }

          # Member-level continuation
          if (!is.null(doc)) {
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
            } else if (identical(last_tag, "todo") && !is.na(last_todo_index)) {
              doc$todo[last_todo_index] <- private$append_text(doc$todo[last_todo_index], text)
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

        # --- Code lines (signatures, blanks, attributes) ---
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

      list(
        externals = externals,
        internals = internals,
        tocs = tocs,
        header = header
      )
    },

    build_markdown = function(
      class_name,
      externals,
      internals,
      class_names,
      header = list(),
      interface_mode = FALSE
    ) {
      output <- c()

      # Folder metadata (consumed by build-site.R for index grouping)
      folder_name <- if (private$has_text(header$folder)) header$folder else ""
      output <- c(output, glue("<!-- folder: {folder_name} -->"), "")

      # Class heading
      output <- c(output, glue("# {class_name}"), "")

      # Class description
      if (private$has_text(header$description)) {
        output <- c(output, header$description, "")
      }

      # Class-level depends
      if (!is.null(header$depends) && length(header$depends) > 0) {
        valid_deps <- header$depends[nzchar(str_trim(header$depends))]
        if (length(valid_deps) > 0) {
          dep_links <- vapply(valid_deps, function(d) {
            d <- str_trim(d)
            if (d %in% class_names) as.character(glue("[{d}]({d}.html)")) else d
          }, character(1))
          output <- c(
            output,
            glue("**Depends on:** {paste(dep_links, collapse = ', ')}"),
            ""
          )
        }
      }

      # Version / Author
      if (private$has_text(header$version)) {
        output <- c(output, glue("**Version:** {header$version}"), "")
      }
      if (private$has_text(header$author)) {
        output <- c(output, glue("**Author:** {header$author}"), "")
      }

      # Exported members
      external_md <- private$resolve_doc(externals, class_names)
      if (private$has_text(external_md)) {
        output <- c(output, external_md, "")
      }

      # Internal members (not for interfaces)
      if (!interface_mode) {
        internal_md <- private$resolve_doc(internals, class_names)
        if (private$has_text(internal_md)) {
          output <- c(
            output,
            "",
            "<details>",
            "<summary>Internal members (not exported)</summary>",
            "",
            internal_md,
            "",
            "</details>"
          )
        }
      }

      paste(output, collapse = "\n")
    },

    resolve_links = function(text, class_names) {
      str_replace_all(text, "see::([A-Za-z0-9_]+)", function(m) {
        cls <- str_match(m, "see::([A-Za-z0-9_]+)")[, 2]
        if (cls %in% class_names) glue("[{cls}]({cls}.html)") else cls
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

        # Use the function/sub/property name as the h3 title
        member_name <- if (private$has_text(doc$name)) {
          as.character(doc$name)[1]
        } else if (private$has_text(doc$label)) {
          as.character(doc$label)[1]
        } else {
          "UnnamedMember"
        }

        # Label shown as subtitle when it differs from the name
        label <- if (private$has_text(doc$label)) {
          as.character(doc$label)[1]
        } else {
          NULL
        }

        summary <- NULL
        if (private$has_text(doc[["prop-title"]])) {
          summary <- doc[["prop-title"]]
        } else if (private$has_text(doc[["sub-title"]])) {
          summary <- doc[["sub-title"]]
        }

        header <- glue("### {member_name}")

        block <- c("", header)

        # Show label as subtitle if it differs from the function name
        if (!is.null(label) && label != member_name) {
          block <- c(block, "", glue("*{label}*"))
        }

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
          return_type <- if (private$has_text(doc$return$type)) glue("{doc$return$type} -- ") else ""
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

        if (!is.null(doc$todo) && length(doc$todo) > 0) {
          todo_block <- character()
          for (todo_entry in doc$todo) {
            if (private$has_text(todo_entry)) {
              todo_block <- c(todo_block, glue("  - {todo_entry}"))
            }
          }
          if (length(todo_block) > 0) {
            block <- c(block, "", "**Todo:**", todo_block)
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
