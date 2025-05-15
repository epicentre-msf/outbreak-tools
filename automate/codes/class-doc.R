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
      output_folder = here(folder, "docs"),
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
      i <- 1L
      n <- length(lines)

      while (i <= n) {
        line <- str_trim(lines[i])

        if (str_detect(line, "^'@label:")) {
          doc <- list()
          headers <- list()
          doc$label <- str_remove(line, "^'@label:\\s*")
          i <- i + 1L

          while (i <= n && str_detect(lines[i], "^'")) {
            l <- str_trim(lines[i])
            tag_match <- str_match(l, "^'@([\\-a-z]+):?\\s*(.*)$")
            if (!is.na(tag_match[1])) {
              tag <- str_to_lower(tag_match[2])
              content <- str_trim(tag_match[3])
              if (tag == "details") {
                desc <- content
                i <- i + 1L
                while (i <= n && str_detect(lines[i], "^'[^@]")) {
                  desc <- paste0(desc, "\n", str_remove(lines[i], "^'"))
                  i <- i + 1L
                }
                doc$details <- desc
                next
              } else if (tag == "param") {
                if (is.null(doc$params)) doc$params <- list()
                param_parts <- str_match(content, "(\\w+)(.*)$")
                if (!is.na(param_parts[1])) {
                  doc$params <- append(
                    doc$params,
                    list(list(
                      name = param_parts[2],
                      details = str_trim(param_parts[3])
                    ))
                  )
                }
              } else if (tag %in% c("section", "sub-title", "prop-title")) {
                headers <- list(entry = content, tag = tag)
                doc[[tag]] <- content
              } else {
                doc[[tag]] <- content
              }
            }
            i <- i + 1L
          }

          if (
            i <= n &&
              str_detect(
                str_trim(lines[i]),
                "^(Public|Private)\\s+(Sub|Function|Property)"
              )
          ) {
            signature <- lines[i]
            while (!str_detect(lines[i], "\\)")) {
              i <- i + 1L
              signature <- glue("{signature}{lines[i]}")
            }
            doc$signature <- str_replace_all(
              signature,
              ",\\s*_\\s+",
              ", _\n     "
            )

            if (!is.null(doc$export)) {
              tocs <- append(tocs, list(headers))
              externals <- append(externals, list(doc))
            } else {
              internals <- append(internals, list(doc))
            }
            i <- i + 1L
          }
        } else {
          i <- i + 1L
        }
      }
      list(externals = externals, internals = internals, tocs = tocs)
    },

    build_markdown = function(
      class_name,
      externals,
      internals,
      tocs,
      class_names
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

      output <- c(
        output,
        "</details>",
        "\n---\n",
        private$resolve_doc(externals, class_names),
        ""
      )

      output <- c(
        output,
        "",
        "::: {.callout-note collapse=\"true\" title=\"Additional not exported Subs \"}",
        glue("{private$resolve_doc(internals, class_names)}"),
        ":::"
      )
      paste(output, collapse = "\n")
    },

    resolve_links = function(text, class_names) {
      str_replace_all(text, "see::([A-Za-z0-9_]+)", function(m) {
        cls <- str_match(m, "see::([A-Za-z0-9_]+)")[, 2]
        if (cls %in% class_names) glue("[{cls}]({cls}.md)") else cls
      })
    },

    resolve_doc = function(lst_doc, class_names) {
      output <- ""
      for (doc in lst_doc) {
        label <- doc$label
        sig <- doc$signature
        desc <- doc$details

        if (!is.null(doc[["prop-title"]])) {
          anchor <- tolower(doc[["prop-title"]]) |>
            trimws() |>
            str_remove_all("[[:punct:]]") |>
            str_replace_all("\\s+", "-")
          title <- glue("`{label}` {{#sec-{anchor}}}")
          desc <- glue("{doc[['prop-title']]}")
        } else {
          anchor <- tolower(doc[["sub-title"]]) |>
            trimws() |>
            str_remove_all("[[:punct:]]") |>
            str_replace_all("\\s+", "-")
          title <- glue("`{label}` {{#sec-{anchor}}}")
          desc <- glue("{doc[['sub-title']]} ")
        }

        block <- c(
          glue("\n### {title}"),
          "",
          glue("**{desc}**"),
          "",
          "**Signature:**",
          "\n```vb",
          sig,
          "```",
          ""
        )

        if (!is.null(doc$params)) {
          block <- c(block, "**Parameters:**", "")
          for (p in doc$params) {
            block <- c(block, glue("  - `{p$name}`: {p$details}"))
          }
          block <- c(block, "")
        }

        if (!is.null(desc)) {
          desc <- private$resolve_links(desc, class_names)
          block <- c(block, glue("**Details:**\n\n{desc}"), "")
        }

        if (!is.null(doc$returned)) {
          block <- c(block, glue("**Return: {doc$returned}**"))
        }

        output <- c(output, block, "\n---\n")
      }
      paste(output, collapse = "\n")
    }
  )
)

# Example usage:
# parser <- VBADocParser$new("src")
# parser$parse()
# parser$detect_usages()
# parser$extract_enums()
