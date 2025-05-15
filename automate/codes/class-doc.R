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
      dir_create(output_folder)
      cli::cli_inform("Parser initialized for folder: {folder}")
      cli::cli_inform("Parser will write in: {output_folder}")
    },

    init_folder = function() private$folder,
    write_folder = function() private$output_folder,
    doc_classes = function() private$class_names,

    parse = function() {
      cls_files <- dir_ls(
        self$init_folder(),
        regexp = "\\.cls$",
        recurse = TRUE
      )
      private$class_names <- path_ext_remove(path_file(cls_files))

      for (idx in seq_along(cls_files)) {
        file <- cls_files[idx]
        class_name <- self$doc_classes()[idx]
        cli::cli_inform("Parsing: {path_file(file)}")

        lines <- read_lines(file)
        doc_info <- private$extract_doc_blocks(lines)

        if (length(doc_info$internals) > 0 || length(doc_info$externals) > 0) {
          md <- private$build_markdown(
            class_name,
            doc_info$externals,
            doc_info$internals,
            doc_info$tocs,
            self$doc_classes()
          )
          write_file(md, path(self$write_folder(), glue("{class_name}.md")))
          cli::cli_inform("Written: {class_name}.md")
        } else {
          cli::cli_alert_warning(
            "No valid documentation blocks found in {class_name}"
          )
        }
      }
    },

    detect_usages = function() {
      all_files <- dir_ls(
        self$init_folder(),
        recurse = TRUE,
        regexp = "\\.(cls|bas)$"
      )
      usage_map <- setNames(
        vector("list", length(self$doc_classes())),
        self$doc_classes()
      )

      for (file in all_files) {
        content <- read_file(file)
        for (class_name in self$doc_classes()) {
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
          }
          write_lines(md, md_path)
          cli::cli_inform(glue("Updated usages for: {class_name}"))
        }
      }
    }
  ),

  private = list(
    folder = NULL,
    output_folder = NULL,
    class_names = NULL,
    proj_path = NULL,

    extract_doc_blocks = function(lines) {
      external <- list()
      internal <- list()
      toc <- list()
      i <- 1
      n <- length(lines)

      while (i <= n) {
        line <- str_trim(lines[i])

        if (str_detect(line, "^'@label:")) {
          doc <- list()
          headers <- list()

          doc$label <- str_remove(line, "^'@label:\\s*")

          i <- i + 1
          # Parse doc lines
          while (i <= n && str_detect(lines[i], "^'")) {
            l <- str_trim(lines[i])
            tag_match <- str_match(l, "^'@([\\-a-z]+):?\\s*(.*)$")
            if (!is.na(tag_match[1, 1])) {
              tag <- str_to_lower(tag_match[1, 2])
              content <- str_trim(tag_match[1, 3])
              if (tag == "description") {
                desc <- content
                i <- i + 1
                while (i <= n && str_detect(lines[i], "^'[^@]")) {
                  desc <- paste0(desc, "\n", str_remove(lines[i], "^'"))
                  i <- i + 1
                }
                doc$description <- desc
                next
              } else if (tag == "param") {
                if (is.null(doc$params)) doc$params <- list()
                param_parts <- str_match(content, "(\\w+)(.*)$")
                if (!is.na(param_parts[1, 1])) {
                  params <- list(
                    name = param_parts[1, 2],
                    description = param_parts[1, 3]
                  )
                  doc$params <- append(doc$params, list(params))
                }
              } else if (tag %in% c("section", "sub-title", "prop-title")) {
                headers <- list(entry = content, tag = tag)
                doc[[tag]] <- content
              } else {
                doc[[tag]] <- content
              }
            }
            i <- i + 1
          }

          # check for Sub/Property line
          if (
            i <= n &&
              str_detect(
                str_trim(lines[i]),
                "^(Public|Private)\\s+(Sub|Function|Property)"
              )
          ) {
            signature <- lines[i]
            # check for the whole signature
            while (!str_detect(lines[i], "\\)")) {
              i <- i + 1
              signature <- glue("{signature}{lines[i]}")
            }

            doc$signature <- str_remove_all(signature, "_")

            if (!is.null(doc$export)) {
              toc <- append(toc, list(headers))
              external <- append(external, list(doc))
            } else {
              # @labels without export, internal content
              internal <- append(internal, list(doc))
            }

            i <- i + 1
          }
        } else {
          i <- i + 1
        }
      }
      list(externals = external, internals = internal, tocs = toc)
    },

    build_markdown = function(
      class_name,
      externals,
      internals,
      tocs,
      class_names
    ) {
      output <- c(glue("# {class_name}"), "")

      # creating the table of contents output
      output <- c(output, "## Table of Contents")
      for (toc in tocs) {
        if (!is.null(toc$entry)) {
          entry <- toc$entry
          anchor <- tolower(entry) |> trimws() |> str_replace_all("\\s+", "-")
          output <- c(output, glue("- [{entry}](#{anchor})"), "")
        }
      }

      # writing the external parts of the document
      output <- c(output, private$resolve_doc(externals), "")

      # writing the internal parts of the doucment
      output <- c(
        output,
        "",
        "<details>",
        "<summary> Not exported </summary>",
        glue("{private$resolve_doc(internals)}"),
        "</details>"
      )

      return(paste(output, collapse = "\n"))
    },

    resolve_links = function(text, class_names) {
      str_replace_all(text, "see::([A-Za-z0-9_]+)", function(m) {
        cls <- str_match(m, "see::([A-Za-z0-9_]+)")[, 2]
        if (cls %in% class_names) {
          glue("[{cls}]({cls}.md)")
        } else {
          glue("{cls}")
        }
      })
    },
    resolve_doc = function(lst_doc) {
      # resolve the documents
      output <- "---\n"

      for (doc in lst_doc) {
        label <- doc[["label"]]
        sig <- doc[["signature"]]
        desc <- doc[["description"]]

        title <- ifelse(
          !is.null(doc[["prop-title"]]),
          glue("Property `{label}`: {doc[['prop-title']]}"),
          glue("Sub `{label}`: {doc[['sub-title']]}")
        )

        output <- c(output, glue("\n### {title}"))
        output <- c(output, "")
        output <- c(output, glue("**Signature:** `{sig}`"), "")

        if (!is.null(desc)) {
          desc <- private$resolve_links(desc, class_names)
          output <- c(output, glue("**Description:**\n\n{desc}"), "")
        }

        if (!is.null(doc$params)) {
          output <- c(output, "**Parameters:**")
          for (p in doc$params) {
            output <- c(output, glue("- `{p$name}`: {p$description}"))
          }
          output <- c(output, "")
        }
        output <- c(output, "\n---\n")
      }
      paste(output, collapse = "\n")
    }
  )
)

# Example usage:
# parser <- VBADocParser$new(here::here("src/classes"), proj_path = here::here())
# parser$parse()
# parser$detect_usages()
