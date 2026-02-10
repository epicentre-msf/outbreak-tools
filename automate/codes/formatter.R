# write a code that can format a .cls or .bas file and add indentations and
# return lines

#--- working on a module file
pacman::p_load(tidyverse, glue, here, fs)

KEYWORDS <- list(
  INDENT = c(
    "If",
    "For",
    "Do",
    "While",
    "With",
    "Select Case",
    "Function",
    "Sub",
    "Property",
    "Private Sub",
    "Public Sub",
    "Private Property",
    "Public Property",
    "Private Function"
  ),

  OUTDENT = c(
    "End If",
    "Next",
    "Loop",
    "Wend",
    "End With",
    "End Select",
    "End Function",
    "End Sub",
    "End Property",
    "Else",
    "ElseIf",
    "Case"
  ),

  BREAK = c("Then", "Or", "And", "\\+", "\\=")
)

# class for the formatting of one line
print.Lineformatting <- function(f) {
  glue("Line: '{f$formatted_lines}'", "\n", "Indent: {f$indent}")
}
indent.Lineformatting <- function(f) f$indent
stripped_lines.Lineformatting <- function(f) str_trim(f$formatted_lines)
# update the actual line formatting with a new line
update.Lineformatting <- function(f, new_line, indent = NULL) {
  indent_level <- f$indent
  if (!is.null(indent)) indent_level <- indent
  f$formatted_lines <- new_line
  return(f)
}

stripped.Lineformatting <- function(f) {
  output <- list(formatted_lines = stripped_lines(f), indent = indent(f))
}

fmerge <- function(f1, f2) {
  # merge two lineformatting objects together
  output <- list(
    formatted_lines = c(f1$formatted_lines, f2$formatted_lines),
    indent = max(f1$indent, f2$indent)
  )
  class(output) <- "Lineformatting"

  return(output)
}

MAX_LINE_WIDTH <- 80
INDENT_STRING <- "  "

# add indent to a a line
add_indent <- function(
  f, # line formatting
  indent_unit = INDENT_STRING,
  words = KEYWORDS
) {
  # current indent
  indent_level <- indent(f)
  strip_val <- stripped_lines(f)

  # if the stripped line is empty, stop here
  if (strip_val == "") return(stripped(f))

  # outindent regular expression
  out_indent <- glue("(({words$OUTDENT})$)") |>
    glue_collapse(sep = "|") |>
    as.character() |>
    regex(ignore_case = TRUE)

  in_indent <- glue("(({words$INDENT})$)") |>
    glue_collapse(sep = "|") |>
    as.character() |>
    regex(ignore_case = TRUE)

  # if there is a start key word, format
  if (str_detect(line, out_indent)) {
    indent_level <- max(indent_level - 1, 0)
  }

  # if there is a end key word, format
  if (str_detect(line, in_indent)) {
    indent_level <- indent_level + 1
  }

  # duplicate the indentation and paste it
  formatted_lines <- glue("{str_dup(indent_unit, indent_level)}{stripped}")

  output <- list(formatted_lines = formatted_lines, indent = indent_level)
  class(output) <- "Lineformatting"
  return(output)
}


# add break line to a line
break_line <- function(
  f, # line formatting
  words = KEYWORDS,
  max_width = MAX_LINE_WIDTH,
  indent_unit = INDENT_STRING
) {
  if (nchar(line) <= MAX_LINE_WIDTH) return(f)

  # start breaking the line
  remaining <- stripped_lines(f)
  prefix <- str_dup(indent_unit, indent_level)
  formatted_lines <- c()

  while (nchar(remaining) > max_width) {
    substring_to_search <- substr(remaining, 1, max_width)

    break_regex <- glue('((?=(?:[^"]*"[^"]*")*[^"]*$)({words$BREAK}))') |>
      glue_collapse(sep = "|") |>
      as.character() |>
      regex(ignore_case = TRUE)

    split_point <- str_locate(substring_to_search, break_regex)[, 2]
    if (is.na(split_point)) break

    chunk <- str_sub(substring_to_search, start = 1, end = split_point)

    # pasting the chunk with the return character
    formatted_lines <- c(formatted_lines, glue("{prefix}{chunk} _"))
    remaining <- str_sub(remaining, start = split_point + 1) |> str_trim()
  }

  formatted_lines <- c(formatted_lines, glue("{prefix}{remaining}"))

  output <- list(formatted_lines = formatted_lines, indent = indent_level)
  class(output) <- "Lineformatting"
  return(output)
}

# align comment on multiple lines
align_comment <- function(
  f, #Line formatting object
  indent_unit = INDENT_STRING,
  com_col = COMMENT_COLUMN
) {
  line <- stripped_lines(f)
  indent_level <- indent(f)

  if (!str_detect(line, "'")) return(f)

  # separate code and comment on line (especially if there is a
  # comment after the code)
  parts <- str_split_fixed(line, "'", 2)
  code <- str_trim(parts[1])
  prefix <- str_dup(indent_unit, indent_level)

  comment <- str_trim(parts[2])
  if (nchar(line) > 120) {
    # for multi lines, put the comment on top if too long
    formatted_lines <- c(comment, code)
    formatted_lines <- glue("{prefix}{formatted_lines}")
    output <- list(formatted_lines = formatted_lines, indent = indent_level)
    class(output) <- "Lineformatting"
    return(output)
  }

  # here, we have a length < 120 with code
  padding <- str_dup(indent_unit, max(com_col - nchar(code), 2))
  formatted_lines <- glue("{prefix}{code}{padding}{comment}")

  output <- list(formatted_lines = formatted_lines, indent = indent_level)
  class(output) <- "Lineformatting"
  return(output)
}

# align_dim_set <- function(line, indent_level = 1, indent_unit = INDENT_STRING) {
#   # if the line does not starts with a Dim or a Set, return  it as is
#   if (!str_detect(line, "^(Dim|Set)"))
#     return(list(formatted_lines = line, indent = indent_level))

#   if (indent_level > 2)
#     return(list(formatted_lines = line, indent = indent_level))

#   # Split Dim/Set declarations and align with tabs or spaces
#   m <- str_match(line, "^(Dim|Set)\\s+([a-zA-Z0-9_]+)(\\s+As\\s+[^']+)?(.*)")
#   if (is.na(m[1])) return(line)
#   keyword <- m[2]
#   var <- m[3]
#   type <- str_trim(ifelse(is.na(m[4]), "", m[4]))
#   rest <- str_trim(ifelse(is.na(m[5]), "", m[5]))

#   formatted_lines <- sprintf("%-6s %-20s %-25s %s", keyword, var, type, rest)
#   formatted_lines <- glue("{rep(indent_unit, indent_level))}{formatted_lines}")

#   return(list(formatted_lines = formatted_lines, indent = indent_level))
# }

# creating the formatting function
format_vbs <- function(file_path, out_file = file_path) {
  # reading all the lines of the a potential file to clean
  all_lines <- readLines(file_path, warn = FALSE)
  cleaned_lines <- character()
  indent <- 0
  i = 0
  for (line in all_lines) {
    #
    stripped <- str_trim(line)

    if (stripped == "") {
      cleaned_lines <- c(cleaned_lines, "")
      next
    }

    # align the dim and the set of a line
    form_obj <- align_dim_set(stripped, indent_level = indent)
    stripped <- form_obj |> pluck("formatted_lines")

    # add the indent
    form_obj <- add_indent(stripped, indent_level = indent)
    stripped <- form_obj |> pluck("formatted_lines")
    indent <- form_obj |> pluck("indent")

    # break lines in multiple lines
    form_obj <- break_line(stripped, indent_level = indent)
    stripped <- form_obj |> pluck("formatted_lines")

    cleaned_lines <- c(cleaned_lines, stripped)
    indent <- indent
    cleaned_lines
  }

  #
  writeLines(cleaned_lines, out_file)
}


format_vbs(here("automate/codes/tmp.cls"), out_file = "tmp.cls")
