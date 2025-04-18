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

MAX_LINE_WIDTH <- 80
INDENT_STRING <- "  "


# add indent to a a line
add_indent <- function(
  line,
  indent_level = 0,
  indent_unit = INDENT_STRING,
  words = KEYWORDS
) {
  # current indent
  indent <- indent_level
  stripped <- str_squish(line)
  formatted_line <- stripped

  # outindent regular expression
  out_indent <- glue("(({words$OUTDENT})$)") |>
    glue_collapse(sep = "|") |>
    as.character() |>
    regex(ignore_case = TRUE)

  in_indent <- glue("(({words$INDENT})$)") |>
    glue_collapse(sep = "|") |>
    as.character() |>
    regex(ignore_case = TRUE)

  # if the stripped line is empty, stop here
  if (stripped != "") {
    # if there is a start key word, format
    if (str_detect(line, out_indent)) {
      indent <- max(indent - 1, 0)
    }

    # if there is a end key word, format
    if (str_detect(line, in_indent)) {
      indent <- indent + 1
    }

    # duplicate the indentation and paste it
    indent_part <- str_dup(indent_unit, indent)
    formatted_line <- glue("{indent_part}{stripped}")
  }

  return(list(formatted_lines = formatted_line, indent = indent))
}


# add break line to a line
break_line <- function(
  line,
  indent_level,
  words = KEYWORDS,
  max_width = MAX_LINE_WIDTH,
  indent_unit = INDENT_STRING
) {
  if (nchar(line) <= MAX_LINE_WIDTH)
    return(list(formatted_lines = line, indent = indent_level))

  # start breaking the line
  remaining <- str_squish(line)
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
    remaining <- str_sub(remaining, start = split_point + 1) |> str_squish()
  }

  formatted_lines <- c(formatted_lines, glue("{prefix}{remaining}"))

  return(list(formatted_lines = formatted_lines, indent = indent_level))
}

# align comment on multiple lines
align_comment <- function(line, indent_level, com_col = COMMENT_COLUMN) {
  stripped <- str_squish(line)
  if (!str_detect(line, "'"))
    return(list(formatted_lines = stripped, indent = indent_level))

  # separate code and comment on line (especially if there is a
  # comment after the code)
  parts <- str_split_fixed(line, "'", 2)
  code <- part
}
