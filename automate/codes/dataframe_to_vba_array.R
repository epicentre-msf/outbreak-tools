# Convert a data frame into a VBA Array(Array(...)) expression.
df_to_vba_array <- function(df) {
  if (!is.data.frame(df)) {
    stop("df must be a data frame")
  }

  if (!nrow(df)) {
    return("Array()")
  }

  clean_df <- data.frame(lapply(df, as.character), stringsAsFactors = FALSE)
  clean_df[is.na(clean_df)] <- ""

  escape_vba <- function(x) gsub('"', '""', x, fixed = TRUE)

  row_expr <- apply(clean_df, 1, function(row) {
    row <- escape_vba(row)
    paste0('Array("', paste0(row, collapse = '","'), '")')
  })

  row_expr
}

# Reading the dictionary
df <- read.csv(here::here("temp/draft.csv"))
expr_vba <- df_to_vba_array(df)


sink("temp/draft.txt")
cat("Array( _\n")

for (i in 1:length(expr_vba)) {
  cat(expr_vba[i])
  if ((i %% 15) == 0) {
    cat(" _\n)")
    cat("\n\n")
    cat("Array( _\n")
  } else {
    cat(", _\n")
  }
}

cat(", _\n)")
cat("\n\n")
sink()
