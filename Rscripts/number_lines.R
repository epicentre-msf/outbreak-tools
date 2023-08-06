# a function to readLines from a file

read_cls <- function(path) {
  nb_lines <- readLines(path) |> length()
  data.frame(file = basename(path), number_lines = nb_lines)
}

classes_implements <- list.files("src/classes/implements/", full.names = TRUE)
classes_interface <- list.files("src/classes/interfaces", full.names = TRUE)
modules_implements <- list.files("src/modules/implements", full.names = TRUE)
modules_tests <- list.files("src/modules/tests", full.names = TRUE)


# count the number of lines for each files in the classes and modules folder.

df_classes_implements <- purrr::map_dfr(classes_implements, read_cls) |>
  # remove BetterArray class (which is imported)
  dplyr::filter(file != "BetterArray.cls")

df_classes_interface <- purrr::map_dfr(classes_interface, read_cls)
df_modules_implements <- purrr::map_dfr(modules_implements, read_cls)
df_modules_tests <- purrr::map_dfr(modules_tests, read_cls)


# creating the database for the classes

df_classes <- df_classes_interface |>
  dplyr::mutate(interface_name = file) |>
  dplyr::mutate(file = stringr::str_remove(file, "^I")) |>
  dplyr::rename(number_interface_lines = number_lines) |>
  dplyr::right_join(df_classes_implements, by = "file") |>
  dplyr::rename(
    number_implement_lines = number_lines,
    implement_name = file
  ) |>
  dplyr::mutate(tot_number_lines = number_interface_lines +
    number_implement_lines) |>
  dplyr::arrange(desc(tot_number_lines)) |>
  dplyr::select(c(ends_with("name"), ends_with("lines")))


sink("src/classes_lines.txt")
knitr::kable(df_classes)
# print the total number of lines for the class
cat("\n\n")
glue::glue("Total number of classes lines: {sum(df_classes$tot_number_lines)}")

sink()

# creating database for the modules
df_modules_implements <- df_modules_implements |>
  dplyr::rename(module_name = file) |>
  dplyr::arrange(desc(number_lines))

sink("src/modules_lines.txt")

knitr::kable(df_modules_implements)
cat("\n\n")
glue::glue(
  "Total number of modules lines: ",
  "{sum(df_modules_implements$number_lines)}"
)
sink()
