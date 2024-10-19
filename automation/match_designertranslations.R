# This is a quick code to match translations from different_files

library(openxlsx)
library(purrr)
library(dplyr)
library(stringr)

#----- copy the designer translation user file from the drive
obt_onedrivedev_folder <- normalizePath(Sys.getenv("OBTDEV"))
onedrive_translation <- glue::glue(
  "{obt_onedrivedev_folder}/ressources/designer_translations_user.xlsx"
)
obt_repo <- here::here()

file.copy(from = onedrive_translation, to = glue::glue("{obt_repo}/misc/"), 
          overwrite = TRUE)

# path to the master designer_translation workbook
master_workbook_path <- glue::glue("{obt_repo}/misc/designer_translations.xlsx")

# path to a workbook where corrections have been made
child_workbook_path <- glue::glue(
  "{obt_repo}/misc/designer_translations_user.xlsx"
)

#---- Preliminary functions ----------------------------------------------------

# get the tables with their corresponding references in one workbook
get_all_tables <- function(sheet_name, path = master_workbook_path) {
  wb <- loadWorkbook(path)
  tabs_list <- getTables(wb, sheet_name)

  data.frame(
    path = as.character(path),
    sheet_name = sheet_name,
    table_name = tabs_list,
    table_reference = names(tabs_list)
  ) |>
    mutate(
      range_address = glue::glue("{sheet_name}!{table_reference}")
    )
}

# Load a table in a workbook
read_a_table <- function(path, sheet_name, range_address, child_tag = "") {
  df <- readxl::read_excel(path, range = range_address) |>
    janitor::clean_names()

  colnames(df) <- glue::glue("{colnames(df)}{child_tag}")

  new_columns <- colnames(df)
  new_columns[1] <- "msg_id"
  colnames(df) <- new_columns

  df |>
    # select names in that order
    filter(!is.na(msg_id)) |>
    select(
      msg_id, starts_with("eng"), starts_with("fra"),
      starts_with("spa"), starts_with("por"), starts_with("ara")
    )
}

# compare elements and replace between child and master 
#(child contains the corrections)

compare_elements <- function(master_data, child_data) {
  new_data <- left_join(master_data, child_data,
    by = "msg_id"
  )
  new_data <- new_data |>
    mutate(
      eng = case_when(
        (!is.na(eng_child)) & (eng != eng_child) ~ eng_child,
        TRUE ~ eng
      ),
      fra = case_when(
        (!is.na(fra_child)) & (fra != fra_child) ~ fra_child,
        TRUE ~ fra
      ),
      spa = case_when(
        (!is.na(spa_child)) & (spa != spa_child) ~ spa_child,
        TRUE ~ spa
      ),
      por = case_when(
        (!is.na(por_child)) & (por != por_child) ~ por_child,
        TRUE ~ por
      ),
      ara = case_when(
        (!is.na(ara_child)) & (ara != ara_child) ~ ara_child,
        TRUE ~ ara
      )
    ) |>
    select(
      msg_id, eng, fra, spa, por, ara
    ) |>
    arrange(msg_id)

  # upper case of the colnames to ensure everything is clear
  colnames(new_data) <- toupper(colnames(new_data))
  new_data
}

targeted_tables <- c(
  "T_TradLLShapes", "T_TradLLMsg",
  "T_TradLLForms", "T_TradLLRibbon",
  "T_tradMsg", "T_tradRange", "T_tradShape"
)

# loading the master workbook tables ===========================================

master_workbook <- loadWorkbook(file = master_workbook_path)
# get the list of tables of the master workbook in each worksheet
master_sheets_list <- getSheetNames(file = master_workbook_path)

master_tables_df <- purrr::map_dfr(master_sheets_list, get_all_tables)

# load master table
master_list_tables <- master_tables_df |>
  filter((table_name %in% targeted_tables) |
           (table_name %in% tolower(targeted_tables))) |>
  arrange(table_name) |>
  select(path, range_address) |>
  purrr::pmap(read_a_table)

# loading the child workbook tables ============================================

child_workbook <- loadWorkbook(file = child_workbook_path)
child_sheets_list <- getSheetNames(file = child_workbook_path)

child_tables_df <- purrr::map_dfr(
  child_sheets_list, get_all_tables,
  path = child_workbook_path
)

child_list_tables <- child_tables_df |>
  filter((table_name %in% targeted_tables) |
    (table_name %in% tolower(targeted_tables))) |>
  arrange(table_name) |>
  select(path, range_address) |>
  purrr::pmap(read_a_table, child_tag = "_child")

# now join and clean everything

joined_data <- purrr::map2(
  master_list_tables, child_list_tables,
  compare_elements
)

targeted_tables <- sort(targeted_tables)

names(joined_data) <- sort(targeted_tables)


# write back to an empty workbook ==============================================
source(glue::glue("{obt_repo}/automation/functions_tabulations.R"))

header_names <- c(
  "Tables for translations of the linelist",
  "Tables for translations of the designer"
)

wb <- initiate_workbook(
  sheetnames = master_sheets_list,
  headernames = header_names
)

# linelist_translations --------------------------------------------------------

ll_tables_names <- c(
  "T_TradLLShapes", "T_TradLLMsg",
  "T_TradLLForms", "T_TradLLRibbon"
)

ll_tables <- joined_data[ll_tables_names]
ll_tables_labels <- c(
  "Translation of shapes (buttons in the the linelist worksheets)",
  glue::glue("Translation of various messages of the user interface, ",
             "including special worksheets names"),
  "Translation of forms in the linelist",
  "Translation of elements of the ribbon menu in the linelist"
)

names(ll_tables_labels) <- ll_tables_names

push_all_tables(
  wb,
  listoflabels = ll_tables_labels,
  sheetname = master_sheets_list[1],
  listofnames = ll_tables_names,
  listoftables = ll_tables
)

# designer translations --------------------------------------------------------

des_tables_names <- c("T_tradMsg", "T_tradRange", "T_tradShape", "T_tradDrop")

des_tables <- joined_data[des_tables_names]
des_tables_labels <- c(
  glue::glue("Translation of various messages in main worksheet, ",
             "including ribbon menu elements"),
  "Translation of messages in ranges of main worksheet",
  "Translation of shapes (buttons in the main worksheet)"
)

names(des_tables_labels) <- des_tables_names

push_all_tables(
  wb,
  listoflabels = des_tables_labels,
  sheetname = master_sheets_list[2],
  listofnames = des_tables_names,
  listoftables = des_tables
)

saveWorkbook(wb,
  file = glue::glue("{obt_repo}/misc/designer_translations_merged.xlsx"),
  overwrite = TRUE
)

file.copy(
  from = glue::glue("{obt_repo}/misc/designer_translations_merged.xlsx"),
  to = onedrive_translation, overwrite = TRUE
)
