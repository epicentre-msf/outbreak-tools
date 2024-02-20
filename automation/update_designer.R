update_designer <- function(update_status = 0, osname = "windows"){
  # move previous version of my designer
  # replace the dev designer
  if (update_status == 0) {
    file.copy(
      from = "./src/bin/designer_aky.xlsb",
      to = "./automation/", overwrite = TRUE
    )
    cat("Moved the development designer to automation folder", "\n")
    cat("\n")
    # move back and overwrite
    file.copy(
      from = "./automation/designer_aky.xlsb",
      to = "./src/bin/designer_dev.xlsb", overwrite = TRUE
    )
    cat("Replaced previous development designer by aky designer", "\n")
    cat("\n")
  }
  # update the stable version if needed
  if (update_status == 1) {
    # previous stable version
    file.copy(
      from = "./designer.xlsb",
      to = "./automation/designer_prev.xlsb", overwrite = TRUE
    )
    cat("Saved previous main designer", "\n")
    cat("\n")

    # update the new stable version
    file.copy(
      from = "./automation/designer_aky.xlsb",
      to = "./designer.xlsb", overwrite = TRUE
    )
    cat("Replaced the main designer by aky designer", "\n")
    cat("\n")
  }

  # revert back previous stable designer due to corrupt files.
  if (update_status == 2) {
    file.copy(
      from = "./automation/designer_prev.xlsb",
      to = "./designer.xlsb", overwrite = TRUE
    )

    cat("Replace the actual main designer by the previous version", "\n")
    cat("\n")

    file.copy(
      from = "./automation/designer_prev.xlsb",
      to = "./src/bin/designer_dev.xlsb", overwrite = TRUE
    )
    cat("Replaced the development designer by previous main designer", "\n")
    cat("\n")

    file.copy(
      from = "./automation/designer_prev.xlsb",
      to = "./src/bin/designer_aky.xlsb", overwrite = TRUE
    )

    cat("Replaced the aky designer by previous main designer---", "\n")
    cat("\n")
  }

  # copy the file from the mock designer
  if (update_status == 3) {
    file.copy(
      from = glue::glue("./src/.mock/designer_mock_{osname}.xlsb"),
      to = "./src/bin/designer_aky.xlsb", overwrite = TRUE
    )

    cat("Replaced the aky designer by previous mock designer", "\n")

  }
}

sysname  <-  tolower(Sys.info()[["sysname"]])

# copy the mock file for development
update_designer(update_status = 3, osname = sysname)
# update the dev file
update_designer(update_status = 0, osname = sysname)
# update the file on root
update_designer(update_status = 1, osname = sysname)
