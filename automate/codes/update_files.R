#'@param tag: a tag to tell R which file to replace
#' (different mock files depending on the OS)

pacman::p_load(fs, cli, glue, here)

# update the designer
update_designer <- function(tag = "mock") {
  # move previous version of my designer
  # replace the dev designer
  switch(
    tag,

    # update the actual mock file
    mock = {
      file_copy(
        here(".mock", "designer_mock.xlsb"),
        here("src", "bin", "designer", "designer_dev.xlsb"),
        overwrite = TRUE
      )
      cli_alert_success("Successfully copied the designer mock file")
    },

    # udate the main designer file on the designer folder
    #(the one without the _dev tag)
    main = {
      file_copy(
        here("src", "designer", "designer_dev.xlsb"),
        here("src", "designer", "designer.xlsb"),
        overwrite = TRUE
      )
      cli_alert_success("Sucessfully replaced the designer main file")
    }
  )
  return(invisible())
}


# update the setup file
update_setup <- function(tag = "mock") {
  # move previous version of my designer
  # replace the dev designer
  switch(
    tag,
    # update the actual mock file
    mock = {
      file_copy(
        here(".mock", "setup_mock.xlsb"),
        here("src", "bin", "setup", "setup_dev.xlsb"),
        overwrite = TRUE
      )
      cli_alert_success("Sucessfully copied the setup mock file")
    },

    # udate the main designer file on the designer folder
    #(the one without the _dev tag)
    main = {
      file_copy(
        here("src", "bin", "setup", "setup_dev.xlsb"),
        here("src", "bin", "setup", "setup.xlsb"),
        overwrite = TRUE
      )
      cli_alert_success("Successfully replaced the setup main file")
    }
  )
  return(invisible())
}


# update the master setup file
update_master_setup <- function(tag = "mock") {
  # move previous version of my designer
  # replace the dev designer
  switch(
    tag,

    # update the actual mock file
    mock = {
      file_copy(
        here(".mock", "disease_mock.xlsb"),
        here("src", "bin", "master-setup", "disease_setup_dev.xlsb"),
        overwrite = TRUE
      )
      cli_alert_success("Sucesfully replaced the disease dev file")
    },

    # udate the main designer file on the designer folder
    #(the one without the _dev tag)
    main = {
      file_copy(
        here("src", "bin", "master-setup", "disease_setup_dev.xlsb"),
        here("src", "bin", "master-setup", "disease_setup.xlsb"),
        overwrite = TRUE
      )
      cli_alert_success("Sucessfully replace the disease main file")
    }
  )
  return(invisible())
}

# copy the mock file for development
update_designer(tag = "mock")
update_setup(tag = "mock")
update_master_setup(tag = "mock")
