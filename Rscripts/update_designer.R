# update And replace the current by aky one, and update also the config
#for import


# source can take two values: Codes or Github:

# If update is done from github, will just update the linelist_
# designer_aky file with linelist_designer file

# If update is done from the source, will just update the
# linelist_designer file with linelist_designer_aky file

update_designer  <- function(src = "codes") {

    if (src == "codes") {
        #move previous version of my designer
        file.copy(from = "./linelist_designer_aky.xlsb",
                to = "./Rscripts/", overwrite = TRUE)

        # rename it
        file.rename(from = "./Rscripts/linelist_designer_aky.xlsb",
                    to = "./Rscripts/linelist_designer.xlsb")

        # move back and overwrite
        file.copy(from = "./Rscripts/linelist_designer.xlsb",
                to = "./linelist_designer.xlsb", overwrite = TRUE)

        # last designer version
        # rename it
        file.rename(from = "./Rscripts/linelist_designer.xlsb",
                    to = "./Rscripts/linelist_designer_aky.xlsb")
    }

    if (src == "github") {
         #move previous version of my designer
        file.copy(from = "./linelist_designer.xlsb",
                to = "./Rscripts/", overwrite = TRUE)

        # delete the linelist_designer_aky file
        file.remove("./Rscripts/linelist_designer_aky.xlsb")

         # rename the linelist_designer file in the ./Rscript folder
        file.rename(from = "./Rscripts/linelist_designer.xlsb",
                    to = "./Rscripts/linelist_designer_aky.xlsb")

        # move the new designer file and replace the previous one
        file.copy(from = "./Rscripts/linelist_designer_aky.xlsb",
                to = "./linelist_designer_aky.xlsb", overwrite = TRUE)

    }

}

# clear the outpout folder (I have too many outputs)

clear_output  <- function(outdir = "./output") {
    if (dir.exists(outdir)) unlink(outdir, recursive = TRUE)
    dir.create(outdir)
}


# prepare for the a demo

prepare_demo  <- function(fake_dataset = "",
                          fake_geobase = "./input/geobase/default_geobase.xlsx",
                          setup_filename = "setup_measles_SSD_ASH",
                          demo_folder = "./demo") {
    if (dir.exists(demo_folder)) unlink(demo_folder, recursive = TRUE)

    dir.create(demo_folder)

    # add the setup file
    file.copy(from = glue::glue("./input/outbreak-tools-setup/{setup_filename}"), # nolint
              to = glue::glue("{demo_folder}/setup.xlsb"))

    # add the geobase files
    file.copy(from = glue::glue("./input/outbreak-tools-setup/{fake_geobase}"),
              to = glue::glue("{demo_folder}/{fake_geo}.xlsb")

    # add the fake dataset file
    fdata  <- basename(fake_dataset)
    file.copy(from = fake_dataset,
              to = glue::glue("{demo_folder}/{fdata}")
    )

    # add the designer
    file.copy(from = "./linelist_designer.xlsb",
              to = "{demo_folder}/linelist_designer.xlsb")
}


# code for preparing the demo

