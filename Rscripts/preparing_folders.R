# Prepare demo and output folders

# clear the outpout folder (I have too many outputs)
clear_output  <- function(outdir = "./output") {
    if (dir.exists(outdir)) unlink(outdir, recursive = TRUE)
    dir.create(outdir)
}

# code for preparing the demo
prepare_demo  <- function(fake_dataset = "",
                          fake_geobase = "./input/geobase/default_geobase.xlsx",
                          setup_filename = "setup_measles_SSD_ASH",
                          demo_folder = "./demo") {
    if (dir.exists(demo_folder)) unlink(demo_folder, recursive = TRUE)

    dir.create(demo_folder)

    # add the setup file
    file.copy(from = setup_filename, # nolint
              to = demo_folder)

    fake_geo  <- basename(fake_geobase)

    # add the geobase files
    file.copy(from = fake_geobase,
              to = demo_folder)

    # add the fake dataset file

    fdata  <- basename(fake_dataset)

    file.copy(from = fake_dataset,
              to = demo_folder
    )

    # add the designer
    file.copy(from = "./designer.xlsb",
              to = demo_folder)
}

# path_fake_dataset  <-  "D:/MSF/OutbreakTools - Library/03 - Test/TEST MARINE/20230131/LL-test-Marine_with fake data.xlsb" # nolint
# path_fake_geobase  <- "D:/MSF/OutbreakTools - Library/03 - Test/TEST MARINE/20230131/geobase_obt_yem_20230112.xlsx" # nolint
# setup_filename  <- "D:/MSF/OutbreakTools - Library/03 - Test/TEST MARINE/20230131/setup_measles_Yemen_20230201.xlsb" # nolint
