source("./Rscripts/update_designer.R")

# update from codes
update_designer(src = "codes")

# update from github
update_designer(src = "github")

# clean the output
clear_output()

# preparing the demo folder

path_fake_dataset  <-  "D:/MSF/OutbreakTools - Library/06 - Pipelines/01 - Test/03 - Fake data set/ll_export_SSD_ASH 1808 20220921-1143.xlsx" # nolint
path_fake_geobase  <- "./input/geobase/OUTBREAK-TOOLS-GEOBASE-SSD-2022-09-07.xlsx" # nolint
setup_filename  <- "setup_measles_SSD_ASH.xlsb"

# preparing the demo folder
prepare_demo(fake_dataset = fake_dataset, fake_geobase = fake_geobase, setup_filename = setup_filename) # nolint
