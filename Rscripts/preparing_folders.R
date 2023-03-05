#
source("./Rscripts/update_designer.R")

# update from codes
update_designer()

# update by adding the stable
#update_designer(update_stable = 1)

# clean the output
clear_output()

# preparing the demo folder

path_fake_dataset  <-  "D:/MSF/OutbreakTools - Library/03 - Test/TEST MARINE/20230131/LL-test-Marine_with fake data.xlsb" # nolint
path_fake_geobase  <- "D:/MSF/OutbreakTools - Library/03 - Test/TEST MARINE/20230131/geobase_obt_yem_20230112.xlsx" # nolint
setup_filename  <- "D:/MSF/OutbreakTools - Library/03 - Test/TEST MARINE/20230131/setup_measles_Yemen_20230201.xlsb" # nolint
demo_folder  <- "./demo"

# preparing the demo folder
clear_output("./demo")

prepare_demo(fake_dataset = path_fake_dataset,
             fake_geobase = path_fake_geobase,
             setup_filename = setup_filename) # nolint


# create a class