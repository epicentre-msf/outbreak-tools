
update_designer  <- function(update_status = 0) {

   #move previous version of my designer
   # replace the dev designer
   if (update_status == 0){
        file.copy(from = "./src/bin/designer_aky.xlsb",
             to = "./Rscripts/", overwrite = TRUE)
        # move back and overwrite
        file.copy(from = "./Rscripts/designer_aky.xlsb",
                to = "./src/bin/designer_dev.xlsb", overwrite = TRUE)
   }
    # update the stable version if needed
    if (update_status == 1) {
       # previous stable version
         file.copy(from = "./designer.xlsb",
             to = "./Rscripts/designer_prev.xlsb", overwrite = TRUE)
       # update the new stable version
         file.copy(from = "./Rscripts/designer_aky.xlsb",
             to = "./designer.xlsb", overwrite = TRUE)
    }

    # revert back previous stable designer due to corrupt files.
    if (update_status == 2) {
         file.copy(from = "./Rscripts/designer_prev.xlsb",
             to = "./designer.xlsb", overwrite = TRUE)
         file.copy(from = "./Rscripts/designer_prev.xlsb",
                 to = "./src/bin/designer_dev.xlsb", overwrite = TRUE)
         file.copy(from = "./Rscripts/designer_prev.xlsb",
                 to = "./src/bin/designer_aky.xlsb", overwrite = TRUE)
    }

    #copy the file from the mock designer
    if (update_status == 3) {
        file.copy(from = "./src/.mock/designer_mock.xlsb",
                to = "./src/bin/designer_aky.xlsb", overwrite = TRUE)
    }
}

#update the dev file
update_designer(update_status = 0) #nolint
#update the file on root
update_designer(update_status = 1) #nolint
#copy the mock file for development
update_designer(update_status = 3)
