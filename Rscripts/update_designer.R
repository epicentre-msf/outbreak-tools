# update And replace the current by aky one, and update also the config for import

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
            to = "./Rscripts/linelist_designer_aky.xlsb"
)
