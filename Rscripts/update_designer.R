# update And replace the current by aky one, and update also the config
#for import


# source can take two values: Codes or Github:

# If update is done from github, will just update the linelist_
# designer_aky file with linelist_designer file

# If update is done from the source, will just update the
# linelist_designer file with linelist_designer_aky file

update_designer  <- function(update_stable = 0) {
    #move previous version of my designer
   file.copy(from = "./linelist_designer_aky.xlsb",
             to = "./Rscripts/", overwrite = TRUE)
    # move back and overwrite
    file.copy(from = "./Rscripts/linelist_designer_aky.xlsb",
             to = "./linelist_designer_dev.xlsb", overwrite = TRUE)
    # update the stable version if needed
    if (update_stable == 1) {
       # previous stable version
         file.copy(from = "./linelist_designer.xlsb",
             to = "./Rscripts/linelist_designer_prev.xlsb", overwrite = TRUE)
       # update the new stable version
         file.copy(from = "./Rscripts/linelist_designer_aky.xlsb",
             to = "./linelist_designer.xlsb", overwrite = TRUE)
    }
    # revert back previous stable designer due to corrupt files.
    if (update_stable == 2) {
         file.copy(from = "./Rscripts/linelist_designer_prev.xlsb",
             to = "./linelist_designer.xlsb", overwrite = TRUE)
         file.copy(from = "./Rscripts/linelist_designer_prev.xlsb",
                 to = "./linelist_designer_dev.xlsb", overwrite = TRUE)
         file.copy(from = "./Rscripts/linelist_designer_prev.xlsb",
                 to = "./linelist_designer_aky.xlsb", overwrite = TRUE)
    }
}

# clear the outpout folder (I have too many outputs)
clear_output  <- function(outdir = "./output") {
    if (dir.exists(outdir)) unlink(outdir, recursive = TRUE)
    dir.create(outdir)
}


# prepare for the a demo

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
    file.copy(from = "./linelist_designer.xlsb",
              to = demo_folder)
}


# create en empty file for creating an interface for each of the class

class_header  <- function(class_name, description = "", module_description = "",
 interface = FALSE){ #nolint
interf  <- ifelse(interface, "I", "")
predeclare_id  <- ifelse(interface, "False", "True")

glue::glue("
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = \"{interf}{class_name}\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = {predeclare_id}
Attribute VB_Exposed = False
Attribute VB_Description = \"{description}\"

'@Folder(\"Dictionary\")
'@ModuleDescription(\"{module_description}\")
'@IgnoreModule

Option Explicit

'Exposed methods
")

}

test_header  <- function(class_name) {

glue::glue(
"
Attribute VB_Name = \"Test{class_name}\"

Option Explicit
Option Private Module

'@TestModule
'@Folder(\"Tests\")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject(\"Rubberduck.AssertClass\")
    Set Fakes = CreateObject(\"Rubberduck.FakesProvider\")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'This method runs before every test in the module..

'@TestInitialize
Private Sub TestInitialize()

End Sub
")
}

create_class  <- function(class_name, description = "", module_description = ""){ #nolint


        class_name_header  <- class_header(class_name,
                                           description = description,
                                           module_description = module_description) #nolint
        class_interface_header  <- class_header(class_name, interface = TRUE,
                                                description = description,
                                                module_description = glue::glue("Interface of {module_description}")) #nolint

        class_test_header  <- test_header(class_name)
        #create the class
        cat(class_name_header,
            file = glue::glue("./src/vba-files/Class/{class_name}.cls"))
        #create the interface of the class
        cat(class_interface_header,
        file = glue::glue("./src/vba-files/Class/I{class_name}.cls"))
        # Add Test for the class
        cat(class_test_header,
        file = glue::glue("./src/vba-files/Module/Test{class_name}.bas"))
}