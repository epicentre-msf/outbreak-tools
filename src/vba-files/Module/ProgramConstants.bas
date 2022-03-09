Attribute VB_Name = "ProgramConstants"
'This modules describe constants used within the program

'SHEETS  ==========================================================================================================================================================================

'Sheets in the designer file

Public Const C_sSheetGeo                As String = "Geo"                                'Sheet for storing the geo data // Feuille des donnees geo
Public Const C_sSheetAdmin              As String = "Admin"                              'Name of the sheet for admin data (metadata on a linelist)
Public Const C_sSheetDesTrans           As String = "designer-translation"               'Sheet for management of designer translation
Public Const C_sSheetPassWord           As String = "Password"                           'sheet for password management
Public Const C_sSheetFormulas           As String = "ControleFormule"                    'Sheet for formula management (mainly in the validation)

'Sheets in the setup file: The all starts with Param

Public Const C_sParamSheetDict          As String = "Dictionary"                         'Dictionnary Sheet in the setup file
Public Const C_sParamSheetExport        As String = "Exports"                            'Sheet with configurations for the export  in the setup filefile
Public Const C_sParamSheetChoices       As String = "Choices"                            'Sheet with configurations for the choices in the setup file
Public Const C_sParamSheetAdminName     As String = "Admin"                              'Name of a sheet of type admin for metadata on linelist

'DICTIONARY PARAMETERS ============================================================================================================================================================

'Headers of the dictionnary in the setup file Headers are fixed -------------------------------------------------------------------------------------------------------------------

Public Const C_sDictHeaderVarName       As String = "variable name"                      'Variable Name (unique identifier of a variable in lowercase without spaces)
Public Const C_sDictHeaderMainLab       As String = "main label"                         'Variable Label
Public Const C_sDictHeaderSubLab        As String = "sub label"                          'Variable Sub Label (sub label of the variable name in gray)
Public Const C_sDictHeaderNote          As String = "note"                               'Notes to show for the variable
Public Const C_sDictHeaderSheetName     As String = "sheet name"                         'Name of the sheet to add to the linelist
Public Const C_sDictHeaderSheetType     As String = "Sheet Type"                         'Type of the sheet to add to the linelist
Public Const C_sDictHeaderMainSec       As String = "main section"                       'Main Section of the variable (to show in heading)
Public Const C_sDictHeaderSubSec        As String = "sub section"                        'Sub Section of the variable to show under the headings
Public Const C_sDictHeaderStatus        As String = "status"                             'Status of the variable
Public Const C_sDictHeaderId            As String = "personal identifier"                'Is the variable a personal identifier?
Public Const C_sDictHeaderType          As String = "type"                               'Type of the variable ()
Public Const C_sDictHeaderControl       As String = "control"
Public Const C_sDictHeaderFormula       As String = "formula"
Public Const C_sDictHeaderChoices       As String = "choices"
Public Const C_sDictHeaderUnique        As String = "unique"
Public Const C_sDictHeaderSource        As String = "source"
Public Const C_sDictHeaderHxl           As String = "hxl"
Public Const C_sDictHeaderExport1       As String = "export 1"
Public Const C_sDictHeaderExport2       As String = "export 2"
Public Const C_sDictHeaderExport3       As String = "export 3"
Public Const C_sDictHeaderExport4       As String = "export 4"
Public Const C_sDictHeaderExport5       As String = "export 5"
Public Const C_sDictHeaderMin           As String = "min"
Public Const C_sDictHeaderMax           As String = "max"
Public Const C_sDictHeaderAlert         As String = "alert"
Public Const C_sDictHeaderMessage       As String = "message"
Public Const C_sDictHeaderBranchLogic   As String = "branching logic"
Public Const C_sDictHeaderCondFormat    As String = "conditional formating"

'Sheet Types ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictSheetTypeLL As String = "linelist"
Public Const C_sDictSheetTypeAdm As String = "admin"

'Status ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictStatusMan As String = "mandatory"
Public Const C_sDictStatusOpt As String = "optional"
Public Const C_sDictStatusHid As String = "hidden"

'Control --------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictControlHf As String = "hf"
Public Const C_sDictControlForm As String = "formula"
Public Const C_sDictControlCho As String = "choices"
Public Const C_sDictControlGeo As String = "geo"

'YesNo ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictYes As String = "yes"
Public Const C_sDictNo As String = "no"

'Alert ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictAlertWar As String = "warning"
Public Const C_sDictAlertErr As String = "error"

'MODULES ==========================================================================================================================================================================
Public Const C_sModMain                 As String = "M_Main"                             'Main module of the designer
Public Const C_sModTrans                As String = "M_traduction"                       'Translation module for the designer
Public Const C_sModExport               As String = "M_Export"                           'Export module in the linelist
Public Const C_sModGeo                  As String = "M_Geo"                              'Geo module in the linelist for geo formula design
Public Const C_sModLinelist             As String = "M_linelist"                         'Events and buttons in one sheet of type linelist
Public Const C_sModMigration            As String = "M_Migration"                        'Migration of the linelist
Public Const C_sModShowHide             As String = "M_NomVisible"                       'ShowHide logic
Public Const C_sModValidation           As String = "M_validation"                       'Validation of the formulas
Public Const C_sModConstants            As String = "ProgramConstants"                   'Constants of the program
Public Const C_sLLChange                As String = "linelist_sheet_change"              'Change Event linked to one sheet of type linelist

'FORMS ============================================================================================================================================================================

Public Const C_sFormExport               As String = "F_Export"                          'Export Frame
Public Const C_sFormGeo                  As String = "F_Geo"                             'Geo Frame
Public Const C_sFormShowHide             As String = "F_NomVisible"                      'ShowHide Frame

'RANGES AND MESSAGES ==============================================================================================================================================================

'Ranges in the main sheet of the designer
Public Const C_sRngPathDic               As String = "RNG_PathDico"                      'Range with path to the dictionary
Public Const C_sRngEdition               As String = "RNG_Edition"                       'Range for messages and editions
Public Const C_sRngLLDir                 As String = "RNG_LLDir"                         'Range for the linelist Dir (where to save the linelist)
Public Const C_sRngLLName                As String = "RNG_LLName"                        'Range for the linelist name in the designer
Public Const C_sRngLangDes               As String = "RNG_LangDesigner"                  'Range for the language of the designer
Public Const C_sRngLangSetup             As String = "RNG_LangSetup"                     'Range for the language of the setup file
Public Const C_sRngLabLangDes            As String = "RNG_LabLangDesigner"               'Range for label for designer language
Public Const C_sRngLangGeo               As String = "RNG_LangGeo"                       'Range for the language of the headings of the geo
Public Const C_sRngPathGeo               As String = "RNG_PathGeo"                       'Range to the path to the geo file

'Ranges in the linelist sheet
Public Const C_sRngPublickey             As String = "RNG_PublicKey"                     'Name of the range for publickey
Public Const C_sRngPrivatekey            As String = "RNG_PrivateKey"                    'Name of the range for the private key

'Messages in the designer ---------------------------------------------------------------------------------------------------------------------------------------------------------

'Inform the designer's user to check for incoherences
Public Const C_sMsgCheckErrorSheet      As String = "MSG_ErrorSheet"                     'Something wrong with the one sheet
Public Const C_sMsgCheckCloseLL         As String = "MSG_CloseLL"                        'Inform designer user to close the linelist
Public Const C_sMsgCheckLL              As String = "MSG_CheckLL"                        'Check the linelist (if something is weird or it will be replaced)
Public Const C_sMsgCheckLLName          As String = "MSG_LLName"                         'Check the linelist name
Public Const C_sMsgCheckExits           As String = "MSG_exists"                         'Check the linelist exists in the folder
Public Const C_sMsgCheckPathLL          As String = "MSG_PathLL"
Public Const C_sMsgCheckSheetType       As String = "MSG_SheetType"                      'Ask the user to check the type of one sheet in the setup file

'Inform the designer user something is going on
Public Const C_sMsgPathLoaded           As String = "MSG_ChemFich"                       'Inform designer user the path is loaded
Public Const C_sMsgCancel               As String = "MSG_OpeAnnule"                      'Inform designer user about cancelling something
Public Const C_sMsgCleaning             As String = "MSG_NetoPrec"                       'Edition Message for cleaning previous data
Public Const C_sMsgDone                 As String = "MSG_Fini"                           'End of computation
Public Const C_sMsgReadDic              As String = "MSG_ReadDic"                        'Edition message for reading the dictionary
Public Const C_sMsgReadList             As String = "MSG_ReadList"                       'Edition message for reading the Lists
Public Const C_sMsgReadExport           As String = "MSG_ReadExport"                     'Edition message for reading the export
Public Const C_sMsgBuildLL              As String = "MSG_BuildLL"                        'Edition message building the linelist
Public Const C_sMsgCreatedLL            As String = "MSG_LLCreated"                      'Edition message to inform linelist is created
Public Const C_sMsgOngoing              As String = "MSG_EnCours"                        'Message for something ongoing
Public Const C_sMsgTrans                As String = "MSG_Traduit"                        'Message for translation done
Public Const C_sMsgInvForm              As String = "MSG_InvalidFormula"                 'Message for invalid formula
Public Const C_sMsgPathTooLong          As String = "MSG_PathTooLong"                    'Message for path too long
Public Const C_sMsgLLCreated            As String = "MSG_LLCreated"                      'Edition message saying Linelist created
Public Const C_sMsgCorrect              As String = "MSG_Correct"                        'Edition message saying everything is fine
Public Const C_sMsgSet                  As String = "MSG_Set"                            'As the user to set the values for the parameters
Public Const C_sCreatedSheet            As String = "MSG_CreatedSheet"                   'Tell the user one sheet has been created


'ux helpers
Public Const C_sMsgChooseFile           As String = "MSG_ChooseFile"                      'Pick one file
Public Const C_sMsgChooseDir            As String = "MSG_ChooseDir"                       'Pick one directory

'Messsages in the linelist --------------------------------------------------------------------------------------------------------------------------------------------------------

'Linelist incoherences
Public Const C_sMsgWrongCells           As String = "MSG_WrongCells"

'ux helpers
Public Const C_sMsgFileSaved            As String = "MSG_FileSaved"
Public Const C_sMsgNewPass              As String = "MSG_NewPassWord"

'TABLES LISTOBJECTS ===============================================================================================================================================================

'Admin levels tables in the Geo Sheet
Public Const C_sTabADM1                  As String = "T_ADM1"                              'ADM1 Table name
Public Const C_sTabADM2                  As String = "T_ADM2"                              'ADM2 Table name
Public Const C_sTabADM3                  As String = "T_ADM3"                              'ADM3 Table name
Public Const C_sTabADM4                  As String = "T_ADM4"                              'ADM4 Table name
Public Const C_sTabHF                    As String = "T_HF"                                'Health Facility Table
Public Const C_sTabNames                 As String = "T_NAMES"
Public Const C_sTabHistoGeo              As String = "T_HistoGeo"                          'Historic data for the geo
Public Const C_sTabHistoHF               As String = "T_Histo_HF"                          'Historic data for the Health Facility

'Languages tables in the designer translation sheet
Public Const C_sTabLang                  As String = "T_Lang"                              'Languages table
Public Const C_sTabTransRange            As String = "T_tradRange"                         'Ranges translation table
Public Const C_sTabTransMsg              As String = "T_tradMsg"                           'Messages translation table
Public Const C_sTabTransShapes           As String = "T_tradShape"                         'Shapes translation table

'Formulas and functions tables
Public Const C_sTabExcelFunctions        As String = "T_XlsFonctions"                      'Excel functions to keep in formulas
Public Const C_sTabASCII                 As String = "T_ascii"                             'Ascii characters table

'STRING CONSTANTS =================================================================================================================================================================
Public Const C_sLinelistPassword         As String = "1234"                                'Default password for the linelist file (if no one is set)
Public Const C_sDesignerPassword         As String = "1234"                                'Default password for the designer

'INTEGERS CONSTANTS ===============================================================================================================================================================
Public Const C_iLLButtonHeight As Integer = 25 'Row height of the first two rows on one sheet of type Linelist for buttons




'ENUMERATION LISTS ================================================================================================================================================================

'Startlines for data in various source files
Public Enum C_StartLines
    C_eStartLinesDictHeaders = 2                                                               'Starting lines for dictionary headers
    C_eStartLinesDictData = 3                                                                  'Starting lines for dictionary data
    C_eStartLinesLLTitle1 = 3                                                                  'Starting lines for first title of the linelist
    C_eStartLinesLLTitle2 = 4                                                                  'Starting lines for second title of the linelist
    C_eStartLinesLLData = 5                                                                    'Starting lines for the linelist data
    C_eStartLinesExportTitle = 1                                                               'Starting lines for export titles
    C_eStartLinesExportSource = 5                                                              'Starting lines for export sources
    C_eStartLinesChoicesHeaders = 1
    C_eStartLinesChoicesData = 2
    C_eStartLinesExportData = 2
End Enum



