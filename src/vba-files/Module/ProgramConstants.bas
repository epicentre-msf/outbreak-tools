Attribute VB_Name = "ProgramConstants"
'This modules describe constants used within the program

'SHEETS  =============================================================

'Sheets in the designer file

Public Const C_SheetGeo = "Geo"                  'Sheet for storing the geo data // Feuille des donnees geo
Public Const C_SheetAdmin = "Admin"              'Name of the sheet for admin data (metadata on a linelist)
Public Const C_SheetDesTrans = "designer-translation" 'Sheet for management of designer translation
Public Const C_SheetPassWord = "Password"        'sheet for password management
Public Const C_SheetFormulas = "ControleFormule" 'Sheet for formula management (mainly in the validation)

'Sheets in the setup file: The all starts with Param

Public Const C_ParamSheetDict = "Dictionary"     'Dictionnary Sheet in the setup file
Public Const C_ParamSheetExport = "Exports"      'Sheet with configurations for the export  in the setup filefile
Public Const C_ParamSheetChoices = "Choices"     'Sheet with configurations for the choices in the setup file


'MODULES =======================================================================
Public Const C_ModMain = "M_Main"                'Main module of the designer
Public Const C_ModTrans = "M_traduction"         'Translation module for the designer
Public Const C_ModExport = "M_Export"            'Export module in the linelist
Public Const C_ModGeo = "M_Geo"                  'Geo module in the linelist for geo formula design
Public Const C_ModLinelist = "M_linelist"        'Events and buttons in one sheet of type linelist
Public Const C_ModMigration = "M_Migration"      'Migration of the linelist
Public Const C_ModShowHide = "M_NomVisible"      'ShowHide logic
Public Const C_ModValidation = "M_validation"    'Validation of the formulas
Public Const C_ModConstants = "ProgramConstants" 'Constants of the program
Public Const C_LLChange = "linelist_sheet_change" 'Change Event linked to one sheet of type linelist

'FORMS =========================================================================

Public Const C_FormExport = "F_Export"
Public Const C_FormGeo = "F_Geo"
Public Const C_FormShowHide = "F_NomVisible"

'RANGES AND MESSAGES ===========================================================

'Ranges in the main sheet of the designer
Public Const C_RNG_PathDic = "RNG_PathDico"      'Range with path to the dictionary
Public Const C_RNG_Edition = "RNG_Edition"       'Range for messages and editions
Public Const C_RNG_LLDir = "RNG_LLDir"           'Range for the linelist Dir (where to save the linelist)
Public Const C_RNG_LLName = "RNG_LLName"         'Range for the linelist name in the designer
Public Const C_RNG_LangDes = "RNG_LangDesigner"  'Range for the language of the designer
Public Const C_RNG_LangSetup = "RNG_LangSetup"   'Range for the language of the setup file
Public Const C_RNG_LabLangDes = "RNG_LabLangDesigner" 'Range for label for designer language
Public Const C_RNGLangGeo = "RNG_LangGeo"        'Range for the language of the headings of the geo

'Ranges in the linelist sheet
Public Const C_RNG_Publickey = "RNG_PublicKey"   'Name of the range for publickey
Public Const C_RNG_Privatekey = "RNG_PrivateKey" 'Name of the range for the private key

'Messages in the designer ------------------------------------------------------

'Inform the designer's user to check for incoherences
Public Const C_MSG_CheckErrorSheet = "MSG_ErrorSheet" 'Something wrong with the one sheet
Public Const C_MSG_CheckCloseLL = "MSG_CloseLL"  'Inform designer user to close the linelist
Public Const C_MSG_CheckLL = "MSG_CheckLL"       'Check the linelist (if something is weird or it will be replaced)
Public Const C_MSG_CheckLLName = "MSG_LLName"    'Check the linelist name
Public Const C_MSG_CheckExits = "MSG_exists"     'Check the linelist exists in the folder
Public Const C_MSG_CheckPathLL = "MSG_PathLL"

'Inform the designer user something is going on
Public Const C_MSG_PathLoaded = "MSG_ChemFich"   'Inform designer user the path is loaded
Public Const C_MSG_Cancel = "MSG_OpeAnnule"      'Inform designer user about cancelling something
Public Const C_MSG_Cleaning = "MSG_NetoPrec"     'Edition Message for cleaning previous data
Public Const C_MSG_Done = "MSG_Fini"             'End of computation
Public Const C_MSG_ReadDic = "MSG_ReadDic"       'Edition message for reading the dictionary
Public Const C_MSG_ReadList = "MSG_ReadList"     'Edition message for reading the Lists
Public Const C_MSG_ReadExport = "MSG_ReadExport" 'Edition message for reading the export
Public Const C_MSG_BuildLL = "MSG_BuildLL"       'Edition message building the linelist
Public Const C_MSG_CreatedLL = "MSG_LLCreated"   'Edition message to inform linelist is created
Public Const C_MSG_Ongoing = "MSG_EnCours"       'Message for something ongoing
Public Const C_MSG_Trans = "MSG_Traduit"         'Message for translation done
Public Const C_MSG_InvForm = "MSG_InvalidFormula" 'Message for invalid formula
Public Const C_MSG_PathTooLong = "MSG_PathTooLong" 'Message for path too long
Public Const C_MSG_LLCreated = "MSG_LLCreated"   'Edition message saying Linelist created
Public Const C_MSG_Correct = "MSG_Correct"       'Edition message saying everything is fine

'ux helpers
Public Const C_MSG_ChooseFile = "MSG_ChooseFile" 'Pick one file
Public Const C_MSG_ChooseDir = "MSG_ChooseDir"   'Pick one directory

'Messsages in the linelist -----------------------------------------------------

'Linelist incoherences
Public Const C_MSG_WrongCells = "MSG_WrongCells"

'ux helpers
Public Const C_MSG_FileSaved = "MSG_FileSaved"
Public Const C_MSG_NewPass = "MSG_NewPassWord"

'TABLES LISTOBJECTS ============================================================

'Admin levels tables in the Geo Sheet
Public Const C_TAB_ADM1 = "T_ADM1"               'ADM1 Table
Public Const C_TAB_ADM2 = "T_ADM2"               'ADM2 Table
Public Const C_TAB_ADM3 = "T_ADM3"               'ADM3 Table
Public Const C_TAB_ADM4 = "T_ADM4"               'ADM4 Table
Public Const C_TAB_HF = "T_HF"                   'Health Facility Table
Public Const C_TAB_Names = "T_NAMES"
Public Const C_TAB_HistoGeo = "T_HistoGeo"       'Historic data for the geo
Public Const C_TAB_HistoHF = "T_Histo_HF"        'Historic data for the Health Facility

'Languages tables in the designer translation sheet
Public Const C_TAB_Lang = "T_Lang"               'Languages table
Public Const C_TAB_TransRange = "T_tradRange"    'Ranges translation table
Public Const C_TAB_TransMsg = "T_tradMsg"        'Messages translation table
Public Const C_TAB_TransShapes = "T_tradShape"   'Shapes translation table

'Formulas and functions tables
Public Const C_TAB_ExcelFunctions = "T_XlsFonctions" 'Excel functions to keep in formulas
Public Const C_TAB_ASCII = "T_ascii"             'Ascii characters table

'STRING CONSTANTS ==============================================================
Public Const C_LinelistPassword = "1234"         'Default password for the linelist file (if no one is set)
Public Const C_DesignerPassword = "1234"         'Default password for the designer


'INTEGERS CONSTANTS ============================================================





'ENUMERATION LISTS =============================================================

'Startlines for data in various source files
Public Enum StartLines
    StartLinesDict = 3                           'Starting lines for dictionary
    StartLinesLLTitle1 = 3                       'Starting lines for first title of the linelist
    StartLinesLLTitle2 = 4                       'Starting lines for second title of the linelist
    StartLinesLLData = 5                         'Starting lines for the linelist data
    StartLinesExportTitle = 1                    'Starting lines for export titles
    StartLinesExportSource = 5                   'Starting lines for export sources
End Enum

