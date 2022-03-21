Attribute VB_Name = "LinelistConstants"

Option Explicit

Public Const C_sSheetGeo                As String = "GEO"                                'Sheet for storing the geo data // Feuille des donnees geo
Public Const C_sSheetAdmin              As String = "Admin"                              'Name of the sheet for admin data (metadata on a linelist)
Public Const C_sSheetPassword           As String = "Password"                           'sheet for password management
Public Const C_sSheetFormulas           As String = "ControleFormule"                    'Sheet for formula management (mainly in the validation)
Public Const C_sSheetLLTranslation      As String = "linelist-translation"

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
Public Const C_sDictHeaderSheetType     As String = "sheet type"                         'Type of the sheet to add to the linelist
Public Const C_sDictHeaderMainSec       As String = "main section"                       'Main Section of the variable (to show in heading)
Public Const C_sDictHeaderSubSec        As String = "sub section"                        'Sub Section of the variable to show under the headings
Public Const C_sDictHeaderStatus        As String = "status"                             'Status of the variable
Public Const C_sDictHeaderId            As String = "personal identifier"                'Is the variable a personal identifier?
Public Const C_sDictHeaderType          As String = "type"                               'Type of the variable
Public Const C_sDictHeaderControl       As String = "control"                            'Control for the variable (one of the differents types of control)
Public Const C_sDictHeaderFormula       As String = "formula"                            'Formulas for the variable
Public Const C_sDictHeaderChoices       As String = "choices"                            'Different types of choices for one variable
Public Const C_sDictHeaderUnique        As String = "unique"                             '
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

'Added headers to the dictionnary
Public Const C_sDictHeaderIndex         As String = "column index"
Public Const C_sDictHeaderVisibility    As String = "visibility"

'Sheet Types ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictSheetTypeLL         As String = "linelist"
Public Const C_sDictSheetTypeAdm        As String = "admin"

'Status ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictStatusMan As String = "mandatory"
Public Const C_sDictStatusOpt As String = "optional"
Public Const C_sDictStatusHid As String = "hidden"

'Control --------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictControlHf As String = "hf"
Public Const C_sDictControlForm As String = "formula"
Public Const C_sDictControlChoice As String = "choices"
Public Const C_sDictControlGeo As String = "geo"
Public Const C_sDictControlCustom As String = "custom"
Public Const C_sDictControlTitle As String = "title"

'YesNo ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictYes As String = "yes"
Public Const C_sDictNo As String = "no"

'Alert ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Const C_sDictAlertWar As String = "warning"
Public Const C_sDictAlertErr As String = "error"

'Data Types
Public Const C_sDictTypeInt As String = "integer"
Public Const C_sDictTypeText As String = "text"
Public Const C_sDictTypeDate As String = "date"
Public Const C_sDictTypeDec As String = "decimal"


'CHOICES PARAMETERS ==========================================================================

Public Const C_sChoiHeaderLab As String = "label" 'Choice label name headers
Public Const C_sChoiHeaderList As String = "list name" 'Choice list headers

'EXPORTS PARAMETERS ===========================================================================
Public Const C_sExportActive As String = "active"
Public Const C_sExportInactive As String = "inactive"


'FORMS ============================================================================================================================================================================

Public Const C_sFormExport               As String = "F_Export"                          'Export Frame
Public Const C_sFormGeo                  As String = "F_Geo"                             'Geo Frame
Public Const C_sFormShowHide             As String = "F_NomVisible"                      'ShowHide Frame

'TABLES LISTOBJECTS ===============================================================================================================================================================

'Admin levels tables in the Geo Sheet

Public Const C_sTabADM1                  As String = "T_ADM1"                              'ADM1 Table name
Public Const C_sTabADM2                  As String = "T_ADM2"                              'ADM2 Table name
Public Const C_sTabADM3                  As String = "T_ADM3"                              'ADM3 Table name
Public Const C_sTabADM4                  As String = "T_ADM4"                              'ADM4 Table name
Public Const C_sTabHF                    As String = "T_HF"                                'Health Facility Table
Public Const C_sTabNames                 As String = "T_NAMES"
Public Const C_sTabHistoGeo              As String = "T_HistoGeo"                          'Historic data for the geo
Public Const C_sTabHistoHF               As String = "T_HistoHF"                          'Historic data for the Health Facility

'Formulas and functions tables

Public Const C_sTabExcelFunctions        As String = "T_XlsFonctions"                      'Excel functions to keep in formulas
Public Const C_sTabASCII                 As String = "T_ascii"                             'Ascii characters table


'PROGRAM NAMES ==================================================================================
'Program names are used for setting programs to buttons added in the linelist

Public Const C_sCmdShowHideName = "ClicCmdShowHide" 'ShowHideCommand
Public Const C_sCmdAddRowsName = "ClicCmdAddRows"
Public Const C_sCmdShowGeoApp = "ClicCmdGeoApp"
Public Const C_sCmdExportMigration = "ClicExportMigration"
Public Const C_sCmdImportMigration = "ClicImportMigration"
Public Const C_sCmdExport = "ClicCmdExport"

'TABLES LISTOBJECTS ===============================================================================================================================================================

Public Const C_sTabkeys = "T_Keys"


'RANGES, MESSAGES AND SHAPES =======================================================================================================================================================

'Shapes----------------------------------------------------------
Public Const C_sShpShowHide = "SHP_ShowHide"
Public Const C_sShpAddRows = "SHP_Add200L"
Public Const C_sShpGeo = "SHP_GeoApps"
Public Const C_sShpExpMigration = "SHP_ExportMig"
Public Const C_sShpImpMigration = "SHP_ImportMig"
Public Const C_sShpExport = "SHP_Export"

'Ranges in the linelist sheet
Public Const C_sRngPublickey             As String = "RNG_PublicKey"                     'Name of the range for publickey
Public Const C_sRngPrivatekey            As String = "RNG_PrivateKey"                    'Name of the range for the private key



'STRING CONSTANTS =================================================================================================================================================================
Public Const C_sLLPassword               As String = "1234"                                'Default password for the linelist file (if no one is set)                                'Default password for the designer
Public Const C_sAdmName                  As String = "adm"   'Name of administrative levels names
Public Const C_sYes                     As String = "yes"
Public Const C_sNo                      As String = "no"

'INTEGERS CONSTANTS ===============================================================================================================================================================
Public Const C_iNbLinesLLData           As Integer = 200                                    'Number of linest to add by default


'Startlines for data in various source files

Public Enum C_StartLines
    C_eStartLinesDictHeaders = 2                                                             'Starting lines for dictionary headers
    C_eStartLinesDictData = 3                                                                'Starting lines for dictionary data
    C_eStartLinesLLMainSec = 3                                                                'Starting lines for first title of the linelist
    C_eStartLinesLLSubSec = 4                                                                'Starting lines for second title of the linelist
    C_eStartLinesLLData = 5                                                                  'Starting lines for the linelist data
    C_eStartLinesExportTitle = 1                                                             'Starting lines for export titles
    C_eStartLinesAdmData = 15                                                                'Starting lines for a Adm data
    C_eStartLinesExportSource = 5                                                            'Starting lines for export sources
    C_eStartLinesChoicesHeaders = 1                                                          'Starting lines of the choices Headers
    C_eStartLinesChoicesData = 2                                                             'Starting lines of the choices Data
    C_eStartLinesExportData = 2                                                              'Starting lines of the export data
End Enum

