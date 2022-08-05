Attribute VB_Name = "DesignerConstants"

'This modules describe constants used within the program

'SHEETS  ==========================================================================================================================================================================

'Sheets in the designer file (And not in linelist)
Public Const C_sSheetDesTrans           As String = "designer-translation" 'Sheet for management of designer translation

'MODULES ==========================================================================================================================================================================

Public Const C_sModMain                 As String = "DesignerMain"          'Main module of the designer
Public Const C_sModDesTrans             As String = "DesignerTranslation"   'Translation module for the designer
Public Const C_sModLLExport             As String = "LinelistExport"        'Export module in the linelist
Public Const C_sModLLGeo                As String = "LinelistGeo"           'Geo module in the linelist for geo formula design
Public Const C_sModLinelist             As String = "LinelistEvents"        'Events and buttons in one sheet of type linelist
Public Const C_sModLLMigration          As String = "LinelistMigration"     'Migration of the linelist
Public Const C_sModLLShowHide           As String = "LinelistShowHide"      'ShowHide logic
Public Const C_sModLLConstants          As String = "LinelistConstants"     'Constants of the linelist program
Public Const C_sModEsthConstants        As String = "EstheticConstants"   'Constants for esthetic things in program
Public Const C_sModDesHelpers           As String = "DesignerMainHelpers"
Public Const C_sModHelpers              As String = "Helpers"
Public Const C_sModLLChange             As String = "LinelistChange"     'Change Event linked to one sheet of type linelist
Public Const C_sModLLTrans              As String = "LinelistTranslation"
Public Const C_sModLLDict               As String = "LinelistDictionary"
Public Const C_sModLLAnaChange          As String = "LinelistAnalysisChange"
Public Const C_sModLLCustFunc           As String = "LinelistCustomFunctions"


'CLASSES ==========================================================================================================================================================================
Public Const C_sClaBA                   As String = "BetterArray"

'RANGES AND MESSAGES ==============================================================================================================================================================

'Ranges in the main sheet of the designer

Public Const C_sRngPathDic               As String = "RNG_PathDico"     'Range with path to the dictionary
Public Const C_sRngEdition               As String = "RNG_Edition"      'Range for messages and editions
Public Const C_sRngLLDir                 As String = "RNG_LLDir"        'Range for the linelist Dir (where to save the linelist)
Public Const C_sRngLLName                As String = "RNG_LLName"       'Range for the linelist name in the designer
Public Const C_sRngLangDes               As String = "RNG_LangDesigner" 'Range for the language of the designer
Public Const C_sRngLangSetup             As String = "RNG_LangSetup" 'Range for the language of the setup file
Public Const C_sRngLabLangDes            As String = "RNG_LabLangDesigner" 'Range for label for designer language
Public Const C_sRngLangGeo               As String = "RNG_LangGeo" 'Range for the language of the headings of the geo
Public Const C_sRngPathGeo               As String = "RNG_PathGeo" 'Range to the path to the geo file
Public Const C_sRngLLFormLang            As String = "RNG_LLForm" 'Languages for the forms in the linelist
Public Const C_sRngLLPassword            As String = "RNG_LLPassword" 'Password for debugging the linelist
Public Const C_sRngUpdate                As String = "RNG_Update"

'Messages in the designer ---------------------------------------------------------------------------------------------------------------------------------------------------------

'Inform the designer's user to check for incoherences

Public Const C_sMsgCheckErrorSheet      As String = "MSG_ErrorSheet" 'Something wrong with the one sheet
Public Const C_sMsgCheckCloseLL         As String = "MSG_CloseLL"   'Inform designer user to close the linelist
Public Const C_sMsgCheckLL              As String = "MSG_CheckLL"   'Check the linelist (if something is weird or it will be replaced)
Public Const C_sMsgCheckLLName          As String = "MSG_LLName"    'Check the linelist name
Public Const C_sMsgCheckExits           As String = "MSG_exists"    'Check the linelist exists in the folder
Public Const C_sMsgCheckPathLL          As String = "MSG_PathLL"
Public Const C_sMsgCheckSheetType       As String = "MSG_SheetType" 'Ask the user to check the type of one sheet in the setup file

'Inform the designer user something is going on

Public Const C_sMsgPathLoaded           As String = "MSG_ChemFich" 'Inform designer user the path is loaded
Public Const C_sMsgCancel               As String = "MSG_OpeAnnule" 'Inform designer user about cancelling something
Public Const C_sMsgCleaning             As String = "MSG_NetoPrec" 'Edition Message for cleaning previous data
Public Const C_sMsgDone                 As String = "MSG_Fini" 'End of computation
Public Const C_sMsgReadDic              As String = "MSG_ReadDic" 'Edition message for reading the dictionary
Public Const C_sMsgReadList             As String = "MSG_ReadList" 'Edition message for reading the Lists
Public Const C_sMsgReadExport           As String = "MSG_ReadExport" 'Edition message for reading the export
Public Const C_sMsgBuildLL              As String = "MSG_BuildLL" 'Edition message building the linelist
Public Const C_sMsgOngoing              As String = "MSG_EnCours" 'Message for something ongoing
Public Const C_sMsgTrans                As String = "MSG_Traduit" 'Message for translation done
Public Const C_sMsgInvForm              As String = "MSG_InvalidFormula" 'Message for invalid formula
Public Const C_sMsgPathTooLong          As String = "MSG_PathTooLong" 'Message for path too long
Public Const C_sMsgLLCreated            As String = "MSG_LLCreated" 'Edition message saying Linelist created
Public Const C_sMsgCorrect              As String = "MSG_Correct" 'Edition message saying everything is fine
Public Const C_sMsgSet                  As String = "MSG_Set" 'As the user to set the values for the parameters
Public Const C_sCreatedSheet            As String = "MSG_CreatedSheet" 'Tell the user one sheet has been created

'ux helpers

Public Const C_sMsgChooseFile           As String = "MSG_ChooseFile" 'Pick one file
Public Const C_sMsgChooseDir            As String = "MSG_ChooseDir" 'Pick one directory

'Messsages in the linelist --------------------------------------------------------------------------------------------------------------------------------------------------------

'Linelist incoherences

Public Const C_sMsgWrongCells           As String = "MSG_WrongCells"

'ux helpers
Public Const C_sMsgFileSaved            As String = "MSG_FileSaved"
Public Const C_sMsgNewPass              As String = "MSG_NewPassWord"


'SHAPES ===========================================================================================================================================================================

'TABLES LISTOBJECTS ===============================================================================================================================================================

'Languages tables in the designer translation sheet

Public Const C_sTabLang                  As String = "T_Lang" 'Languages table
Public Const C_sTabTransRange            As String = "T_tradRange" 'Ranges translation table
Public Const C_sTabTransMsg              As String = "T_tradMsg" 'Messages translation table
Public Const C_sTabTransShapes           As String = "T_tradShape" 'Shapes translation table
'Analysis tables
Public Const C_sTabGS                    As String = "Tab_global_summary" 'Global Summary
Public Const C_sTabUA                    As String = "Tab_Univariate_Analysis" 'Univariate Analysis
Public Const C_sTabBA                    As String = "Tab_Bivariate_Analysis" 'Bivariate Analysis
Public Const C_sTabTA                    As String = "Tab_TimeSeries_Analysis" ' Time Series Analysis
Public Const C_sTabSA                    As String = "Tab_SpatialAnalysis" 'Spatial analysis

'STRING CONSTANTS =================================================================================================================================================================

Public Const C_sDesignerPassword         As String = "1234" 'Default password for the designer

'INTEGERS CONSTANTS ===============================================================================================================================================================

'ENUMERATION LISTS ================================================================================================================================================================

'constant linked to the different columns to be translated in the workbook sheets
Public Const sCstColDictionary As String = "Main label|Sub-label|Note|Sheet Name|Main section|Sub-section|Formula|Message"
Public Const sCstColChoices As String = "label_short|label"
Public Const sCstColExport As String = "Label button"
