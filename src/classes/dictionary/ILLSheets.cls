VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Dictionary")
'@ModuleDescription("Interface of Linelist sheets")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'@section Enumerations
'===============================================================================

'@description Bound selectors used by DataBounds. The original spelling of
'@description RowSart is preserved for backwards compatibility.
Public Enum SheetBound
    RowSart = 1
    RowEnd = 2
    ColStart = 3
    ColEnd = 4
End Enum

'@description Metadata selectors consumed by SheetInfo.
Public Enum SheetInfoType
    SheetInfoSheetType = 1
    SheetInfoSheetTable = 2
End Enum

'@section Properties
'===============================================================================

'@prop-title Dictionary accessor
'@return ILLdictionary backing store for the sheets helper.
Public Property Get Dictionary() As ILLdictionary: End Property

'@section Queries
'===============================================================================

'@sub-title Retrieve a specific metadata entry for the supplied sheet
'@param sheetName String. Sheet identifier.
'@param info Byte. Selector from SheetInfoType.
'@return String metadata value.
Public Function SheetInfo(ByVal sheetName As String, Optional ByVal info As Byte = SheetInfoType.SheetInfoSheetType) As String: End Function

'@jump:Contains
'@sub-title Determine whether the dictionary contains the supplied sheet
'@param sheetName String. Sheet identifier.
'@return True when the sheet exists.
Public Function Contains(ByVal sheetName As String) As Boolean: End Function

'@sub-title Return positional bounds for the sheet layout
'@param sheetName String. Sheet identifier.
'@param Bound Byte. Selector from SheetBound.
'@return Long bound value.
Public Function DataBounds(ByVal sheetName As String, Optional ByVal Bound As Byte = SheetBound.RowSart) As Long: End Function

'@sub-title Test whether a sheet exposes the requested control type
'@param sheetName String. Sheet identifier.
'@param control Optional String. Control type to search for. Defaults to "list_auto".
'@param colName Optional String. Column that contains control metadata. Defaults to "control".
'@return True when at least one variable matches the control.
Public Function ContainsControl(ByVal sheetName As String, _
                                Optional ByVal control As String = "list_auto", _
                                Optional ByVal colName As String = "control") As Boolean: End Function

'@sub-title Build a cell address pointing to the variable location
'@param varName String. Variable identifier.
'@param onSheet Optional String. Target sheet to suppress sheet prefix when matching.
'@return String Excel address for the variable.
Public Function VariableAddress(ByVal varName As String, Optional ByVal onSheet As String = vbNullString) As String: End Function

'@sub-title Retrieve the worksheet row containing the sheet metadata row
'@param sheetName String. Sheet identifier.
'@return Long worksheet row index (0 when not found).
Public Function RowIndex(ByVal sheetName As String) As Long: End Function

'@sub-title Count the variables assigned to the sheet
'@param sheetName String. Sheet identifier.
'@return Long number of variables.
Public Function NumberOfVars(ByVal sheetName As String) As Long: End Function

'Reset internal caches after dictionary structure changes.
Public Sub InvalidateCaches(): End Sub
