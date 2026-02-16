VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLChoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Choice class"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of Choice class")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ILLChoices
'@description
'Defines the public contract for the LLChoices class, exposing
'management of choice lists used for dropdowns and cascading selectors.
'Supports adding, removing, sorting, translating, importing, and
'exporting choices along with their long and short label categories.


'@section Elements
'===============================================================================
'@description Properties exposing the underlying DataSheet and worksheet.

'@jump:choice-data
'@prop-title Associated DataSheet object
'@details
'Returns the IDataSheet wrapping the choice worksheet data.
'@return IDataSheet. The backing DataSheet.
Public Property Get ChoiceData() As IDataSheet: End Property

'@jump:wksh
'@prop-title Worksheet hosting the choice lists
'@details
'Returns the worksheet that stores the choice data.
'@return Worksheet. The host worksheet.
Public Property Get Wksh() As Worksheet: End Property


'@section Choices
'===============================================================================
'@description Operations for managing individual choice entries.

'@jump:name
'@prop-title Name of the choice object
'@details
'Returns the assigned name, or the worksheet name when none was provided.
'@return String. The choice object name.
Public Property Get Name() As String: End Property

'@jump:choice-exists
'@prop-title Check whether a choice exists
'@details
'Searches the list name column for the specified choice.
'@param choiceName String. Choice name to search for.
'@return Boolean. True when the choice is found.
Public Property Get ChoiceExists(ByVal choiceName As String) As Boolean: End Property

'@jump:sort
'@sub-title Sort choices in the worksheet
'@details
'Sorts by list name ascending, then by ordering column ascending.
Public Sub Sort(): End Sub

'@jump:add-choice
'@sub-title Add a new choice with its categories
'@details
'Appends a new choice entry with long labels and optional short labels.
'Silently exits when the choice already exists or the data is empty.
'@param choiceName String. Name of the choice to add.
'@param longLabelData BetterArray. Long label values for the categories.
'@param shortLabelData Optional Object. BetterArray of short label values. Defaults to Nothing.
Public Sub AddChoice(ByVal choiceName As String, _
                     ByVal longLabelData As BetterArray, _
                     Optional ByVal shortLabelData As Object = Nothing)
End Sub

'@jump:remove-choice
'@sub-title Remove one or multiple choices
'@details
'Removes each choice whose name appears in the separator-delimited
'string. Silently skips choices that do not exist.
'@param choiceNames String. Choice names separated by sep.
'@param sep Optional String. Delimiter between names. Defaults to ";".
Public Sub RemoveChoice(ByVal choiceNames As String, Optional ByVal sep As String = ";")
End Sub

'@jump:translate
'@sub-title Translate the choice worksheet labels
'@details
'Applies the supplied TranslationObject to the label and short label
'columns. Silently exits when headers are invalid.
'@param trads ITranslationObject. Translation provider.
Public Sub Translate(ByVal trads As ITranslationObject): End Sub

'@jump:add-rows
'@sub-title Add rows to the choice table
Public Sub AddRows(): End Sub

'@jump:insert-rows
'@sub-title Insert rows at the selected position
'@details
'Mirrors the selection height inside the choices ListObject.
'@param targetCell Range. Selection anchoring the insertion point.
'@param insertShift Optional Boolean. When True, inserts worksheet rows. Defaults to False.
Public Sub InsertRows(ByVal targetCell As Range, _
                      Optional ByVal insertShift As Boolean = False)
End Sub

'@jump:delete-rows
'@sub-title Delete rows intersecting the selection
'@param targetCell Range. Selection identifying rows to delete.
'@param includeIds Optional Boolean. When True, renumbers the ID column. Defaults to True.
'@param forceShift Optional Boolean. When True, forces worksheet row deletion. Defaults to False.
Public Sub DeleteRows(ByVal targetCell As Range, _
                      Optional ByVal includeIds As Boolean = True, _
                      Optional ByVal forceShift As Boolean = False)
End Sub

'@jump:remove-rows
'@sub-title Remove empty rows from the choice table
Public Sub RemoveRows(): End Sub

'@jump:manage-rows
'@sub-title Add or remove rows via a single toggle
'@param del Optional Boolean. When True, removes rows; when False, adds rows. Defaults to False.
Public Sub ManageRows(Optional ByVal del As Boolean = False): End Sub

'@jump:all-choices
'@prop-title All distinct choice names
'@details
'Scans the list name column and returns every unique choice name.
'@return BetterArray. The distinct choice names.
Public Property Get AllChoices() As BetterArray: End Property


'@section Categories
'===============================================================================
'@description Category retrieval and formatting helpers.

'@jump:categories
'@prop-title Categories of a choice
'@details
'Returns the long label categories for the specified choice. When
'useShortlabels is True, short labels replace long labels where
'available.
'@param choiceName String. The choice to retrieve categories from.
'@param useShortlabels Optional Boolean. When True, uses short labels. Defaults to False.
'@return BetterArray. The category values.
Public Property Get Categories( _
  ByVal choiceName As String, _
  Optional ByVal useShortlabels As Boolean = False) As BetterArray
End Property

'@jump:concatenate-categories
'@prop-title Concatenate all categories of a choice into a string
'@details
'Joins the categories with a separator. Optionally translates each
'category via the supplied TranslationObject.
'@param choiceName String. The choice whose categories to concatenate.
'@param sep Optional String. Separator between categories. Defaults to " | ".
'@param trads Optional Object. A TranslationObject for translation. Defaults to Nothing.
'@return String. The concatenated categories.
Public Property Get ConcatenateCategories( _
        ByVal choiceName As String, _
        Optional ByVal sep As String = " | ", _
        Optional ByVal trads As Object = Nothing) As String
End Property


'@section DataExchange
'===============================================================================
'@description Import, export, and header initialisation operations.

'@jump:add-headers
'@sub-title Write the default headers to the choice worksheet
Public Sub AddHeaders(): End Sub

'@jump:export
'@sub-title Export the choice object to a workbook
'@param toWkb Workbook. Destination workbook.
'@param Hide Optional Long. Worksheet visibility. Defaults to xlSheetHidden.
Public Sub Export(ByVal toWkb As Workbook, Optional ByVal Hide As Long = xlSheetHidden)
End Sub

'@jump:import
'@sub-title Import choices from a worksheet
'@param fromWksh Worksheet. Source worksheet.
'@param fromStartRow Long. Start row in the source.
'@param fromStartcol Long. Start column in the source.
'@param clearSheet Optional Boolean. When True, clears the sheet before import. Defaults to False.
Public Sub Import(ByVal fromWksh As Worksheet, _
                  ByVal fromStartRow As Long, _
                  ByVal fromStartcol As Long, _
                  Optional ByVal clearSheet As Boolean = False)
End Sub

'@jump:import-choice
'@sub-title Import choices from another ILLChoices object
'@param choiObj ILLChoices. Source choice object.
Public Sub ImportChoice(ByVal choiObj As ILLChoices): End Sub


'@section Checkings
'===============================================================================
'@description Diagnostic logging access.

'@jump:has-checkings
'@prop-title Whether the choice object has logged diagnostic messages
'@details
'Returns True when at least one diagnostic entry has been recorded.
'@return Boolean. True when diagnostic entries exist.
Public Property Get HasCheckings() As Boolean: End Property

'@jump:checking-values
'@prop-title Retrieve the diagnostic log entries
'@details
'Returns the internal IChecking object containing all logged messages.
'Returns Nothing when no entries have been recorded.
'@return Object. An IChecking instance, or Nothing.
Public Property Get CheckingValues() As Object: End Property
