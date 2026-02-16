VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDropdownLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of dropdowns"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of dropdowns")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class IDropdownLists
'@description
'Defines the public contract for the DropdownLists class, exposing
'worksheet-based dropdown management built on one-column ListObjects.
'Supports adding, removing, updating, sorting, translating, and
'validating dropdown entries, plus hyperlink navigation and diagnostic
'logging.


'@section Elements
'===============================================================================
'@description Properties exposing the dropdown worksheet and list metadata.

'@jump:wksh
'@prop-title Worksheet hosting the dropdown lists
'@details
'Returns the worksheet that contains the one-column ListObjects.
'@return Worksheet. The host worksheet.
Public Property Get Wksh() As Worksheet: End Property

'@jump:exists
'@prop-title Check whether a dropdown exists
'@details
'Searches the internal registry for the specified dropdown name.
'@param listName String. Dropdown name to search for.
'@return Boolean. True when the dropdown is found.
Public Property Get Exists(ByVal listName As String) As Boolean: End Property

'@jump:name
'@prop-title Name of the dropdown object
'@details
'Returns the name of the worksheet that hosts the dropdown lists.
'@return String. The worksheet name.
Public Property Get Name() As String: End Property

'@jump:label-range
'@prop-title Label cell above a dropdown ListObject
'@details
'Returns the Range of the cell directly above the header row of the
'specified dropdown ListObject. Returns Nothing when the dropdown does
'not exist.
'@param listName String. Dropdown name to look up.
'@return Range. The label cell, or Nothing.
Public Property Get LabelRange(ByVal listName As String) As Range: End Property

'@jump:values
'@prop-title Retrieve the values of a dropdown
'@details
'Returns a BetterArray of the entries stored in the specified dropdown.
'Optionally includes the header row.
'@param listName String. Dropdown name to read from.
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@return BetterArray. The dropdown entries.
Public Property Get Values(ByVal listName As String, _
               Optional ByVal includeHeaders As Boolean = False) As BetterArray
End Property

'@jump:length
'@prop-title Total number of dropdowns on the worksheet
'@details
'Returns the count of registered dropdown lists.
'@return Long. The dropdown count.
Public Property Get Length() As Long: End Property

'@jump:all-dropdowns
'@prop-title All registered dropdown names
'@details
'Returns a BetterArray containing every dropdown name registered on
'the worksheet.
'@return BetterArray. The dropdown name list.
Public Property Get AllDropdowns() As BetterArray: End Property


'@section Modify
'===============================================================================
'@description Methods that add, remove, update, sort, and translate dropdowns.

'@jump:add
'@sub-title Add a new dropdown to the worksheet
'@details
'Creates a new one-column ListObject populated with the supplied values.
'Optionally adds a visible label above the ListObject and applies a
'table style. Silently exits when the dropdown already exists.
'@param listValues BetterArray. String values to populate the ListObject.
'@param listName String. Name to assign to the dropdown.
'@param addLabel Optional Boolean. When True, writes a label above the ListObject. Defaults to False.
'@param counterPrefix Optional String. Prefix for the label text. Defaults to vbNullString.
'@param tableStyle Optional String. Excel table style name. Defaults to "TableStyleLight13".
Public Sub Add(ByVal listValues As BetterArray, ByVal listName As String, _
               Optional ByVal addLabel As Boolean = False, _
               Optional ByVal counterPrefix As String = vbNullString, _
               Optional ByVal tableStyle As String = "TableStyleLight13")
End Sub

'@jump:remove
'@sub-title Remove a dropdown from the worksheet
'@details
'Deletes the ListObject, its workbook-level name reference, and the
'internal registry entry. Silently exits when the dropdown does not exist.
'@param listName String. Dropdown name to remove.
Public Sub Remove(ByVal listName As String): End Sub

'@jump:clear-list
'@sub-title Clear entries of a dropdown
'@details
'Removes all data rows from the dropdown ListObject without deleting
'the dropdown itself. Silently exits when the dropdown does not exist.
'@param listName String. Dropdown name to clear.
Public Sub ClearList(ByVal listName As String)
End Sub

'@jump:update
'@sub-title Update a dropdown with new values
'@details
'Replaces or appends values in the specified dropdown. When
'pasteAtBottom is True, appends below existing data and optionally
'removes duplicates.
'@param UpdateData BetterArray. New values to write.
'@param listName String. Dropdown name to update.
'@param pasteAtBottom Optional Boolean. When True, appends below existing rows. Defaults to False.
'@param removeDuplicates Optional Boolean. When True, removes duplicate entries. Defaults to True.
Public Sub Update(ByVal UpdateData As BetterArray, ByVal listName As String, _
                  Optional ByVal pasteAtBottom As Boolean = False, _
                  Optional ByVal removeDuplicates As Boolean = True)
End Sub

'@jump:sort
'@sub-title Sort a dropdown in the worksheet
'@details
'Sorts the dropdown ListObject using the specified sort order.
'Silently exits when the dropdown does not exist.
'@param listName String. Dropdown name to sort.
'@param sortOrder Optional Long. xlSortOrder enumeration value. Defaults to xlAscending.
Public Sub Sort(ByVal listName As String, Optional ByVal sortOrder As Long = xlAscending)
End Sub

'@jump:translate
'@sub-title Translate all dropdowns on the worksheet
'@details
'Applies the supplied TranslationObject to the data body of every
'registered dropdown. Skips translation when the current language
'matches and force is False.
'@param trads Object. A TranslationObject providing translations.
'@param force Optional Boolean. When True, translates even if the language matches. Defaults to False.
'@param exclude Optional String. Substring filter to skip matching dropdowns.
Public Sub Translate(ByVal trads As Object, Optional ByVal force As Boolean = False, Optional ByVal exclude As String)
End Sub


'@section Operations
'===============================================================================
'@description Validation and hyperlink operations using dropdown references.

'@jump:set-validation
'@sub-title Apply dropdown validation to a Range
'@details
'Sets data validation on the target Range using the specified dropdown
'as the validation source. The alertType controls the severity shown
'when the entered value is not in the list.
'@param cellRng Range. Target range to validate.
'@param listName String. Dropdown name to use as the validation source.
'@param alertType Optional String. Alert severity ("info", "warning", "error"). Defaults to "info".
'@param message Optional String. Message shown on validation failure. Defaults to vbNullString.
'@param showError Optional Boolean. When True, displays validation feedback. Defaults to True.
Public Sub SetValidation(ByVal cellRng As Range, ByVal listName As String, _
                         Optional ByVal alertType As String = "info", _
                         Optional ByVal message As String = vbNullString, _
                         Optional ByVal showError As Boolean = True)
End Sub

'@jump:add-return-link
'@sub-title Add a return hyperlink on the dropdown header
'@details
'Adds a hyperlink on the dropdown label cell that navigates back to
'the specified range.
'@param listName String. Dropdown whose label receives the link.
'@param returnRng Range. Target destination of the hyperlink.
Public Sub AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
End Sub

'@jump:add-hyperlink
'@sub-title Add a hyperlink on a cell targeting a dropdown
'@details
'Adds a hyperlink on the specified range that navigates to the label
'cell of the dropdown.
'@param listName String. Dropdown to navigate to.
'@param linkRng Range. Cell that receives the hyperlink.
Public Sub AddHyperLink(ByVal listName As String, ByVal linkRng As Range)
End Sub


'@section Checkings
'===============================================================================
'@description Diagnostic logging access.

'@jump:has-checkings
'@prop-title Whether the dropdown object has logged diagnostic messages
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
