VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHiddenNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for managing hidden worksheet name/value storage"

'@Interface
'@Folder("General")
'@ModuleDescription("Contract for managing hidden worksheet name/value storage")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class IHiddenNames
'@description
'Defines the public contract for the HiddenNames class, exposing
'hidden worksheet/workbook-level name-value storage with typed
'serialization, metadata caching, discovery, and cross-scope
'import/export operations.


'@section Enumerations
'===============================================================================
'@description Supported data types when persisting hidden name values.

'@sub-title Supported data types when persisting hidden name values
'@details
'The enumeration informs serialization when writing names and coercion
'when retrieving values.
Public Enum HiddenNameValueType
    HiddenNameTypeVariant = 0
    HiddenNameTypeString = 1
    HiddenNameTypeBoolean = 2
    HiddenNameTypeLong = 3
End Enum


'@section Introspection
'===============================================================================
'@description Properties exposing the host and name existence checks.

'@jump:sheet
'@prop-title Target worksheet reference
'@details
'Returns the worksheet bound to this HiddenNames instance.
'@return Worksheet. The host worksheet.
Public Property Get Sheet() As Worksheet
End Property

'@jump:has-name
'@sub-title Check whether a hidden name exists
'@details
'Determines whether the specified name is defined on the host scope.
'@param nameId String. Identifier of the hidden name.
'@return Boolean. True when the name exists.
Public Function HasName(ByVal nameId As String) As Boolean
End Function


'@section Hidden Name Management
'===============================================================================
'@description Methods for creating, reading, updating, and removing hidden names.

'@jump:ensure-name
'@sub-title Ensure a hidden name exists with the provided default
'@details
'Creates the name with the initial value if it does not yet exist.
'Returns the cached Name definition so callers can retain a reference.
'@param nameId String. Identifier of the hidden name (scope-level).
'@param initialValue Variant. Value written on first creation.
'@param valueType Optional Byte. Type annotation guiding serialization. Defaults to HiddenNameTypeVariant.
'@return Name. The Name definition after ensuring existence.
Public Function EnsureName(ByVal nameId As String, _
                           ByVal initialValue As Variant, _
                           Optional ByVal valueType As Byte = HiddenNameTypeVariant) As Name
End Function

'@jump:value
'@sub-title Retrieve a hidden value without coercion
'@details
'Returns the stored value for the specified name. When the name is
'absent and a default is provided, returns the default; otherwise
'raises an error.
'@param nameId String. Identifier of the hidden name.
'@param defaultValue Optional Variant. Fallback when the name is absent.
'@return Variant. The stored value, or the default.
Public Function Value(ByVal nameId As String, _
                       Optional ByVal defaultValue As Variant) As Variant
End Function

'@jump:set-value
'@sub-title Update the stored value using the cached metadata
'@details
'Writes a new value to an existing hidden name. Raises an error when
'the name does not exist.
'@param nameId String. Identifier of the hidden name to update.
'@param value Variant. The new value to persist.
Public Sub SetValue(ByVal nameId As String, ByVal value As Variant)
End Sub

'@jump:value-as-string
'@sub-title Retrieve the hidden value coerced to String
'@param nameId String. Identifier of the hidden name.
'@param defaultValue Optional String. Fallback when absent. Defaults to vbNullString.
'@return String. The coerced value.
Public Function ValueAsString(ByVal nameId As String, _
                              Optional ByVal defaultValue As String = vbNullString) As String
End Function

'@jump:value-as-boolean
'@sub-title Retrieve the hidden value coerced to Boolean
'@param nameId String. Identifier of the hidden name.
'@param defaultValue Optional Boolean. Fallback when absent. Defaults to False.
'@return Boolean. The coerced value.
Public Function ValueAsBoolean(ByVal nameId As String, _
                               Optional ByVal defaultValue As Boolean = False) As Boolean
End Function

'@jump:value-as-long
'@sub-title Retrieve the hidden value coerced to Long
'@param nameId String. Identifier of the hidden name.
'@param defaultValue Optional Long. Fallback when absent. Defaults to 0.
'@return Long. The coerced value.
Public Function ValueAsLong(ByVal nameId As String, _
                            Optional ByVal defaultValue As Long = 0) As Long
End Function

'@jump:remove-name
'@sub-title Remove a hidden name definition
'@details
'Deletes the hidden name from the host scope and removes its metadata
'cache entry.
'@param nameId String. Identifier of the hidden name to remove.
Public Sub RemoveName(ByVal nameId As String)
End Sub


'@section Discovery
'===============================================================================
'@description Enumerate tracked hidden names.

'@jump:list-names
'@sub-title Enumerate recorded hidden names
'@details
'Returns metadata for all tracked hidden names, optionally filtered
'by a case-insensitive prefix.
'@param prefix Optional String. Prefix filter. Defaults to vbNullString (all names).
'@return BetterArray. Array of metadata records: Array(nameId, HiddenNameValueType, lastUpdated).
Public Function ListNames(Optional ByVal prefix As String = vbNullString) As BetterArray
End Function


'@section Synchronisation
'===============================================================================
'@description Import/export operations across worksheets and workbooks.

'@jump:export-names
'@sub-title Export all tracked hidden names to another worksheet
'@param targetSh Worksheet. Destination worksheet.
'@param prefix Optional String. Prefix filter. Defaults to vbNullString.
Public Sub ExportNames(ByVal targetSh As Worksheet, _
                       Optional ByVal prefix As String = vbNullString)
End Sub

'@jump:import-names
'@sub-title Import hidden names from another worksheet
'@param sourceSh Worksheet. Source worksheet providing the names.
'@param prefix Optional String. Prefix filter. Defaults to vbNullString.
'@param overwriteExisting Optional Boolean. When False, existing values are preserved. Defaults to True.
Public Sub ImportNames(ByVal sourceSh As Worksheet, _
                       Optional ByVal prefix As String = vbNullString, _
                       Optional ByVal overwriteExisting As Boolean = True)
End Sub

'@jump:export-names-to-workbook
'@sub-title Export tracked names to another workbook scope
'@param targetWb Workbook. Destination workbook.
'@param prefix Optional String. Prefix filter. Defaults to vbNullString.
Public Sub ExportNamesToWorkbook(ByVal targetWb As Workbook, _
                                 Optional ByVal prefix As String = vbNullString)
End Sub

'@jump:import-names-from-workbook
'@sub-title Import hidden names from another workbook scope
'@param sourceWb Workbook. Source workbook providing the names.
'@param prefix Optional String. Prefix filter. Defaults to vbNullString.
'@param overwriteExisting Optional Boolean. When False, existing values are preserved. Defaults to True.
Public Sub ImportNamesFromWorkbook(ByVal sourceWb As Workbook, _
                                   Optional ByVal prefix As String = vbNullString, _
                                   Optional ByVal overwriteExisting As Boolean = True)
End Sub

'@jump:set-list-object-header
'@sub-title Create a workbook-level hidden name referencing a ListObject header
'@details
'Adds or replaces a workbook-scoped hidden name whose RefersTo formula
'points at the specified ListObject column.
'@param nameId String. Identifier of the workbook name to assign.
'@param listObject ListObject. Target ListObject hosting the header.
'@param headerName String. Caption of the header column to reference.
Public Sub SetListObjectHeader(ByVal nameId As String, _
                               ByVal listObject As ListObject, _
                               ByVal headerName As String)
End Sub
