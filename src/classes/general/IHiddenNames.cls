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

'@section Enumerations
'===============================================================================
'@sub-title Supported data types when persisting hidden name values
'@details The enumeration informs serialization when writing names and coercion when retrieving values.
Public Enum HiddenNameValueType
    HiddenNameTypeVariant = 0
    HiddenNameTypeString = 1
    HiddenNameTypeBoolean = 2
    HiddenNameTypeLong = 3
End Enum


'@section Introspection
'===============================================================================
'@sub-title Target sheet reference
Public Property Get Sheet() As Worksheet
End Property

'@sub-title Determines whether a name already exists on the worksheet
Public Function HasName(ByVal nameId As String) As Boolean
End Function


'@section Hidden name management
'===============================================================================
'@sub-title Ensure a hidden name exists with the provided default
'@details Returns the cached Name definition so callers can retain a reference when needed.
'@param nameId Identifier of the hidden name (sheet-scoped).
'@param initialValue Value written on first creation.
'@param valueType Optional type annotation guiding serialization and coercion.
'@return Name definition after ensuring existence.
Public Function EnsureName(ByVal nameId As String, _
                           ByVal initialValue As Variant, _
                           Optional ByVal valueType As Byte = HiddenNameTypeVariant) As Name
End Function

'@sub-title Retrieve a hidden value without coercion
'@param nameId Identifier of the hidden name.
'@param defaultValue Optional default when the name is absent.
'@return Stored value, or default when provided.
Public Function Value(ByVal nameId As String, _
                       Optional ByVal defaultValue As Variant) As Variant
End Function

'@sub-title Update the stored value using the cached metadata
'@param nameId Identifier of the hidden name to update.
'@param value Variant to persist.
Public Sub SetValue(ByVal nameId As String, ByVal value As Variant)
End Sub

'@sub-title Retrieve the hidden value coerced to string
Public Function ValueAsString(ByVal nameId As String, _
                              Optional ByVal defaultValue As String = vbNullString) As String
End Function

'@sub-title Retrieve the hidden value coerced to boolean
Public Function ValueAsBoolean(ByVal nameId As String, _
                               Optional ByVal defaultValue As Boolean = False) As Boolean
End Function

'@sub-title Retrieve the hidden value coerced to long
Public Function ValueAsLong(ByVal nameId As String, _
                            Optional ByVal defaultValue As Long = 0) As Long
End Function

'@sub-title Remove a hidden name definition from the worksheet
Public Sub RemoveName(ByVal nameId As String)
End Sub


'@section Discovery
'===============================================================================
'@sub-title Enumerate recorded hidden names (optionally filtered by prefix)
'@param prefix Optional prefix (case-insensitive) to filter returned names.
'@return BetterArray of metadata arrays: Array(nameId, HiddenNameValueType, lastUpdated As Date).
Public Function ListNames(Optional ByVal prefix As String = vbNullString) As BetterArray
End Function


'@section Synchronisation
'===============================================================================
'@sub-title Export all tracked hidden names into another sheet
'@param targetSh Destination sheet that receives the hidden names.
'@param prefix Optional prefix (case-insensitive) to restrict exported names.
Public Sub ExportNames(ByVal targetSh As Worksheet, _
                       Optional ByVal prefix As String = vbNullString)
End Sub

'@sub-title Import hidden names from another sheet
'@param sourceSh Sheet providing the hidden names.
'@param prefix Optional prefix (case-insensitive) to restrict imported names.
'@param overwriteExisting When False, existing values are preserved.
Public Sub ImportNames(ByVal sourceSh As Worksheet, _
                       Optional ByVal prefix As String = vbNullString, _
                       Optional ByVal overwriteExisting As Boolean = True)
End Sub
